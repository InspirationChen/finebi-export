#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FineBI 分享链接 -> 数据采集器（零第三方依赖）
支持一个仪表板上的**多个组件**，自动发现、自动重建载荷，导出 csv / xlsx。

链路（实测 FineBI 6.0.27 / FineReport 11.5.12；6.1.x 按同一套 v5 路由设计，见 README「版本兼容」）：
  1) GET  /webroot/decision/link/<code>
        -> HTML 里拿到 sessionId、分享令牌 JWT、reportId，以及 reportWidgets（组件清单）
  2) POST /webroot/decision/v5/design/report/share/report/<reportId>/widgets/detail
        body: {"ids":[realWidgetId, ...]}      # 批量，一次拿全部组件的数据模型
  3) POST /webroot/decision/v5/design/report/share/widget/data?engineType=2&reportId=..&entryType=2&link=<JWT>
        header: sessionId
        body  : 由 detail 重建的组件模型（翻页只改 page）
        resp  : data.header 列名 / data.items 行 / data.hasNext
  4) 逐个组件翻页到 hasNext=false，落到同一个 xlsx 的不同 sheet

关键坑（都是实测出来的）：
  * 取数路由少了 /report/share 这一节会报 "Token is empty!"（报鉴权错，实为路由错）
  * detail 重建的载荷里 settings.tableStyle.totalRows 决定"最多取多少行"，detail 默认只给 1000，
    必须置 0（不限），否则静默只拿到 1000 行
  * wId 要用外层 widgetId；detail 返回的 wId 是 realWidgetId，两者不同
  * 6.1 起取数必须带 cid，否则 FineEmptyCIDException: CID is missing（见 make_cid）
  * 6.1 的分组表/交叉表首次取数 page=0，且改用 rowCounts/colCounts 一次取全（见 fetch_table）
  * 筛选条件是「两半」，缺一半服务端不报错、静默返回全量（见 build_payload 的 ★ 筛选段）

用法：
    python3 finebi_export.py                          # 导出全部表格组件 -> 多 sheet
    python3 finebi_export.py --list-widgets           # 只看页面有哪些组件
    python3 finebi_export.py --widget 客户明细         # 只导指定组件（可逗号分隔/多次）
    python3 finebi_export.py --split                  # 每个组件一个文件
    python3 finebi_export.py --code 'nQcq?性别=女'      # 带筛选导出（过滤组件名=值）
    python3 finebi_export.py --template payload_template.json   # 退回单组件 + 固定模板模式
"""

import argparse
import copy
import datetime as _dt
import hashlib
import json
import os
import re
import sys
import time
import uuid
import urllib.error
import urllib.parse
import urllib.request

from xlsx_io import write_workbook, write_csv, safe_sheet_name
from table_norm import (cell_text, build_table, flatten_header, row_values, did_index,
                        detect_kind, from_response, is_cross_response, is_tree_rows,
                        KIND_DETAIL, KIND_GROUP, KIND_CROSSTAB)

# ---------------------------------------------------------------- 默认配置
CONFIG = {
    "base": "http://localhost:37799",   # FineBI 站点根，内网部署时换成对应地址
    "code": "z8Pz",                     # /webroot/decision/link/z8Pz 里的那段
    "out_dir": os.path.join(os.path.dirname(os.path.abspath(__file__)), "out"),
    "timeout": 90,
    "retries": 3,
    "page_delay": 0.05,                 # 翻页间隔，太猛容易被限流
}

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 " \
     "(KHTML, like Gecko) Chrome/158.0.0.0 Safari/537.36"

# tableStyle.totalRows：0 = 不限制（取全量）。若某个版本对 0 的解释变了，
# 可用环境变量覆盖，例如 FINEBI_TOTAL_ROWS=-1，不用改代码。
TOTAL_ROWS = int(os.environ.get("FINEBI_TOTAL_ROWS", "0") or 0)
# 6.1 的分组表/交叉表不按页码翻页，而是「一次要多大给多大」，
# 靠 tableStyle.rowCounts/colCounts 控制。给一个足够大的值就能一次取全，
# 太小会出现 hasNext=True / 交叉表只有可视区域的问题。
BIG_ROWS = int(os.environ.get("FINEBI_BIG_ROWS", "200000") or 200000)
# 明细表仍然是「按页翻」，每页行数由 payload 里的
# settings.tableStyle.detailRowCounts 决定，服务端默认 100。
# 11 万行按默认值要 1100 次往返（实测 286s），放大到 5000 只剩 22 次（实测约 20s）；
# 再往上放大单次响应会到 260MB 量级，性价比反而下降，所以默认取 5000。
# 设 0 = 不干预，沿用服务端默认。
DETAIL_PAGE_ROWS = int(os.environ.get("FINEBI_DETAIL_PAGE_ROWS", "5000") or 0)

# reportWidgets.title 在用户没改过标题时是 BI 的默认占位符，形如
#   ${"alt":"__widget_name_id__","name":"组件名称"}
# 这种值不能当组件名用（否则一堆组件都叫「组件名称」），真实名字在 detail.name 里。
_PLACEHOLDER_TITLE = re.compile(r"__widget_name_id__|^\s*组件名称\s*$")

# detail 缺失但服务端需要的字段，用这些空值补齐（明细表口径）
_PAYLOAD_DEFAULTS = {
    "explainConf": {"fieldCombinationList": [[]]},
    "fields": [],
    "filterValues": [],
    "drillOrder": [],
    "resultFilter": [],
    "uploadedImages": [],
    "injection": None,
    "legendFilter": None,
    "filter": {"filterType": 34, "filterValue": []},
    "allData": True,
    "realData": True,
    "page": 1,
}


def _first(text, rules):
    """按顺序试多条正则，返回第一个命中的分组 1；全都没命中返回 None。"""
    for pat in rules:
        m = re.search(pat, text)
        if m:
            return m.group(1)
    return None


def _all(text, rules, limit=20):
    """按顺序试多条正则，收集所有命中的分组 1（去重、保序）。"""
    for pat in rules:
        got = []
        for m in re.finditer(pat, text):
            v = m.group(1)
            if v not in got:
                got.append(v)
            if len(got) >= limit:
                break
        if got:
            return got
    return []


def _as_int(v, default=None):
    """宽松转 int：容忍首尾空白和 Excel 粘贴常见的引号，转不了返回 default（不抛异常）。"""
    try:
        return int(str(v).strip().strip('"'))
    except Exception:
        return default


def filter_kind(bean):
    """
    按过滤组件 bean 的值结构粗判类型，只用于错误提示，不参与逻辑判断。

    文本类 value 是字符串数组；日期/数值区间类 value 里套的是对象
    （日期是 {year,month,day,...}），两者从结构上一眼能分开。
    """
    cur = (bean or {}).get("value")
    if isinstance(cur, dict) and isinstance(cur.get("value"), dict):
        return "日期或区间类"
    return "文本类"


def shape_filter_value(bean, values):
    """
    按组件的原始 value 结构决定怎么写入筛选值，写不了返回 None。

    文本/下拉类（TreeControlBean、TextControlBean 等）的 value 形如
        {"type": 1, "value": ["女"], "assist": []}
    值就是字符串数组，直接覆盖没问题。

    日期类的 value 是 DateValueWrapperBean，里面再套 DateStaticValueBean
    （{year, month, day, quarter, period, hour, minute, second}）。把字符串数组
    塞进去，服务端反序列化当场失败：
        Cannot deserialize DateStaticValueBean from Array value
    表现是 HttpMessageNotReadableException，且与组件数量有关——一个日期参数会让
    整个仪表板的所有组件全部取数失败（日志里每个组件一条同样的堆栈）。

    所以这类「对象型值」一律返回 None：宁可明确报不支持，也不发出必然 400 的载荷
    ——那种失败会伪装成「接口坏了」，排查成本极高。
    """
    cur = (bean or {}).get("value")
    if cur is None or isinstance(cur, list):
        return {"type": 1, "value": list(values), "assist": []}
    if isinstance(cur, dict):
        inner = cur.get("value")
        # inner 缺失或本身是数组 -> 文本型；是 dict -> 日期/区间型的对象值
        if inner is None or isinstance(inner, list):
            return {"type": 1, "value": list(values), "assist": []}
        return None
    return None


def probe_version(base, timeout=10):
    """
    探测目标 FineBI 版本，返回 (版本号, 构建时间)，取不到返回 (None, None)。
    版本号来自登录页里 BI-Foundation 模块的 minorVersion（如 6.0.27 / 6.1.10）。
    用于日志留痕和升级前后对比，取不到不影响导出。
    """
    url = "%s/webroot/decision/login" % base.rstrip("/")
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": UA}),
                                    timeout=timeout) as r:
            html = r.read().decode("utf-8", "replace")
    except Exception:
        return None, None
    flat = html.replace('\\"', '"')
    ver = _first(flat, [
        r'BI-Foundation_Module_Name"\s*,\s*"version"\s*:\s*"[^"]*"\s*,\s*"minorVersion"\s*:\s*"([^"]+)"',
        r'"nameLocaleKey"\s*:\s*"BI-Foundation_Module_Name"[^}]*?"minorVersion"\s*:\s*"([^"]+)"',
        r'minorVersion"\s*:\s*"(6\.\d+\.\d+)"',
    ])
    build = ""
    if ver:
        m = re.search(r'"minorVersion"\s*:\s*"%s"\s*,\s*"jarTime"\s*:\s*"([^"]*)"'
                      % re.escape(ver), flat)
        build = m.group(1) if m else ""
    return ver, build


class FineBiError(RuntimeError):
    """本工具统一的业务异常。消息直接给运维看，不夹堆栈、不夹内部字段名。"""
    pass


class Widget(object):
    """仪表板上的一个组件。"""

    def __init__(self, widget_id, real_id, title, type_=None):
        """
        widget_id 外层 id（载荷 wId / queryInfo.reportWidgetId）
        real_id   真实 id（detail 的 key / queryInfo.widgetId）；过滤组件没有这个值，传 None
        type_     组件类型（1 分组表 / 2 交叉表 / 4 明细表 / 32 过滤组件），未知为 None
        """
        self.widget_id = widget_id          # 外层 id（载荷里的 wId / queryInfo.reportWidgetId）
        self.real_id = real_id              # 真实 id（detail 的 key / queryInfo.widgetId）
        self.title = (title or "").strip()      # 空标题在解析完成后按序号补
        # 名字是不是「序号兜底」出来的。真实名字要等 widgets/detail 的 name 字段，
        # 所以这里先记下，拿到 detail 后再覆盖。
        self.title_is_auto = not self.title
        self.type = type_
        self.payload = None                 # 重建后的取数载荷
        self.detail = None                  # 组件原始模型（判表类型、认维度/指标要用）
        self.skip_reason = None             # 非 None 表示跳过（如非表格组件）
        self.notes = None                   # 兜底/降级提示（不阻塞导出）
        # 过滤组件（type=32 / 页面里的 widgetBean）：不导出，只用来把筛选值带进载荷
        self.bean = None                    # widgetBean 原文，构造 widgetIdValueMap 要用
        self.is_filter = False
        self.field = ""                     # 过滤组件绑定的字段名，便于日志和匹配

    def __repr__(self):
        """日志里用短 id + 名称，避免刷一屏 32 位 hex。"""
        return "<Widget %s %r>" % (self.widget_id[:8], self.title or self.widget_id[:8])


class FineBiShareClient:
    """针对 FineBI「公共分享链接」的最小取数客户端（支持多组件）。"""

    def __init__(self, base, code, timeout=90, retries=3, payload_file=None):
        """
        base       站点根地址，如 http://10.0.0.5:37799
        code       分享链接末段（可带 ?参数=值，参数当筛选条件用，见 set_filter）
        payload_file 非 None 则走「固定模板 + 单组件」旧模式（兼容老代码，正常不用）
        """
        self.base = base.rstrip("/")
        self.code = code
        self.timeout = timeout
        self.retries = retries
        self.payload_file = payload_file     # 非 None 则走"固定模板 + 单组件"旧路径
        self.session_id = None
        self.jwt = None
        self.report_id = None
        self.report_ids = []                # 页面里出现的全部 reportId（6.1 多资源链接会有多个）
        self.report_name = ""
        self.template = None
        self.widgets = []
        self.last_failed = []               # 上一轮 export 里失败/跳过的组件 [(标题, 原因)]
        # 6.1 起取数请求必须带 cid（浏览器并发控制 ID），6.0 不带也能跑
        self.cid = None
        # 排错用：设了就把原始 HTML / 请求体 / 响应体落盘，便于比对不同版本的接口结构
        self.dump_dir = None
        self._post_seq = 0
        self.version = ""
        # 筛选：{过滤组件 widgetId: {"name": 显示名, "bean": widgetBean, "values": [...]}}
        # 值最终写进载荷的 queryInfo.widgetIdValueMap + filter（见 build_payload）
        self.filters = {}
        self.url_params = {}                # 链接里 ?a=b 形式的参数
        self.unmatched_params = []          # 传了但没有同名过滤组件的参数，用于提示
        self.unsupported_filters = []       # 名字对上了、但值结构不支持文本写入的组件

    # ------------------------------------------------------------ 原始报文留存
    def _dump(self, name, data):
        """排错留档：把原始报文落盘到 dump_dir。没开 dump 就什么都不做，写失败也不影响主流程。"""
        if not self.dump_dir:
            return
        try:
            os.makedirs(self.dump_dir, exist_ok=True)
            with open(os.path.join(self.dump_dir, name), "wb") as f:
                f.write(data if isinstance(data, bytes) else str(data).encode("utf-8"))
        except Exception:
            pass                                # 留档失败不影响主流程

    # ------------------------------------------------------------ 底层请求
    def _open(self, req):
        """
        发请求并读回字节。

        401/403/404 视为「链接失效 / 无权限」，直接抛错不重试——重试只会把一次明确的
        失败拖成三次，还掩盖真实原因；网络抖动类（超时、连接重置）才退避重试。
        """
        last = None
        for attempt in range(1, self.retries + 1):
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    return resp.read()
            except urllib.error.HTTPError as e:
                last = e
                if e.code in (401, 403, 404):
                    raise FineBiError(
                        "HTTP %s：链接无效/已失效，或服务端拒绝了分享访问" % e.code)
            except Exception as e:                      # 网络抖动类
                last = e
            if attempt < self.retries:
                time.sleep(1.5 * attempt)
        raise FineBiError("请求失败：%s" % last)

    def _get(self, url):
        """GET 一个页面（只带 UA）。分享链接允许匿名访问，不需要 cookie/登录态。"""
        return self._open(urllib.request.Request(url, headers={"User-Agent": UA}))

    def _post_json(self, url, body):
        """POST JSON 并解析响应。开了 dump_dir 时把请求体和响应体各存一份，便于跨版本比对。"""
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST", headers={
            "User-Agent": UA,
            "Content-Type": "application/json;charset=UTF-8",
            "sessionId": self.session_id or "",
        })
        self._post_seq += 1
        tag = re.sub(r"[^A-Za-z0-9]+", "_", url.split("?")[0].rsplit("/", 1)[-1])[:32] or "api"
        if self.dump_dir:
            self._dump("%02d_req_%s.json" % (self._post_seq, tag), data)
        raw = self._open(req)
        if self.dump_dir:
            self._dump("%02d_resp_%s.json" % (self._post_seq, tag), raw)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            raise FineBiError("响应不是 JSON：%s" % raw[:200])

    # ------------------------------------------------------------ 会话引导
    def bootstrap(self):
        """访问分享链接，换取匿名会话、令牌与组件清单。"""
        # 链接允许写成 nQcq?性别=女&姓名=张三：参数只用于筛选，不会出现在请求路径里
        # （服务端对 URL 参数完全不处理，筛选是前端行为，所以这里自己解析）。
        code, _, qs = str(self.code).partition("?")
        self.code = code
        if qs:
            self.url_params = {k: list(v) for k, v in
                               urllib.parse.parse_qs(qs, keep_blank_values=True).items()}
        html = self._get("%s/webroot/decision/link/%s" % (self.base, self.code)) \
            .decode("utf-8", "replace")
        self._dump("00_share_page.html", html.encode("utf-8"))
        flat = html.replace('\\"', '"')

        # 下面四项都按「多条候选规则」取，因为不同 FineBI 小版本的页面模板会换写法。
        # 只认一种写法的话，升级后一旦换模板就会报「找不到 sessionId」，很容易被误判成链接失效。
        self.session_id = _first(flat, [
            r'BI\.sessionId\s*[:=]\s*"([^"]+)"',
            r'"sessionId"\s*:\s*"([^"]{8,})"',
        ])
        if not self.session_id:
            raise FineBiError("页面里找不到 sessionId，链接可能失效或需要登录")

        # 优先取 link 参数上下文里的令牌：页面里可能还有别的 eyJ 开头串（埋点/埋码），
        # 直接搜第一个 eyJ 有拿错的风险。
        self.jwt = _first(flat, [
            r'"link"\s*:\s*"(eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+)"',
            r'[?&]link=(eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+)',
            r'(eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+)',
        ])
        if not self.jwt:
            raise FineBiError("页面里找不到分享令牌（link）")

        # reportId 早期版本是 32 位 hex，不能假设 6.1 仍然如此，所以放宽到 16~64 位 ID 串。
        self.report_id = _first(flat, [
            r'reportId"\s*:\s*"([0-9a-f]{32})"',
            r'reportId"\s*:\s*"([0-9a-zA-Z_\-]{16,64})"',
        ])
        if not self.report_id:
            raise FineBiError("页面里找不到 reportId")

        # FineBI 6.1 起公共链接可以「一次分享多个资源」（主题内多个仪表板/组件/分析文档），
        # 这种链接的页面里会有多个 reportId。这里记下全部，由调用方决定是否告警；
        # 本客户端仍只处理第一个（= 当前打开的仪表板）。
        self.report_ids = _all(flat, [
            r'reportId"\s*:\s*"([0-9a-zA-Z_\-]{16,64})"',
        ])

        self.report_name = _first(flat, [r'reportName"\s*:\s*"([^"]*)"']) or ""

        self.widgets = self._parse_widgets(html)

        # 链接上的参数按过滤组件名（或它绑定的字段名）落到具体的过滤组件上
        for name, values in self.url_params.items():
            if not self.set_filter(name, values):
                self.unmatched_params.append(name)

        # 一旦有参数落到对象型值的组件上，整条链路都不可能成功：
        # 载荷里的 value 结构不对，服务端每个组件都会回 400，最终表现成
        # 「14 个组件全部失败」这种看着像接口坏了的假象。这里一次性说清楚。
        if self.unsupported_filters:
            raise FineBiError(
                "链接参数落到了不支持的过滤组件：%s。\n"
                "这类组件的值是对象结构（日期是 year/month/day，区间是起止对象），"
                "按文本传会被服务端拒绝（Cannot deserialize ... from Array value）。\n"
                "处理办法：① 暂时去掉链接里的该参数，导出全量；"
                "② 或运行 python probe_filters.py --base <地址> --code <链接码> "
                "导出该组件的结构，用它适配日期类筛选。"
                % "、".join(self.unsupported_filters))

        if self.payload_file:
            with open(self.payload_file, "r", encoding="utf-8") as f:
                self.template = json.load(f)
        return self

    # ------------------------------------------------------------ 筛选条件
    def filter_widgets(self):
        """仪表板上的过滤组件（文本下拉/日期区间等），不参与导出。"""
        return [w for w in self.widgets if w.is_filter]

    def set_filter(self, name, values, mode="replace"):
        """
        按过滤组件名或它绑定的字段名设置筛选值。

        名字取组件标题（页面左上角显示的那个）优先，其次取字段名，两者都试，
        避免用户分不清「姓名」是组件标题还是「客户姓名」这个字段名。
        返回命中的组件数，0 表示没有同名过滤组件。
        """
        vals = [str(v) for v in (values if isinstance(values, (list, tuple)) else [values])
                if str(v) != ""]
        hit = 0
        for w in self.filter_widgets():
            if name not in (w.title, w.field):
                continue
            # 值结构不支持的（日期、数值区间）先记下来，bootstrap 结束时报错退出。
            # 不能只是「跳过不写」——那样筛选静默不生效，导出的是全量，
            # 而调用方以为筛过了，这类错误比直接失败危险得多。
            if shape_filter_value(w.bean, vals) is None:
                self.unsupported_filters.append(
                    "%s（%s，字段 %s）" % (w.title or w.field, filter_kind(w.bean), w.field or "?"))
                hit += 1
                continue
            cur = self.filters.get(w.widget_id, {}).get("values", [])
            merged = (cur + vals) if mode == "append" else vals
            # 去重保序，重复传参时不会把同一个值算两遍
            seen, uniq = set(), []
            for v in merged:
                if v not in seen:
                    seen.add(v)
                    uniq.append(v)
            self.filters[w.widget_id] = {"name": w.title or w.field,
                                         "bean": w.bean, "values": uniq}
            hit += 1
        return hit

    def filter_summary(self):
        """生效筛选的可读描述，写进表注和日志。"""
        if not self.filters:
            return ""
        parts = []
        for f in self.filters.values():
            vals = f["values"]
            parts.append("%s=%s" % (f["name"], "、".join(vals) if vals else "（全部）"))
        return "；".join(parts)

    # ------------------------------------------------------------ 组件清单
    @staticmethod
    def _brace_block(text, key):
        """按括号配对取出 '"key":{...}' 的完整 JSON 文本，长度不限（跳过字符串内的括号）。"""
        i = text.find('"%s"' % key)
        if i < 0:
            return None
        j = text.find("{", i)
        if j < 0:
            return None
        depth, in_str, esc = 0, False, False
        for k in range(j, len(text)):
            ch = text[k]
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return text[j:k + 1]
        return None

    @staticmethod
    def _deep_get(obj, want, depth=4):
        """在嵌套结构里找第一个 want 键的值（组件对象层级可能变，别假设一定在顶层）。"""
        if depth < 0 or not isinstance(obj, dict):
            return None
        if obj.get(want):
            return obj[want]
        for v in obj.values():
            if isinstance(v, dict):
                got = FineBiShareClient._deep_get(v, want, depth - 1)
                if got:
                    return got
        return None

    @staticmethod
    def _clean_title(raw):
        """标题清洗：兼容纯文本，以及 ${"alt":"__widget_name_id__","name":"客户明细"} 这种占位写法。"""
        t = str(raw or "").strip()
        mn = re.search(r'"name"\s*:\s*"([^"]+)"', t)
        if mn:
            t = mn.group(1)
        t = t.replace("\\", "").strip()
        if not t or t.startswith("${") or "__widget_name_id__" in t:
            return ""
        return t

    @staticmethod
    def _bean_title(bean):
        """
        过滤组件的显示名：name 字段是一段 HTML（<p class="align-left">性别</p>），
        去标签取文本；取不到就退回它绑定的字段名。
        """
        raw = str(bean.get("name") or "")
        t = re.sub(r"<[^>]+>", "", raw).strip()
        if t:
            return t
        return FineBiShareClient._bean_field(bean)

    @staticmethod
    def _bean_field(bean):
        """过滤组件绑定的第一个字段名（「性别」「客户姓名」这类），仅用于日志和匹配。"""
        for d in (bean.get("dimensions") or {}).values():
            nm = str(d.get("name") or "").strip()
            if nm:
                return nm
        return ""

    @classmethod
    def _parse_widgets(cls, html):
        """
        解析 reportWidgets 得到全部组件。两条路：
          1) 结构化（首选）：按括号配对截出 reportWidgets 的 JSON 原文再 json.loads。
             不受字段顺序变化、新增字段影响——6.1 只要在组件对象里多加字段，
             老的「固定 600 字符窗口」正则就会静默漏组件。
          2) 正则兜底：结构不规整（非标准 JSON）时按特征串扫，窗口放宽到 4000 字符。
        """
        out = []
        # 6.1 的分享页把 reportWidgets 整段放在 JS 字符串里，且是多层转义：
        # 原始 HTML 里键写作 \"reportWidgets\"，title 里的引号还要再叠一层。
        # 一层层反转义并试解析，第一个能 json.loads 成功的就是对的（实测第 2 层），
        # 比写死层数抗版本变化。在原始 html 上直接找 '"reportWidgets"' 永远找不到，
        # 会静默退回正则兜底——而兜底要求条目带 realWidgetId，过滤组件恰好没有，
        # 结果就是「过滤组件整批漏掉」（表现为筛选条件无从设置）。
        flat, block, data = html, None, None
        for _ in range(3):
            block = cls._brace_block(flat, "reportWidgets")
            if block:
                try:
                    data = json.loads(block)
                    break
                except Exception:
                    data = None
            flat = flat.replace('\\"', '"')
        if isinstance(data, dict):
            for wid, v in data.items():
                if not isinstance(v, dict):
                    continue
                # 过滤组件：整份定义挂在 widgetBean 里，没有 realWidgetId，
                # 也不能导出数据——它只负责把筛选值带进取数载荷。
                bean = v.get("widgetBean")
                if isinstance(bean, dict):
                    fw = Widget(str(wid), None, cls._bean_title(bean), 32)
                    fw.is_filter = True
                    fw.bean = bean
                    fw.field = cls._bean_field(bean)
                    out.append(fw)
                    continue
                # 只认像组件的条目，避免把布局/占位对象也算进来
                if not any(k in v for k in ("realWidgetId", "type", "title", "bounds")):
                    continue
                real = v.get("realWidgetId") or v.get("widgetId") \
                    or cls._deep_get(v, "realWidgetId") or wid
                out.append(Widget(str(wid), str(real), cls._clean_title(v.get("title")),
                                  _as_int(v.get("type"))))

        if not out:                              # 结构化失败 -> 正则兜底
            flat = html.replace('\\"', '"')
            start = flat.find('"reportWidgets"')
            if start < 0:
                return []
            seg = flat[start:]
            for m in re.finditer(
                    r'"([0-9a-fA-F]{16,64})"\s*:\s*\{(.{0,4000}?)'
                    r'"realWidgetId"\s*:\s*"([0-9a-fA-F]{16,64})"', seg, re.S):
                wid, mid, real = m.group(1), m.group(2), m.group(3)
                title = ""
                mt = re.search(r'"title"\s*:\s*"(.*?)(?="\s*,\s*"showTitle")', mid, re.S)
                if mt:
                    title = cls._clean_title(mt.group(1))
                mtype = re.search(r'"type"\s*:\s*"?(\d+)"?', mid)
                out.append(Widget(wid, real, title, _as_int(mtype.group(1)) if mtype else None))

        # 去重（同一段可能被命中多次）
        seen, uniq = set(), []
        for w in out:
            if w.widget_id not in seen:
                seen.add(w.widget_id)
                uniq.append(w)
        for i, w in enumerate(uniq, 1):
            if not w.title:
                w.title = "组件%d" % i        # 未命名组件：用序号，避免 sheet 名变成一串 id
        return uniq

    # ------------------------------------------------------------ 组件详情
    def fetch_details(self, widgets=None):
        """批量拿组件数据模型，并按规则重建成取数载荷。过滤组件没有 real_id，跳过。"""
        widgets = widgets if widgets is not None else self.widgets
        todo = [w for w in widgets if not w.is_filter and w.real_id]
        if not todo:
            return widgets
        url = ("%s/webroot/decision/v5/design/report/share/report/%s/widgets/detail"
               "?reportId=%s&entryType=2&link=%s"
               % (self.base, self.report_id, self.report_id, self.jwt))
        res = self._post_json(url, {"ids": [w.real_id for w in todo]})
        if not res.get("success"):
            raise FineBiError("取组件详情失败：%s" % res.get("errorMsg"))
        details = res.get("data") or {}

        for w in todo:
            detail = details.get(w.real_id)
            if not detail:
                w.skip_reason = "服务端未返回该组件模型（可能已被删除或无权限）"
                continue
            w.detail = detail
            # 组件真实名字在 detail.name 里；reportWidgets.title 常常只是默认占位符
            # ${"alt":"__widget_name_id__","name":"组件名称"}，那种情况不能用它当文件名
            # （用户看到的组件名是「全部明细」「客户号+名」这类，存在 detail.name）
            real_name = str(detail.get("name") or "").strip()
            if real_name and not _PLACEHOLDER_TITLE.search(real_name) \
                    and (w.title_is_auto or _PLACEHOLDER_TITLE.search(w.title or "")):
                w.title = real_name
                w.title_is_auto = False
            w.payload = self.build_payload(detail, w)
            if w.payload is None:
                continue
        return widgets

    def build_payload(self, detail, widget):
        """
        把 detail（组件数据模型）重建成 widget/data 可接受的取数载荷。
        非表格类组件（图表等）返回的 data 结构不同，这里判定为暂不支持。
        """
        settings = detail.get("settings") or {}
        if "tableStyle" not in settings:
            # 明细表判定：早期版本靠 settings.tableStyle 存在与否。若某版本把表格样式
            # 挪了位置，这里用「有字段、无图表特征」兜底，避免整页组件被误判成图表全跳过
            # （那种情况的表现是「没有取到任何可导出的表格组件」，很像接口坏了）。
            looks_table = ("chart" not in settings and "axis" not in settings
                           and bool(detail.get("fields") or settings.get("fields")))
            if not looks_table:
                widget.skip_reason = "非明细表组件（图表类），数据结构为 series 非行列表，暂不导出"
                return None
            settings.setdefault("tableStyle", {})
            detail["settings"] = settings
            widget.notes = "未找到 settings.tableStyle，按明细表兜底处理"

        p = copy.deepcopy(detail)
        p.pop("bounds", None)
        p["wId"] = widget.widget_id          # 注意：不是 realWidgetId

        tn = p.get("tableName")
        if isinstance(tn, str) and tn:
            p["tableName"] = [tn]

        qi = p.setdefault("queryInfo", {})
        qi.pop("bounds", None)
        qi.update({
            "reportId": self.report_id,
            "reportName": self.report_name,
            "reportWidgetId": widget.widget_id,
            "widgetId": widget.real_id,
            "sessionId": self.session_id,
            "title": widget.title,
        })

        for k, v in _PAYLOAD_DEFAULTS.items():
            p.setdefault(k, copy.deepcopy(v))
        p["page"] = 1

        # ★ 核心：totalRows 是"最多取多少行"，detail 默认给的是预览值（1000）
        p.setdefault("settings", {}).setdefault("tableStyle", {})["totalRows"] = TOTAL_ROWS

        # ★ 筛选：把过滤组件的选中值写进 queryInfo.widgetIdValueMap（值放 bean.value），
        # 再让顶层 filter 用 filterType=56 引用这几个组件。这是 6.1 前端在页面上
        # 点筛选时发出的真实结构，缺任何一半服务端都不认（只改 value 不写 filter，
        # 或只写 filter 不带 map，都会静默返回全量）。
        if self.filters:
            value_map, refs = {}, []
            for wid, f in self.filters.items():
                bean = copy.deepcopy(f["bean"])
                shaped = shape_filter_value(bean, f["values"])
                if shaped is None:
                    # 正常流程在 bootstrap 就拦住了，这里是第二道防线：
                    # 宁可少一个筛选条件，也不要发出会让服务端 400 的载荷。
                    continue
                bean["value"] = shaped
                value_map[wid] = bean
                refs.append({"filterType": 56, "filterValue": {"widgetId": wid}})
            p.setdefault("queryInfo", {})["widgetIdValueMap"] = value_map
            p["filter"] = {"filterType": 34, "filterValue": refs}
        return p

    def resolve_widgets(self, wanted=None):
        """按名称/ID 筛选组件；wanted 为空表示全部。过滤组件不参与导出，始终排除。"""
        pool = [w for w in self.widgets if not w.is_filter]
        if not wanted:
            return pool
        keys = [k.strip().lower() for k in wanted if k and k.strip()]
        picked = []
        for w in pool:
            hay = (w.title or "").lower()
            if any(k in hay or k == w.widget_id.lower() or k == (w.real_id or "").lower()
                   or hay == k for k in keys):
                picked.append(w)
        return picked

    # ------------------------------------------------------------ 取数
    @staticmethod
    def make_cid():
        """
        生成 cid。6.1 起取数请求必须带它，否则报 FineEmptyCIDException: CID is missing。

        格式：<32位hex>#<13位毫秒时间戳>#<校验码>
        校验码 = MD5(前两段以 # 相连) 的末 8 位。规则的出处是 6.1.12 的
        com.finebi.dashboard.api.structure.session.BIBrowserConcurrencyManager#verifyCid
        （反编译得出），已用浏览器真实抓包对经验证。
        """
        raw = "%s#%d" % (uuid.uuid4().hex, int(time.time() * 1000))
        return raw + "#" + hashlib.md5(raw.encode("utf-8")).hexdigest()[-8:]

    # cid 缓存文件：同一台服务端复用同一个 cid
    CID_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cid_cache")

    def _resolve_cid(self):
        """
        取 cid，优先复用缓存。

        为什么要复用：服务端按 <地址, 设备, cid> 三元组记并发，浏览器 license 通常只给
        1~2 个。若每次运行都造新 cid，跑第二次就会撞
        FineBrowserConcurrencyOverflowException: BI concurrency has reached the limit
        并且旧记录要等 session 过期才清理。固定 cid 后无论跑多少次都只占一个名额。

        可用环境变量 FINEBI_CID 强制指定；设 FINEBI_CID=random 则每次新生成。
        """
        env = os.environ.get("FINEBI_CID", "").strip()
        if env == "random":
            return self.make_cid()
        if env:
            return env

        key = self.base.rstrip("/")
        try:
            if os.path.exists(self.CID_CACHE):
                with open(self.CID_CACHE, "r", encoding="utf-8") as f:
                    for line in f:
                        parts = line.rstrip("\n").split("\t")
                        if len(parts) == 2 and parts[0] == key and parts[1]:
                            return parts[1]
        except Exception:
            pass

        cid = self.make_cid()
        try:
            old = ""
            if os.path.exists(self.CID_CACHE):
                with open(self.CID_CACHE, "r", encoding="utf-8") as f:
                    old = f.read()
            keep = [ln for ln in old.splitlines()
                    if ln and not ln.startswith(key + "\t")]
            keep.append("%s\t%s" % (key, cid))
            with open(self.CID_CACHE, "w", encoding="utf-8") as f:
                f.write("\n".join(keep) + "\n")
        except Exception:
            pass
        return cid

    def _data_url(self):
        """
        取数接口的完整 URL。

        三个参数是踩坑换来的，别删：
          showSectionError=true  让服务端返回具体错因（否则只给一句 success=false）
          cid                    6.1 必填的浏览器并发控制 ID，惰性解析（只有真取数才需要）
          templateId=templateHelperId_export  固定值，缺了会走到别的模板分支
        """
        if not self.cid:
            self.cid = self._resolve_cid()
        return ("%s/webroot/decision/v5/design/report/share/widget/data"
                "?engineType=2&reportId=%s&templateId=templateHelperId_export"
                "&entryType=2&link=%s&showSectionError=true&cid=%s"
                % (self.base, self.report_id, self.jwt,
                   urllib.parse.quote(self.cid, safe="")))

    def fetch_page(self, page, payload=None):
        """
        取一页，返回 data 节点（含 header / items / hasNext / row）。

        page 的语义按表型不同，别乱传：明细表从 1 开始逐页翻；
        分组表/交叉表首次必须是 0（传 1 会被当成「下一页」，服务端回
        cache Page_State expired——这个报错文案是误导，实际是页码不对）。
        """
        body = dict(payload if payload is not None else self.template)
        body["page"] = page
        res = self._post_json(self._data_url(), body)
        if not res.get("success"):
            raise FineBiError("取数失败 page=%s: %s %s"
                              % (page, res.get("errorMsg"), res.get("detailErrorMsg") or ""))
        data = res.get("data") or {}
        if not isinstance(data, dict):
            raise FineBiError("返回结构异常 page=%s: %s" % (page, str(data)[:200]))
        return data

    def fetch_all(self, max_pages=0, log=print, verbose=False, payload=None, label=""):
        """
        逐页取数，返回 (原始表头, 已对齐的数据行, 服务端总行数)。

        行对齐交给 table_norm.row_values()：能按 cell 的 dId 落位，
        所以交叉表/分组表里「合并区只返回首格」的行不会错列。
        """
        header, cols, did_map, rows, page = None, None, None, [], 1
        total = None
        t0 = time.time()
        while True:
            data = self.fetch_page(page, payload=payload)
            if header is None:
                header = data.get("header") or []
                cols = flatten_header(header)[0]
                did_map = did_index(cols)
                total = data.get("row")
            items = data.get("items") or []
            for item in items:
                rows.append(row_values(item, cols, did_map))
            if verbose or page == 1 or page % 100 == 0:
                log("  %s第 %-5d 页 +%d 行，累计 %d / %s"
                    % (label, page, len(items), len(rows), total if total is not None else "?"))
            if not data.get("hasNext") or not items:
                break
            if max_pages and page >= max_pages:
                log("  已达 --max-pages=%d，提前结束" % max_pages)
                break
            page += 1
            time.sleep(CONFIG["page_delay"])
        log("  %s取数完成：%d 行 × %d 列，耗时 %.1fs"
            % (label, len(rows), len(cols or []), time.time() - t0))
        return header or [], rows, total

    # ------------------------------------------------------------ 高层：整页导出
    def fetch_table(self, widget, kind=None, group_mode="merge", max_pages=0,
                    log=print, verbose=False):
        """
        取一个组件的完整数据并落成 Table，自动适配 6.0 / 6.1 的结构差异：

        - 明细表：逐页翻页（page 从 1 开始），行按 dId 对齐；
        - 分组表 / 交叉表：6.1 不再按页码翻页，改成 page=0 + 放大
          tableStyle.rowCounts/colCounts 一次取全，响应是维度树或 top/left 视窗结构，
          交给 table_norm.from_response 解析。

        6.0 上分组表/交叉表走的是平铺结构，from_response 会返回 None，
        这时自动退回原来的平铺解析，行为与旧版一致。
        """
        k, _why = detect_kind(widget.detail or {}, None)
        if kind:
            k = kind
        body = copy.deepcopy(widget.payload)
        body["allData"] = True
        body["realData"] = True

        # 明细表：放大每页行数。detailRowCounts 是 6.x 的字段，旧版本若不认
        # 这个字段只是每页仍按默认行数返回，翻页靠 hasNext 驱动，不影响正确性。
        if k == KIND_DETAIL and DETAIL_PAGE_ROWS:
            ts = body.setdefault("settings", {}).setdefault("tableStyle", {})
            cur = ts.get("detailRowCounts")
            if not isinstance(cur, int) or 0 < cur < DETAIL_PAGE_ROWS:
                ts["detailRowCounts"] = DETAIL_PAGE_ROWS

        if k in (KIND_GROUP, KIND_CROSSTAB):
            ts = body.setdefault("settings", {}).setdefault("tableStyle", {})
            ts["rowCounts"] = BIG_ROWS
            ts["colCounts"] = BIG_ROWS
            t0 = time.time()
            try:
                data = self.fetch_page(0, payload=body)
            except FineBiError as e:
                # 6.0 上分组表/交叉表不接受 page=0，也可能不接受放大的 rowCounts：
                # 这种情况退回原来的逐页取数，不能因为一次尝试失败就丢掉整个组件。
                log("  %s该版本不接受单次取全（%s），改按老方式翻页" % ("    ", str(e)[:60]))
                return self._flat_table(widget, body, kind, group_mode, max_pages, log, verbose)
            t = from_response(widget.title, widget.detail or {}, data, group_mode, kind)
            if t is not None:
                # 是否还有更多数据：分组表看 hasNext，交叉表看行列两个方向
                t.has_next = bool(data.get("hasNext") or data.get("hasXNext")
                                  or data.get("hasYNext"))
                if t.has_next:
                    t.notes.append("服务端标记还有更多数据（hasNext），本次没取全；"
                                   "调大环境变量 FINEBI_BIG_ROWS（当前 %d）后重试" % BIG_ROWS)
                if widget.notes:
                    t.notes.append(widget.notes)
                log("  %s取数完成：%d 行 × %d 列（%s，单次取全）耗时 %.1fs"
                    % ("    ", len(t.rows), len(t.cols), t.kind, time.time() - t0))
                return t
            log("  %s响应是平铺结构（6.0 风格），按老方式翻页取" % "    ")

        return self._flat_table(widget, body, kind, group_mode, max_pages, log, verbose)

    def _flat_table(self, widget, body, kind, group_mode, max_pages, log, verbose):
        """平铺结构（6.0 全表型 / 6.1 明细表）：逐页取 + 按 dId 对齐。"""
        try:
            header, rows, total = self.fetch_all(max_pages, log, verbose, payload=body, label="    ")
        except FineBiError as e:
            # 6.1 上把分组/交叉类组件误判成明细表时会在这炸：page=1 被当作「下一页」，
            # 找不到游标就报 cache Page_State expired（文案误导，实际是起始页码不对）。
            # 补一次 page=0 的单次取全：图表类组件多半还是取不出行列表，但至少
            # 「日期组件/KPI 汇总」这类边缘表型有机会被救回来。
            if "Page_State expired" not in str(e):
                raise
            log("  %s明细路径报 %s，改试 6.1 起始页 page=0" % ("    ", str(e)[:52]))
            ts = body.setdefault("settings", {}).setdefault("tableStyle", {})
            ts["rowCounts"] = BIG_ROWS
            ts["colCounts"] = BIG_ROWS
            data = self.fetch_page(0, payload=body)
            t = from_response(widget.title, widget.detail or {}, data, group_mode, kind)
            if t is None:
                raise FineBiError(
                    "page=0 与 page=1 都取不出表格数据，该组件多半是图表类"
                    "（响应无行列表）；可在配置「表样式」列强制指定后再试")
            t.has_next = bool(data.get("hasNext") or data.get("hasXNext")
                              or data.get("hasYNext"))
            if widget.notes:
                t.notes.append(widget.notes)
            log("  %spage=0 取数成功：%d 行 × %d 列（%s）"
                % ("    ", len(t.rows), len(t.cols), t.kind))
            return t
        return build_table(widget.title, widget.detail or {}, header, rows, total,
                           kind_override=kind, group_mode=group_mode)

    def export(self, wanted=None, max_pages=0, log=print, verbose=False,
               kind_override=None, group_mode="merge"):
        """
        导出指定组件（默认全部表格组件），返回 [(组件, Table)]。
        Table 里已经带好表类型、多级表头、分组列合并方式与告警。
        单个组件失败不影响其他组件，失败原因记在 widget.skip_reason。
        """
        targets = self.resolve_widgets(wanted)
        if not targets:
            raise FineBiError("页面里没有匹配的组件（共 %d 个：%s）"
                              % (len(self.widgets), "、".join(w.title for w in self.widgets) or "无"))

        if self.template is not None:       # 固定模板模式：忽略组件模型，只导一个
            header, rows, total = self.fetch_all(max_pages, log, verbose, payload=self.template)
            t = build_table(targets[0].title, {}, header, rows, total,
                            kind_override=kind_override or KIND_DETAIL, group_mode=group_mode)
            return [(targets[0], t)]

        self.fetch_details(targets)
        results = []
        # 失败/跳过的组件单独留一份：调用方要据此判断「14 个组件只导出了 2 个」，
        # 只看 results 会以为一切正常（[失败] 行是 debug 级，默认看不见）。
        self.last_failed = []
        for i, w in enumerate(targets, 1):
            if not w.payload:
                log("  [跳过] %s：%s" % (w.title, w.skip_reason))
                self.last_failed.append((w.title, w.skip_reason or "无法重建取数载荷"))
                continue
            log("  [%d/%d] %s" % (i, len(targets), w.title))
            try:
                t = self.fetch_table(w, kind=kind_override, group_mode=group_mode,
                                     max_pages=max_pages, log=log, verbose=verbose)
                for n in t.notes:
                    log("    · %s" % n)
                if t.filter_summary:
                    log("    · 组件自带筛选：%s" % "、".join(t.filter_summary))
                results.append((w, t))
            except Exception as e:          # 单组件失败不拖垮整页
                w.skip_reason = str(e)
                log("  [失败] %s：%s" % (w.title, e))
                if "field missing" in str(e).lower():
                    # FineBI 导出前置校验（WidgetFieldMissingCheckStrategy → FineFieldAbsent-
                    # WidgetCheckException）抛的错：组件引用的某个维度字段在原始字段列表、
                    # 计算字段、组件参数里都找不到。这是 FineBI 自家导出链路的校验——
                    # 浏览器里对该组件点「导出」也会报同样的错，不是脚本私造的载荷问题。
                    log("    · 提示：组件引用了失效字段（计算字段被删/字段无权限/跨表关联字段失效），"
                        "浏览器里对该组件点导出也会报同样错误；"
                        "用 probe_table.py --out 存下组件模型可以定位具体字段")
                self.last_failed.append((w.title, str(e)))
        return results


# ---------------------------------------------------------------- 主流程
def main():
    """命令行入口：引导会话 → 选组件 → 取数 → 写 xlsx。返回值即进程退出码。"""
    ap = argparse.ArgumentParser(description="FineBI 分享链接取数并导出（支持多组件，csv/xlsx）")
    ap.add_argument("--base", default=CONFIG["base"], help="FineBI 站点根地址")
    ap.add_argument("--code", default=CONFIG["code"], help="分享链接 code（link/<code>）")
    ap.add_argument("--out", default=None, help="输出路径（csv/xlsx）或目录")
    ap.add_argument("--widget", action="append", help="只导指定组件（名称/ID，可多次或用逗号分隔）")
    ap.add_argument("--list-widgets", action="store_true", help="只列出页面上的组件")
    ap.add_argument("--split", action="store_true", help="每个组件单独一个文件")
    ap.add_argument("--template", default=None,
                    help="固定载荷模板 JSON（旧模式，忽略组件模型）")
    ap.add_argument("--max-pages", type=int, default=0, help="每个组件最多抓几页（0 = 全量）")
    ap.add_argument("--kind", choices=["auto", "detail", "group", "crosstab"], default="auto",
                    help="强制表类型（默认 auto 自动判定；判错时用它纠正）")
    ap.add_argument("--format", choices=["csv", "xlsx"],
                    default=(os.environ.get("FINEBI_OUTPUT_FORMAT", "csv").strip().lower()
                             if os.environ.get("FINEBI_OUTPUT_FORMAT", "").strip().lower()
                             in ("xlsx", "excel", "xls") else "csv"),
                    help="输出格式：csv（默认，无百万行上限）/ xlsx（保留合并单元格与多级表头）")
    ap.add_argument("--group-cols", choices=["merge", "fill", "raw"], default=None,
                    help="分组列处理：merge 合并 / fill 向下填充 / raw 原样；"
                         "默认 csv 时自动 fill、xlsx 时自动 merge")
    ap.add_argument("--verbose", action="store_true", help="逐页打印进度")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    log = (lambda *a, **k: None) if args.quiet else print

    wanted = None
    if args.widget:
        wanted = [x for one in args.widget for x in one.split(",") if x.strip()]

    log("→ 引导会话：%s/webroot/decision/link/%s" % (args.base, args.code))
    client = FineBiShareClient(args.base, args.code,
                               timeout=CONFIG["timeout"],
                               retries=CONFIG["retries"],
                               payload_file=args.template).bootstrap()
    log("  sessionId=%s  reportId=%s  组件数=%d"
        % (client.session_id, client.report_id, len(client.widgets)))

    if args.list_widgets:
        # 组件真实名字在 widgets/detail 里，必须拉一次才能显示（也顺便标出哪些导不出）
        try:
            client.fetch_details()
        except FineBiError as e:
            log("  （取组件详情失败，名称可能不准：%s）" % e)
        print("\n页面：%s（%d 个组件）" % (client.report_name or client.report_id,
                                     len(client.widgets)))
        print("%-3s %-22s %-8s %-8s %-34s" % ("#", "组件名", "类型", "状态", "widgetId"))
        for i, w in enumerate(client.widgets, 1):
            state = "可导出" if w.payload else "跳过"
            kind = "-"
            if w.payload:
                k, _why = detect_kind(w.detail, None)
                kind = k or "-"
            print("%-3d %-22s %-8s %-8s %-34s"
                  % (i, w.title[:20], kind, state, w.widget_id))
            if w.skip_reason:
                print("      └─ %s" % w.skip_reason)
        return 0

    out = args.out or CONFIG["out_dir"]
    ext = ".csv" if args.format == "csv" else ".xlsx"
    as_dir = out.endswith(os.sep) or os.path.isdir(out) \
        or not out.lower().endswith((".xlsx", ".csv"))
    if as_dir:
        os.makedirs(out, exist_ok=True)
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out = os.path.join(out, "finebi_%s%s" % (stamp, ext))

    kind_map = {"detail": KIND_DETAIL, "group": KIND_GROUP, "crosstab": KIND_CROSSTAB}
    kind_override = kind_map.get(args.kind)

    # csv 没有合并单元格：分组列「合并」自动按「填充」处理（用户显式传的值优先）
    if args.group_cols:
        group_mode = args.group_cols
    else:
        group_mode = "fill" if args.format == "csv" else "merge"
        if group_mode == "fill":
            log("  csv 输出：分组列自动按「填充」处理（csv 无合并单元格）")

    log("→ 拉取数据")
    results = client.export(wanted, max_pages=args.max_pages, log=log, verbose=args.verbose,
                            kind_override=kind_override, group_mode=group_mode)
    if not results:
        print("没有取到任何可导出的表格组件。", file=sys.stderr)
        for w in client.widgets:
            print("  - %s：%s" % (w.title, w.skip_reason or "未知原因"), file=sys.stderr)
        return 2

    used = set()
    if args.split:
        base_dir = os.path.dirname(os.path.abspath(out)) or "."
        stem = os.path.splitext(os.path.basename(out))[0]
        paths = []
        for w, t in results:
            sheet = safe_sheet_name(w.title, used)
            p = os.path.join(base_dir, "%s_%s%s" % (stem, sheet, ext))
            if args.format == "csv":
                write_csv(p, t.header, t.rows)
            else:
                write_workbook(p, [t.sheet(sheet)])
            paths.append(p)
            log("→ 已写出：%s（%s，%d 行 × %d 列）"
                % (p, t.kind, len(t.rows), len(t.cols)))
        for p in paths:
            print(p)
    elif args.format == "csv" and len(results) > 1:
        # csv 没有多 sheet：多组件退化为每组件一个文件（与 collector 行为一致）
        log("  csv 无法合成多 sheet，多组件按每组件一个文件输出（要合一个文件请 --format xlsx）")
        base_dir = os.path.dirname(os.path.abspath(out)) or "."
        stem = os.path.splitext(os.path.basename(out))[0]
        paths = []
        for w, t in results:
            sheet = safe_sheet_name(w.title, used)
            p = os.path.join(base_dir, "%s_%s%s" % (stem, sheet, ext))
            write_csv(p, t.header, t.rows)
            log("→ 已写出：%s（%s，%d 行 × %d 列）" % (p, t.kind, len(t.rows), len(t.cols)))
            paths.append(p)
        for p in paths:
            print(p)
    else:
        if args.format == "csv":
            write_csv(out, results[0][1].header, results[0][1].rows)
        else:
            write_workbook(out, [t.sheet(w.title) for w, t in results])
        log("→ 已写出：%s（%d 个组件）" % (out, len(results)))
        print(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
