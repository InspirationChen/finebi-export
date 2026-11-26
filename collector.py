#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FineBI 采集调度器：读配置 Excel -> 按 cron 定时取数 -> 落 Excel

配置 Excel 一行一个任务，列名（中英文都认，见 ALIASES）：
    任务名 | 启用 | 分享链接 | 组件 | 输出方式 | 输出格式 | 输出目录 | 文件名 | cron表达式 | 备注
                                            可选列：payload文件 | 表样式 | 分组列

  组件      留空 = 页面全部表格组件；否则填组件名或 ID，多个用逗号分隔
  输出方式  留空或「多文件」= 每个组件一个文件（默认）；「单文件」= 合成一个文件
  输出格式  留空或「csv」= csv（默认，无百万行上限）；「xlsx」= Excel 格式（保留合并单元格/多级表头）
  表样式    留空或「自动」= 自动判定；也可强制 明细表/分组表/交叉表（判错时纠正用）
  分组列    留空或「合并」= 合并单元格；「填充」= 向下填充；「原样」= 不动（只影响呈现）
            注意：csv 没有合并单元格，输出格式为 csv 时「合并」会自动按「填充」处理
  payload文件 留空 = 自动从页面重建组件载荷，支持多组件；填了就退回固定模板单组件模式

运行：
    python3 collector.py --config config.xlsx           # 常驻，每分钟判一次 cron
    python3 collector.py --config config.xlsx --list    # 只看解析结果和下次执行时间
    python3 collector.py --config config.xlsx --once    # 忽略 cron，立刻把所有启用任务跑一遍
    python3 collector.py --config config.xlsx --dry-run --once   # 只打印不取数

设计要点：
  * 配置每小时强制重载一次；文件 mtime 变化时立刻重载，改完不用重启
  * 每分钟对齐判定 cron，用「任务名@年月日时分」做幂等标记，重启不会重复跑
  * 任务串行执行（FineBI 侧并发取数容易被限流）
"""

import argparse
import datetime as _dt
import json
import logging
import logging.handlers
import os
import re
import signal
import sys
import time
import traceback
from urllib.parse import urlparse

from cron_expr import CronExpr, CronError
from finebi_export import FineBiError, FineBiShareClient
from xlsx_io import read_table, write_csv, write_workbook, safe_filename, safe_sheet_name

HERE = os.path.dirname(os.path.abspath(__file__))

# 输出按「天」归档：输出目录下再建一层 yyyyMMdd 子文件夹，一次运行的所有文件都在里面。
# 这层是每次运行现建的，历史产出互不覆盖；不想要这层就 set FINEBI_DATE_FOLDER=0。
DATE_FOLDER = os.environ.get("FINEBI_DATE_FOLDER", "1") != "0"

# 默认输出格式：csv（xlsx 单 sheet 上限 1,048,576 行，明细一超就写挂；csv 无上限）。
# 环境变量 FINEBI_OUTPUT_FORMAT=xlsx 可全局改回 Excel 输出；配置表「输出格式」列优先级更高。
DEFAULT_FORMAT = (os.environ.get("FINEBI_OUTPUT_FORMAT", "csv") or "csv").strip().lower()
if DEFAULT_FORMAT in ("excel", "xls", "xlsx"):
    DEFAULT_FORMAT = "xlsx"
else:
    DEFAULT_FORMAT = "csv"

ALIASES = {
    "name":     ("任务名", "名称", "name", "job", "任务"),
    "enabled":  ("启用", "是否启用", "enabled", "on", "开关"),
    "url":      ("分享链接", "链接", "url", "link", "地址"),
    "payload":  ("payload文件", "payload", "payload路径", "模型文件", "模型"),
    "widgets":  ("组件", "组件名", "导出组件", "组件列表", "widget", "widgets"),
    "mode":     ("输出方式", "模式", "分表方式", "mode", "split"),
    "outdir":   ("输出目录", "目录", "outdir", "output", "输出路径"),
    "filename": ("文件名", "输出文件名", "filename", "file", "目标文件名"),
    "cron":     ("cron", "cron表达式", "表达式", "schedule", "调度"),
    "remark":   ("备注", "说明", "remark", "note", "描述"),
    "kind":     ("表样式", "表类型", "kind", "tabletype", "类型"),
    "groups":   ("分组列", "分组列处理", "groupcols", "group", "合并方式"),
    "format":   ("输出格式", "格式", "文件格式", "format", "filetype"),
}

TRUTHY = {"是", "y", "yes", "true", "1", "启用", "开", "on"}
FALSY = {"否", "n", "no", "false", "0", "停用", "关", "off"}

# Windows 的「驱动器相对路径」：D:out（不带斜杠）指的是「D 盘当前目录」，
# 语义几乎总是错，而且会拼出 D:\<配置目录>\D:out 这种非法路径，直接报错更省事。
_DRIVE_REL = re.compile(r"^[A-Za-z]:(?![\\/])")

log = logging.getLogger("collector")


def resolve_path(value, base_dir, what="路径"):
    """
    把配置里的路径统一解析成绝对路径。

    处理四件事：
      1. 去掉首尾空白，剥掉 Excel 里常带的包裹引号（"D:\\out" 这种）；
      2. 展开 ~（家目录）和 %USERPROFILE% / $HOME 之类环境变量；
      3. 相对路径以**配置文件所在目录**为基准，而不是进程当前目录 ——
         这样计划任务、Windows 服务、从别的盘启动的调度器都落到同一处；
      4. 归一化（去掉 ./ ..），避免同一个目录被写成两种形式。
    """
    v = str(value or "").strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        v = v[1:-1].strip()
    if not v:
        return ""
    v = os.path.expandvars(os.path.expanduser(v))
    if os.name == "nt" and _DRIVE_REL.match(v):
        raise ValueError(
            "%s 写法有歧义：%s —— 单独一个 %s 表示「%s盘的当前目录」，不是绝对路径。"
            "请补全成 %s\\... 或 %s:/..."
            % (what, v, v[:2], v[:1], v, v[:1]))
    if not os.path.isabs(v):
        # 在非 Windows 上遇到 D:\xxx 这种，只能当相对路径，但几乎肯定是写错了，提醒一句
        if os.name != "nt" and re.match(r"^[A-Za-z]:[\\/]", v):
            log.warning("%s 看起来是 Windows 路径（%s），但当前系统不是 Windows，"
                        "已按相对路径处理；请确认是否写错", what, v)
        v = os.path.join(base_dir, v)
    return os.path.normpath(v)


# ==================================================================== 配置
class Job:
    """配置表里的一行任务，含解析后的 cron / base / code 与校验错误（error 非 None 即跳过）。"""

    def __init__(self, row_no, name, url, payload, outdir, filename, cron, remark, enabled,
                 widgets="", mode="", kind="", groups="", fmt=""):
        """参数与配置列一一对应；带 raw 后缀的保留用户原文，由同名 property 翻译成内部值。"""
        self.row_no = row_no
        self.name = name
        self.url = url
        self.payload = payload
        self.outdir = outdir
        self.filename = filename
        self.cron_raw = cron
        self.remark = remark
        self.enabled = enabled
        self.widgets_raw = widgets or ""
        self.mode = mode or ""
        self.kind_raw = kind or ""          # 表样式：留空/自动 = 自动判定
        self.groups_raw = groups or ""      # 分组列：留空/合并 = merge
        self.fmt_raw = fmt or ""            # 输出格式：留空 = DEFAULT_FORMAT（默认 csv）
        self.cron = None          # 解析成功后填充
        self.base = None
        self.code = None
        self.error = None         # 校验失败原因

    @property
    def key(self):
        """幂等标记用的键。任务名要求唯一，重名在 load_config 里已被加上行号。"""
        return self.name

    @property
    def widget_list(self):
        """组件筛选：空 = 页面上全部表格组件。"""
        if not self.widgets_raw.strip():
            return None
        return [x.strip() for x in re.split(r"[,，;；|/、]", self.widgets_raw) if x.strip()]

    @property
    def split(self):
        """True = 每个组件单独一个文件（默认）；False = 一个文件多个 sheet。"""
        s = _norm(self.mode)
        if not s:
            return True                                   # 留空默认「每组件一个文件」
        if any(k in s for k in ("单文件", "合并", "一个文件", "single")):
            return False
        return True

    @property
    def kind(self):
        """表类型覆盖：None = 自动判定；否则 明细表/分组表/交叉表。"""
        k = _norm(self.kind_raw)
        if not k or k in ("自动", "auto", "自动判定"):
            return None
        if any(x in k for x in ("明细", "detail")):
            return "明细表"
        if any(x in k for x in ("交叉", "crosstab", "pivot")):
            return "交叉表"
        if any(x in k for x in ("分组", "group")):
            return "分组表"
        return None

    @property
    def group_mode(self):
        """分组列处理：merge 合并（默认）/ fill 填充 / raw 原样。"""
        g = _norm(self.groups_raw)
        if not g:
            return "merge"
        if any(x in g for x in ("填充", "fill", "填满")):
            return "fill"
        if any(x in g for x in ("原样", "不处理", "raw", "none")):
            return "raw"
        return "merge"

    @property
    def fmt(self):
        """输出格式：csv（默认）/ xlsx。配置列 > 环境变量 FINEBI_OUTPUT_FORMAT > 内置默认。"""
        f = _norm(self.fmt_raw)
        if not f:
            return DEFAULT_FORMAT
        if f in ("xlsx", "excel", "xls"):
            return "xlsx"
        return "csv"

    @property
    def ext(self):
        """按输出格式推出的文件扩展名。"""
        return ".csv" if self.fmt == "csv" else ".xlsx"

    def target_path(self, now=None, widget=None):
        """
        算出落盘路径。widget 非 None 表示多文件夹模式（每个组件一个文件）。
        多文件模式默认命名：{组件名}_{yyyyMMdd}.xlsx
        若配置的文件名里没写 {widget}，会自动补上，否则各组件会写到同一个文件互相覆盖。
        扩展名由「输出格式」列决定（csv/xlsx）：文件名里手写的扩展名会被替换，
        比如文件名列写 ...xlsx、输出格式 csv -> 落盘 ...csv，避免两种格式混着写错。
        """
        now = now or _dt.datetime.now()
        if widget is not None:
            # 默认命名：任务名_组件名_yyyymmdd。配置里手写了「文件名」就以配置为准
            #（没写 {widget} 会自动补上，否则各组件会写到同一个文件互相覆盖；
            #  手写的扩展名先剥掉，扩展名统一由「输出格式」决定）
            name = self.filename or "{name}_{widget}_{yyyyMMdd}"
            if "{widget}" not in name:
                name = os.path.splitext(name)[0] + "_{widget}"
        else:
            # 单文件模式没有单个组件名，用 任务名_yyyymmdd（多组件合成一个文件）
            name = self.filename or "{name}_{yyyyMMdd}"

        mapping = {
            "yyyy": "%04d" % now.year, "MM": "%02d" % now.month, "dd": "%02d" % now.day,
            "HH": "%02d" % now.hour, "mm": "%02d" % now.minute, "ss": "%02d" % now.second,
            "yyyyMMdd": now.strftime("%Y%m%d"), "HHmmss": now.strftime("%H%M%S"),
            "yyyyMMdd_HHmmss": now.strftime("%Y%m%d_%H%M%S"),
            "name": self.name, "date": now.strftime("%Y-%m-%d"),
            "widget": safe_filename(widget, fallback="组件") if widget is not None else "",
        }
        fn = name
        for k in sorted(mapping, key=len, reverse=True):
            fn = fn.replace("{" + k + "}", mapping[k])
        # 日期子文件夹：out/20261125/xxx.csv。输出目录里已手写 {yyyyMMdd} 的话尊重原样，
        # 不再重复加一层。目录本身不用在这里建，write_workbook/write_csv 会 makedirs。
        out_dir = self.outdir
        if DATE_FOLDER and "{yyyyMMdd}" not in out_dir:
            out_dir = os.path.join(out_dir, now.strftime("%Y%m%d"))
        return os.path.join(out_dir, safe_filename(os.path.splitext(fn)[0],
                                                   fallback="导出")
                            + self.ext)

    def next_run(self, now=None):
        """下一次触发时间；cron 解析失败（job.error 非 None）时返回 None。"""
        return self.cron.next_after(now) if self.cron else None


def _norm(s):
    """列名归一化：去空白/下划线/连字符/括号并转小写，这样「输出目录」「输出 目录」都能匹配。"""
    return re.sub(r"[\s_\-（）()]", "", (s or "")).lower()


def _pick(header, field):
    """按别名在表头里找列下标。"""
    wanted = {_norm(a) for a in ALIASES[field]}
    for i, h in enumerate(header):
        if _norm(h) in wanted:
            return i
    return -1


def _truthy(v, default=True):
    """把「启用」列的写法翻成布尔；留空按 default（默认启用，少写一列不至于任务不跑）。"""
    s = (v or "").strip().lower()
    if not s:
        return default
    if s in TRUTHY:
        return True
    if s in FALSY:
        return False
    return default


def parse_url(url):
    """
    从分享链接里拆出 base 与 code。

    链接上的 ?参数=值 会原样拼回 code，交给客户端当筛选条件用：
        .../link/nQcq?性别=女  ->  base=http://host:port,  code="nQcq?性别=女"
    这一步不能把 ? 后面丢掉，否则配置里写了筛选也静默失效。
    """
    m = re.match(r"^(https?://[^/]+).*/link/([A-Za-z0-9_-]+)/?(?:\?(.*))?$", url.strip())
    if not m:
        raise ValueError("分享链接格式不对，应形如 http://host:port/webroot/decision/link/xxxx")
    code = m.group(2)
    if m.group(3):
        code += "?" + m.group(3)
    return m.group(1).rstrip("/"), code


def load_config(path):
    """读配置 Excel -> 任务列表（含校验结果）。"""
    header, rows = read_table(path)
    if not header:
        raise ValueError("配置表是空的：%s" % path)

    idx = {f: _pick(header, f) for f in ALIASES}
    missing = [f for f in ("url", "cron") if idx[f] < 0]
    if missing:
        raise ValueError("配置表缺少必需列：%s（表头为 %s）"
                         % (["分享链接" if f == "url" else "cron表达式" for f in missing], header))

    base_dir = os.path.dirname(os.path.abspath(path))
    jobs, seen = [], set()

    for n, row in enumerate(rows, start=2):
        def cell(field):
            i = idx[field]
            return row[i].strip() if 0 <= i < len(row) else ""

        url, cron_raw = cell("url"), cell("cron")
        if not url and not cron_raw:
            continue                                  # 整行空，跳过
        if url.startswith("#"):
            continue                                  # 注释行

        name = cell("name") or "任务%d" % (n - 1)
        if name in seen:
            name = "%s_行%d" % (name, n)
        seen.add(name)

        job = Job(row_no=n, name=name, url=url,
                  payload=cell("payload"),          # 留空 = 自动从 detail 重建组件载荷
                  outdir=cell("outdir") or "out",
                  filename=cell("filename"), cron=cron_raw,
                  remark=cell("remark"),
                  enabled=_truthy(cell("enabled")),
                  widgets=cell("widgets"), mode=cell("mode"),
                  kind=cell("kind"), groups=cell("groups"),
                  fmt=cell("format"))

        errs = []
        try:
            job.base, job.code = parse_url(url)
        except ValueError as e:
            errs.append(str(e))

        if job.payload:
            try:
                job.payload = resolve_path(job.payload, base_dir, "payload文件")
            except ValueError as e:
                errs.append(str(e))
            else:
                if not os.path.exists(job.payload):
                    errs.append("payload 文件不存在：%s" % job.payload)

        try:
            job.outdir = resolve_path(job.outdir, base_dir, "输出目录")
        except ValueError as e:
            errs.append(str(e))

        try:
            job.cron = CronExpr(cron_raw)
        except CronError as e:
            errs.append("cron 解析失败：%s" % e)

        job.error = "；".join(errs) or None
        jobs.append(job)

    return jobs


# ==================================================================== 运行
class State:
    """记录每个任务每分钟是否已执行，避免重启后重复跑。"""

    def __init__(self, path):
        """读状态文件；读不出来就当空的——宁可多跑一次，也不要因为状态文件坏了拒绝调度。"""
        self.path = path
        self.data = {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
        except Exception:
            self.data = {}

    def done(self, job, when):
        """这个任务在 when 这一分钟是否已经成功跑过（常驻模式重启后靠它防重复执行）。"""
        return self.data.get("%s@%s" % (job.name, when.strftime("%Y%m%d%H%M"))) == "ok"

    def mark(self, job, when, status):
        """记下执行结果，顺手丢掉 3 天前的记录——否则状态文件会一直长。"""
        self.data["%s@%s" % (job.name, when.strftime("%Y%m%d%H%M"))] = status
        cutoff = (_dt.datetime.now() - _dt.timedelta(days=3)).strftime("%Y%m%d%H%M")
        self.data = {k: v for k, v in self.data.items() if k.split("@")[-1] >= cutoff}
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=1)
        except Exception as e:
            log.warning("状态文件写入失败：%s", e)


def _write_one(path, table, sheet):
    """按扩展名写单个文件（csv 无合并单元格概念，xlsx 保留多级表头与合并）。"""
    try:
        if path.lower().endswith(".csv"):
            return write_csv(path, table.header, table.rows)
        return write_workbook(path, [table.sheet(sheet)])
    except PermissionError as e:
        raise FineBiError(str(e))


def _write_split(job, results, log):
    """
    多文件写出：每个组件一个文件（csv/xlsx 由任务格式决定）。
    组件重名 / Windows 大小写不敏感 -> 文件名去重，避免互相覆盖。
    返回写出的路径列表。
    """
    paths, taken = [], set()
    for w, t in results:
        p = job.target_path(widget=w.title)
        stem, ext = os.path.splitext(p)
        cand, n = p, 2
        while os.path.normcase(cand) in taken:
            cand = "%s_%d%s" % (stem, n, ext)
            n += 1
        taken.add(os.path.normcase(cand))
        _write_one(cand, t, safe_sheet_name(w.title))
        paths.append(cand)
        log.debug("   -> %s（%s，%d 行 × %d 列）", cand, t.kind, len(t.rows), len(t.cols))
        # 类型判定、分组列处理等提示记进 debug，出问题能回溯；
        # 「行数不符」统一交给 _report_rows 打 WARNING，这里跳过避免重复
        for note in t.notes:
            if "行数不符" not in note:
                log.debug("      · %s", note)
    if not paths:
        raise FineBiError("没有可写出的文件")
    return paths


def run_job(job, dry_run=False, max_pages=0):
    """
    执行一个任务。返回 (总行数, 产出的文件路径列表, 耗时秒)。
    多文件夹模式（默认）下每个组件一个文件：{组件名}_{yyyyMMdd}.csv/.xlsx
    """
    t0 = time.time()
    out = job.target_path()
    if dry_run:
        log.info("[dry-run] %s -> %s", job.name, out)
        return 0, [out], 0.0

    log.info("[开始] %s  (%s)", job.name, job.cron_raw)
    client = FineBiShareClient(job.base, job.code,
                               payload_file=job.payload or None).bootstrap()
    wanted = job.widget_list
    tables = client.resolve_widgets()
    filters = client.filter_widgets()
    log.info("  页面「%s」共 %d 个表格组件%s，本次导出：%s",
             client.report_name or client.report_id, len(tables),
             "、%d 个过滤组件" % len(filters) if filters else "",
             "全部表格组件" if not wanted else "、".join(wanted))
    if filters:
        log.info("  过滤组件：%s", "、".join(
            "%s(%s)" % (w.title, w.field) if w.field and w.field != w.title else w.title
            for w in filters))
    # 筛选是否生效必须打在日志里：不生效时数据看起来「很正常」，只是全量而已
    if client.filters:
        log.info("  生效筛选：%s", client.filter_summary())
    if client.unmatched_params:
        log.warning("  链接里的参数没有同名过滤组件，已忽略：%s",
                    "、".join(client.unmatched_params))

    # csv 没有合并单元格：分组列的「合并」自动按「填充」处理，否则维度列大片空白没法用
    gm = job.group_mode
    if job.fmt == "csv" and gm == "merge":
        gm = "fill"
        log.debug("  csv 输出：分组列「合并」自动改为「填充」（csv 无合并单元格）")
    results = client.export(wanted, max_pages=max_pages, log=lambda m: log.debug(m),
                            kind_override=job.kind, group_mode=gm)
    if not results:
        detail = "；".join("%s：%s" % (w.title, w.skip_reason or "未知") for w in tables)
        raise FineBiError("没有取到任何可导出的表格组件（%s）" % (detail or "页面无可导出组件"))

    total_rows = sum(len(t.rows) for _w, t in results)
    cost = time.time() - t0
    failed = list(getattr(client, "last_failed", []) or [])

    # 组件级失败必须显式说出来。默认只写出成功的那几个文件，日志照样打 [完成]，
    # 很容易被当成「跑完了」——实际是 14 个组件只导出了 2 个。
    if failed:
        log.warning("  ! 有 %d 个组件没导出来（成功 %d 个）：", len(failed), len(results))
        for title, why in failed[:12]:
            log.warning("     × %-22s %s", title, str(why).splitlines()[0][:150])
        if len(failed) > 12:
            log.warning("     …另有 %d 个，用 --verbose 看完整堆栈", len(failed) - 12)

    if job.split:
        paths = _write_split(job, results, log)
        size = sum(os.path.getsize(p) for p in paths) / 1024.0 / 1024.0
        log.info("[完成] %s  %d 个文件  %d 行  %.1f MB  %.1fs",
                 job.name, len(paths), total_rows, size, cost)
        log.info("       类型：%s", "、".join("%s=%s" % (w.title, t.kind) for w, t in results))
        for p in paths:
            log.info("       %s", os.path.basename(p))
        _report_rows(log, results)
        _csv_text_hint(log, results, job)
        return total_rows, paths, cost

    # 单文件模式：csv 没有多 sheet 概念——单组件直接写一个 csv；
    # 多组件退化为每组件一个文件（并提示，避免用户以为数据丢了）
    if job.fmt == "csv" and len(results) > 1:
        log.warning("  单文件模式不支持 csv 合并多组件，已改为每组件一个 csv 文件"
                    "（要合成一个文件请把「输出格式」改回 xlsx）")
        paths = _write_split(job, results, log)
        size = sum(os.path.getsize(p) for p in paths) / 1024.0 / 1024.0
        log.info("[完成] %s  %d 个文件  %d 行  %.1f MB  %.1fs",
                 job.name, len(paths), total_rows, size, cost)
        for p in paths:
            log.info("       %s", os.path.basename(p))
        _report_rows(log, results)
        _csv_text_hint(log, results, job)
        return total_rows, paths, cost

    try:
        if job.fmt == "csv":
            _write_one(out, results[0][1], safe_sheet_name(results[0][0].title))
        else:
            write_workbook(out, [t.sheet(w.title) for w, t in results])
    except PermissionError as e:
        raise FineBiError(str(e))
    size = os.path.getsize(out) / 1024.0 / 1024.0
    log.info("[完成] %s  %d 行  %.1f MB  %.1fs -> %s",
             job.name, total_rows, size, cost, os.path.basename(out))
    _report_rows(log, results)
    _csv_text_hint(log, results, job)
    return total_rows, [out], cost


def _csv_text_hint(log, results, job):
    """csv 输出且表里有身份证/客户号类列时提示一次：Excel 双击打开会截断长数字精度。"""
    if job.fmt != "csv":
        return
    from xlsx_io import FORCE_TEXT_HINTS
    hits = set()
    for _w, t in results:
        for h in t.header:
            if any(k in str(h) for k in FORCE_TEXT_HINTS):
                hits.add(str(h))
    if hits:
        log.warning("  ! csv 含长数字列（%s）：Excel 双击打开会丢精度/变科学计数法，"
                    "请用「数据→自文本导入」查看，或该任务改用 xlsx 输出",
                    "、".join(sorted(hits)[:5]))


def _report_rows(log, results):
    """
    行数自证：用服务端报的总行数比对实际写出的行数，少了就打 WARNING。

    为什么单独拎出来：明细表靠翻页取数，任何「少一页 / hasNext 提前为假」都会
    表现为「导出比页面少几条」，而这种现象不会报错、文件也正常打开，只能靠
    这个比对发现（之前它混在 t.notes 里，级别是 DEBUG，不加 --verbose 根本看不见）。

    判据是「服务端自己报的行数」：
      * 服务端报 4400、实取 4341 -> 翻页/截断，脚本侧问题，这里会告警；
      * 服务端报 4341、页面显示 4400 -> 两边口径不同（权限、筛选、实时数据变动、
        组件级的「合并相同项」），脚本无责，这里不出声，要去 BI 侧查。
    非明细表（分组/交叉）的 expected_total 是明细行数，与分组后的行数本来就不等，
    这条路径不会传入，所以不会误报。
    """
    bad = []
    for w, t in results:
        exp = getattr(t, "expected_total", None)
        if isinstance(exp, int) and exp > 0 and exp != len(t.rows):
            bad.append((w.title, exp, len(t.rows)))
    if not bad:
        return
    for title, exp, got in bad:
        log.warning("  ! 行数不符「%s」：服务端报 %d 行，实际取到 %d 行，差 %d 行",
                    title, exp, got, exp - got)
    log.warning("  以上差异属于取数阶段丢行，请用 --max-pages 0 重跑；仍不符就把"
                "日志连同 FINEBI_BIG_ROWS 当前值一起反馈")


def setup_logging(log_dir, verbose=False):
    """日志同时写控制台和 logs/collector.log（按天切分、留 30 天）。文件不可用就只写控制台。"""
    log.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s %(levelname)-5s %(message)s", "%Y-%m-%d %H:%M:%S")

    sh = logging.StreamHandler(sys.stderr)
    sh.setLevel(logging.DEBUG if verbose else logging.INFO)
    sh.setFormatter(fmt)
    log.addHandler(sh)

    try:
        os.makedirs(log_dir, exist_ok=True)
        fh = logging.handlers.TimedRotatingFileHandler(
            os.path.join(log_dir, "collector.log"), when="midnight",
            backupCount=30, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        log.addHandler(fh)
    except Exception as e:
        log.warning("日志文件不可用，仅输出到控制台：%s", e)


def print_jobs(jobs):
    """--list 的输出：任务清单 + 下次执行时间 + 落盘位置（用于上线前核对）。"""
    log.info("%-16s %-6s %-28s %-22s %s", "任务名", "启用", "cron", "下次运行", "输出")
    for j in jobs:
        flags = "是" if j.enabled else "否"
        if j.error:
            log.info("%-16s %-6s %s  ⚠ %s", j.name, flags, "-", j.error)
            continue
        nxt = j.next_run()
        target = j.target_path(widget="组件名") if j.split else j.target_path()
        try:
            # 输出目录和脚本不在同一个盘时（C: vs D:），Windows 的 relpath 会抛 ValueError
            rel = os.path.relpath(target, HERE)
            if rel.startswith(".."):
                rel = target
        except ValueError:
            rel = target
        log.info("%-16s %-6s %-28s %-22s %s", j.name, flags, j.cron_raw,
                 nxt.strftime("%m-%d %H:%M") if nxt else "-", rel)


def _force_utf8_stdio():
    """
    输出编码兜底，分三种情况：
    1. 真控制台（cmd 窗口）：Python 3.6+ 走 Windows Unicode API（WriteConsoleW），
       中文天然不乱码不报错，不动它——强行改成 UTF-8 反而会在 GBK 控制台打出乱码；
    2. 管道（webadmin 抓日志、> 重定向）：编码由调用方决定（webadmin 传
       PYTHONIOENCODING=utf-8），这里只把编不了的字符降级为 ?，避免罕见字符报错；
    3. 其他异常（极老 Python 没有 reconfigure）一律忽略，不影响主流程。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream.isatty():
                continue
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def main():
    """
    命令行入口，四种模式互斥：
      --once          立刻跑一轮（可配 --job / --json 给外部调度器用）
      --list          只打印任务清单和下次执行时间
      --list-widgets  列出各链接页面上的组件（改「组件」列前先跑）
      默认            常驻，每分钟判一次 cron
    """
    _force_utf8_stdio()
    ap = argparse.ArgumentParser(description="FineBI 配置驱动的定时采集")
    ap.add_argument("--config", default=os.path.join(HERE, "config.xlsx"), help="配置 Excel 路径")
    ap.add_argument("--once", action="store_true", help="忽略 cron，立刻执行所有启用任务")
    ap.add_argument("--list", action="store_true", help="只打印解析结果，不执行")
    ap.add_argument("--list-widgets", action="store_true",
                    help="列出配置里各链接页面上的组件（改「组件」列前先跑这个）")
    ap.add_argument("--dry-run", action="store_true", help="不真正取数")
    ap.add_argument("--verbose", action="store_true")
    ap.add_argument("--run-seconds", type=int, default=0, help="调试用：跑 N 秒后退出")
    ap.add_argument("--max-pages", type=int, default=0, help="调试用：每组件最多抓几页（0=全量）")
    ap.add_argument("--job", default=None,
                    help="只跑指定任务（任务名或名字片段，多个用逗号分隔）；给外部调度器用")
    ap.add_argument("--json", action="store_true",
                    help="向 stdout 输出机器可读的 JSON 结果（日志走 stderr），便于 ElasticJob 等解析")
    args = ap.parse_args()

    cfg_path = os.path.abspath(args.config)
    setup_logging(os.path.join(os.path.dirname(cfg_path), "logs"), args.verbose)

    try:
        jobs = load_config(cfg_path)
    except Exception as e:
        log.error("配置读取失败：%s", e)
        return 2

    if not jobs:
        log.warning("配置里没有有效任务：%s", cfg_path)
        return 2

    valid = [j for j in jobs if not j.error]
    bad = [j for j in jobs if j.error]
    for j in bad:
        log.warning("跳过任务 %s（第 %d 行）：%s", j.name, j.row_no, j.error)

    if args.list_widgets:
        seen_url, rc = set(), 0
        for j in valid:
            if j.url in seen_url:
                continue
            seen_url.add(j.url)
            try:
                c = FineBiShareClient(j.base, j.code).bootstrap()
            except Exception as e:
                log.error("[%s] 打开链接失败：%s", j.name, e)
                rc = 2
                continue
            print("\n【%s】%s  链接 %s" % (j.name, c.report_name or c.report_id, j.url))
            # 真实组件名在 widgets/detail 的 name 字段，必须拉一次
            try:
                c.fetch_details()
            except Exception as e:
                log.warning("  取组件详情失败，名称可能不准：%s", e)
            print("  组件 %d 个：" % len(c.widgets))
            for i, w in enumerate(c.widgets, 1):
                state = "可导出" if w.payload else "跳过"
                print("   %-3d %-24s %-6s widgetId=%s"
                      % (i, w.title[:22], state, w.widget_id))
                if w.skip_reason:
                    print("       └─ %s" % w.skip_reason)
            if j.widget_list:
                hit = c.resolve_widgets(j.widget_list)
                miss = [k for k in j.widget_list
                        if not any(k.strip().lower() in (w.title or "").lower() for w in hit)]
                print("  配置的组件筛选：%s  ->  命中 %d 个%s"
                      % ("、".join(j.widget_list), len(hit),
                         "，未命中：%s" % "、".join(miss) if miss else ""))
        return rc

    if args.list or (args.once and args.dry_run):
        print_jobs(jobs)
        if args.list:
            return 0

    state = State(os.path.join(os.path.dirname(cfg_path), ".collector_state.json"))
    stop = {"flag": False}

    def _signal(signum, _frame):
        log.info("收到信号 %s，准备退出…", signum)
        stop["flag"] = True
    for s in (signal.SIGINT, signal.SIGTERM):
        signal.signal(s, _signal)

    if args.once:
        if args.job:
            # 显式点名任务时不受「启用」列限制：点名 = 想跑它（网页上单跑按钮依赖这个，
            # 否则停用的任务永远跑不了）。没点名才默认只跑启用的。
            pool = valid
        else:
            pool = [j for j in valid if j.enabled]
        targets = pool
        if args.job:
            names = [x.strip() for x in args.job.split(",") if x.strip()]
            targets = [j for j in targets
                       if any(n == j.name or n.lower() in j.name.lower() for n in names)]
            if not targets:
                log.error("没有匹配的任务：%s（可选：%s）", args.job,
                          "、".join(j.name for j in valid) or "无")
                return 2

        log.info("一次性执行 %d 个任务", len(targets))
        fails, report = 0, []
        for job in targets:
            try:
                rows, paths, cost = run_job(job, args.dry_run, args.max_pages)
                report.append({"task": job.name, "ok": True, "rows": rows,
                               "seconds": round(cost, 1), "files": paths})
            except Exception as e:
                fails += 1
                report.append({"task": job.name, "ok": False, "error": str(e)})
                log.error("[失败] %s：%s", job.name, e)
                log.debug("%s", traceback.format_exc(limit=3))
        log.info("一次性执行结束，成功 %d 个，失败 %d 个", len(report) - fails, fails)

        if args.json:
            print(json.dumps({
                "ok": fails == 0, "total": len(report), "failed": fails,
                "finishedAt": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "results": report,
            }, ensure_ascii=False, indent=1))
        return 1 if fails else 0

    # ---------------------------------------------------------- 常驻
    log.info("调度器启动：%s，共 %d 个任务（%d 个可用）",
             cfg_path, len(jobs), len([j for j in valid if j.enabled]))
    for j in valid:
        if j.enabled:
            nxt = j.next_run()
            log.info("  %-16s %-20s 下次 %s", j.name, j.cron_raw,
                     nxt.strftime("%Y-%m-%d %H:%M") if nxt else "永不")

    started = time.time()
    last_reload = time.time()
    cfg_mtime = os.path.getmtime(cfg_path)
    running_key = set()

    while not stop["flag"]:
        now = _dt.datetime.now()
        # 睡到下一分钟第 2 秒，避免边界抖动
        nxt_tick = (now.replace(second=0, microsecond=0) + _dt.timedelta(minutes=1, seconds=2))
        while not stop["flag"] and _dt.datetime.now() < nxt_tick:
            if args.run_seconds and time.time() - started > args.run_seconds:
                stop["flag"] = True
                break
            time.sleep(1)
        if stop["flag"]:
            break

        tick = _dt.datetime.now().replace(second=0, microsecond=0)

        # 重载配置：每小时一次，或文件被改动
        try:
            mtime = os.path.getmtime(cfg_path)
        except OSError:
            mtime = cfg_mtime
        if time.time() - last_reload >= 3600 or mtime != cfg_mtime:
            reason = "文件已变更" if mtime != cfg_mtime else "满 1 小时"
            try:
                jobs = load_config(cfg_path)
                valid = [j for j in jobs if not j.error]
                bad = [j for j in jobs if j.error]
                cfg_mtime = mtime
                last_reload = time.time()
                log.info("[重载] %s，%d 个任务（%d 个跳过）", reason, len(jobs), len(bad))
                for j in bad:
                    log.warning("  跳过 %s：%s", j.name, j.error)
            except Exception as e:
                log.error("[重载] 失败，沿用旧配置：%s", e)
                last_reload = time.time()

        for job in valid:
            if stop["flag"]:
                break
            if not job.enabled or job.error or job.cron is None:
                continue
            if not job.cron.matches(tick):
                continue
            if job.name in running_key or state.done(job, tick):
                continue
            running_key.add(job.name)
            try:
                run_job(job, args.dry_run, args.max_pages)
                state.mark(job, tick, "ok")
            except Exception as e:
                state.mark(job, tick, "fail")
                log.error("[失败] %s：%s", job.name, e)
                log.debug("%s", traceback.format_exc())
            finally:
                running_key.discard(job.name)

    log.info("调度器已退出")
    return 0


if __name__ == "__main__":
    sys.exit(main())
