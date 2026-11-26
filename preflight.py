#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FineBI 版本兼容性自检（升级前后对拍用）。

它不导出全量数据，只走一遍最关键的 5 步链路，逐步给出 OK / WARN / FAIL，
并把结果存成一份 fingerprint JSON。升级前存一次、升级后再存一次，
用 --compare 就能直接看出「哪一步变了」。

    python preflight.py                                  # 用 config.xlsx 里第一个启用任务
    python preflight.py --base http://10.0.0.5:37799 --code z8Pz
    python preflight.py --save before.json               # 升级前：存基线
    python preflight.py --compare before.json            # 升级后：与基线对比
    python preflight.py --all-jobs                       # 配置里所有任务都查一遍

退出码：0 = 兼容（可能有 WARN），1 = 不兼容（有 FAIL）。
"""

import argparse
import datetime as _dt
import json
import os
import sys
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from finebi_export import (FineBiShareClient, FineBiError, probe_version, cell_text,  # noqa: E402
                           KIND_DETAIL, detect_kind)

HERE = os.path.dirname(os.path.abspath(__file__))
OK, WARN, FAIL = "OK  ", "WARN", "FAIL"
_fails = []
_warns = []


def say(level, msg, hint=""):
    """打一行结论并记账：FAIL/WARN 会累计到进程退出码与指纹文件里。"""
    print("[%s] %s" % (level, msg))
    if hint:
        print("        %s" % hint)
    if level == FAIL:
        _fails.append(msg)
    elif level == WARN:
        _warns.append(msg)


def step_no(i, title):
    """打印步骤小标题（1~5 走链路，第 6 步是与基线对拍）。"""
    print()
    print("── %d. %s %s" % (i, title, "─" * max(0, 54 - len(title))))


# ------------------------------------------------------------------ 各步骤
def s1_site(base, fp):
    """站点可达 + 版本号。"""
    step_no(1, "站点可达性与版本")
    ver, build = probe_version(base)
    fp["version"], fp["build"] = ver, build
    if not ver:
        say(FAIL, "连不上或识别不出 FineBI 版本：%s" % base,
            "确认地址端口、防火墙；纯内网需在能访问 BI 的机器上跑")
        return False
    say(OK, "FineBI 版本 %s（%s）" % (ver, build or "构建时间未知"))
    # 只做提示，不做拦截：不同小版本走的是同一套 v5 路由
    if not ver.startswith(("6.", "11.")):
        say(WARN, "版本号 %s 不在实测范围（6.0.27 / 6.1.x）" % ver,
            "仍可继续，但要重点看后几步结果")
    return True


def s2_link(client, fp):
    """分享链接 -> 会话 / 令牌 / 报表 ID。"""
    step_no(2, "分享链接解析（sessionId / 令牌 / reportId）")
    try:
        client.bootstrap()
    except FineBiError as e:
        say(FAIL, "分享链接解析失败：%s" % e,
            "链接失效、需要登录，或页面模板变了（可看 finebi_export.py 的 bootstrap 规则）")
        return False
    fp["session"] = {
        "has_sessionId": bool(client.session_id),
        "has_jwt": bool(client.jwt),
        "report_id": client.report_id,
        "report_name": client.report_name,
    }
    say(OK, "sessionId 已取得（%s…）" % (client.session_id or "")[:8])
    say(OK, "分享令牌已取得（%s…，%d 字符）" % ((client.jwt or "")[:12], len(client.jwt or "")))
    say(OK, "reportId=%s  页面名=%s" % (client.report_id, client.report_name or "(空)"))
    ids = [x for x in (client.report_ids or []) if x]
    if len(ids) > 1:
        say(WARN, "页面里有 %d 个 reportId：%s" % (len(ids), "、".join(ids[:4])),
            "这是 6.1 的「一次分享多个资源」链接；本工具只处理第一个仪表板，"
            "其余资源不会导出。需要全要的话，为每个仪表板单独建公共链接、各配一行任务")
    return True


def s3_widgets(client, fp):
    """组件清单解析。"""
    step_no(3, "组件清单解析（reportWidgets）")
    tables = client.resolve_widgets()
    filters = client.filter_widgets()
    fp["widget_count"] = len(tables)
    fp["filter_count"] = len(filters)
    if not tables:
        say(FAIL, "解析出 0 个可导出组件", "页面结构可能变了，或该链接本身没有组件")
        return False
    say(OK, "解析出 %d 个表格组件%s"
        % (len(tables), "、%d 个过滤组件" % len(filters) if filters else "（页面无过滤组件）"))
    for i, w in enumerate(tables, 1):
        # 真实组件名要到第 4 步的 widgets/detail 才能拿到，这里先标明未命名
        print("        %d) %-22s widgetId=%s"
              % (i, "（名称待第 4 步补齐）" if w.title_is_auto else w.title, w.widget_id))
    for w in filters:
        print("        · 过滤组件 %-12s 绑定字段=%s" % (w.title, w.field or "?"))

    # 链接上带的筛选参数：这里必须验证，因为参数名写错时脚本会安静地导全量
    if client.url_params:
        if client.filters:
            say(OK, "生效筛选：%s" % client.filter_summary())
        if client.unmatched_params:
            names = "、".join(client.unmatched_params)
            have = "、".join(w.title for w in filters) or "无"
            say(WARN, "参数 %s 没有同名过滤组件，未生效（可用名：%s）" % (names, have),
                "参数名要写过滤组件的标题，不是数据字段名")
    return True


def s4_details(client, fp):
    """组件数据模型 + 关键字段位置。"""
    step_no(4, "组件数据模型（widgets/detail）")
    try:
        client.fetch_details()
    except FineBiError as e:
        say(FAIL, "取组件详情失败：%s" % e, "接口路径或参数可能在 6.1 有变化")
        return False

    recs, table_like = [], 0
    for w in [x for x in client.widgets if not x.is_filter]:
        settings = (w.payload or {}).get("settings") or {}
        has_ts = "tableStyle" in settings
        if has_ts or w.notes:
            table_like += 1
        recs.append({
            "title": w.title,
            "widget_id": w.widget_id,
            "real_id": w.real_id,
            "type": w.type,
            "has_tableStyle": has_ts,
            "settings_keys": sorted(settings.keys()),
            "payload_ok": w.payload is not None,
            "skip_reason": w.skip_reason,
        })
        mark = OK if w.payload is not None else WARN
        say(mark, "%-22s %s" % (w.title, "可导出" if w.payload is not None else
                                "跳过：" + (w.skip_reason or "未知")))
        if w.notes:
            say(WARN, "  %s：%s" % (w.title, w.notes), "已按兜底逻辑继续，建议核对行数")
    fp["widgets"] = recs
    if table_like == 0:
        say(FAIL, "没有一个组件被识别为明细表", "settings.tableStyle 的位置可能变了")
        return False
    # 详情里也拿不到名字，才会影响输出文件名，这时才值得提示
    auto = [w.title for w in client.widgets if w.title_is_auto and not w.is_filter]
    if auto:
        say(WARN, "%d 个组件连 widgets/detail 里都没有 name，只能用序号命名：%s"
            % (len(auto), "、".join(auto)), "导出文件名会是「组件N_日期.xlsx」")
    return True


def s5_data(client, fp):
    """真实取 2 页：验证 totalRows 与翻页语义。"""
    step_no(5, "取数验证（第 1 页 + 第 2 页）")
    target = None
    for w in client.widgets:
        if w.is_filter:
            continue
        if w.payload:
            target = w
            break
    if target is None:
        say(FAIL, "没有可试取数的组件")
        return False
    kind, why = detect_kind(target.detail or {}, None)
    say(OK, "试取组件：%s（%s，%s）" % (target.title, kind or "未知", why))

    # 用统一的取数入口：明细表走翻页，分组表/交叉表一次取全（6.1 的新形态）
    try:
        t = client.fetch_table(target, max_pages=1, log=lambda m: None)
    except FineBiError as e:
        say(FAIL, "第 1 次取数失败：%s" % e, "载荷结构可能不被该版本接受")
        return False

    header = t.header
    fp["sample"] = {
        "widget": target.title,
        "kind": t.kind,
        "header": header,
        "col_count": len(t.cols),
        "header_height": t.header_height,
        "row_count": len(t.rows),
        "row_total": t.expected_total,
        "has_next": getattr(t, "has_next", None),
        "notes": t.notes,
    }
    say(OK, "表类型=%s，%d 列 × %d 行，表头 %d 层"
        % (t.kind, len(t.cols), len(t.rows), t.header_height))
    if header:
        say(OK, "列名：%s" % (" | ".join(header[:8]) + (" …" if len(header) > 8 else "")))
    for n in t.notes:
        say(OK, "· %s" % n)

    if not t.rows:
        say(FAIL, "没有取到任何数据行")
        return False

    if t.kind != KIND_DETAIL:
        # 6.1 的分组表/交叉表不按页码翻页，能取到数据 + 没有截断标记就算通过
        if getattr(t, "has_next", False):
            say(WARN, "服务端标记还有更多数据（hasNext），当前只取到一部分",
                "把 FINEBI_BIG_ROWS 调大（默认 200000）")
        else:
            say(OK, "单次取全（无翻页需求）")
        return True

    # 明细表：原来的翻页验证照旧
    total = t.expected_total
    if isinstance(total, int) and total in (1000, 2000, 5000, 10000, 20000):
        say(WARN, "服务端回报总行数正好是 %d（常见截断值）" % total,
            "若实际数据远多于该值，用环境变量 FINEBI_TOTAL_ROWS=-1 再试")

    if isinstance(total, int) and total > len(t.rows):
        time.sleep(0.2)
        try:
            d2 = client.fetch_page(2, payload=target.payload)
        except FineBiError as e:
            say(FAIL, "第 2 页取数失败：%s" % e, "翻页参数语义可能变了")
            return False
        items2 = d2.get("items") or []
        fp["sample"]["page2_size"] = len(items2)
        if not items2:
            say(FAIL, "hasNext=true 但第 2 页为空", "翻页可能失效，会导致数据截断")
            return False
        a, b = t.rows[0], [cell_text(c) for c in items2[0]]
        fp["sample"]["page2_first_row_same"] = (a == b)
        if a == b:
            say(FAIL, "第 1、2 页首行完全相同", "page 参数可能失效，全量导出会拿到重复数据")
            return False
        say(OK, "第 2 页 %d 行，首行与第 1 页不同（翻页正常）" % len(items2))
    else:
        say(OK, "单页即取完（无翻页需求）")
    return True


# ------------------------------------------------------------------ 对拍
def compare_with(fp, path):
    """
    与基线指纹逐项对比，把「变化」标出来。

    升级后哪一步变了，这里会直接指出字段名——比重新理解一遍接口结构快得多。
    基线文件读不了只告警不中断：自检本身的结果仍然有效。
    """
    step_no(6, "与基线对比：%s" % os.path.basename(path))
    try:
        with open(path, "r", encoding="utf-8") as f:
            old = json.load(f)
    except Exception as e:
        say(WARN, "基线读不了：%s" % e)
        return

    print("        %-16s %-22s %s" % ("项目", "基线", "当前"))
    print("        " + "-" * 62)

    def row(name, a, b):
        flag = "" if a == b else "  <-- 变化"
        print("        %-16s %-22s %s%s" % (name, str(a)[:22], str(b)[:22], flag))

    row("版本", old.get("version"), fp.get("version"))
    row("组件数", old.get("widget_count"), fp.get("widget_count"))

    on = [w.get("title") for w in old.get("widgets") or []]
    nn = [w.get("title") for w in fp.get("widgets") or []]
    row("组件名", "、".join(on)[:22], "、".join(nn)[:22])

    os_ = old.get("sample") or {}
    ns = fp.get("sample") or {}
    row("列数", os_.get("col_count"), ns.get("col_count"))
    row("表头", " | ".join(os_.get("header") or [])[:22],
        " | ".join(ns.get("header") or [])[:22])
    row("总行数", os_.get("row_total"), ns.get("row_total"))
    row("每页行数", os_.get("page_size"), ns.get("page_size"))

    # settings 字段增删最能反映内部模型有没有动
    def keys_of(x):
        s = set()
        for w in x.get("widgets") or []:
            s.update(w.get("settings_keys") or [])
        return s
    ko, kn = keys_of(old), keys_of(fp)
    row("settings 字段数", len(ko), len(kn))
    if kn - ko:
        say(WARN, "新增 settings 字段：%s" % "、".join(sorted(kn - ko)),
            "一般无碍，留意是否有影响行数上限的新字段")
    if ko - kn:
        say(FAIL, "丢失 settings 字段：%s" % "、".join(sorted(ko - kn)),
            "若有 tableStyle/totalRows 相关字段，需按新结构改 build_payload")
    if not (kn - ko) and not (ko - kn) and old.get("version") == fp.get("version"):
        say(OK, "结构与基线完全一致")


# ------------------------------------------------------------------ 主流程
def run_one(base, code, save, compare, label="", dump=None):
    """
    对一个分享链接跑完 5 步自检，返回指纹 dict。

    fp 就是「指纹」：版本、组件清单、表头、总行数、关键 settings 字段。
    升级前 --save 存一份，升级后 --compare 直接对拍，不必靠记忆比对界面。
    """
    fp = {
        "checked_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "python": sys.version.split()[0],
        "base": base, "code": code,
        "widget_count": 0, "widgets": [], "sample": {},
        "verdict": "",
    }
    print("=" * 66)
    print("FineBI 兼容性自检  %s" % (("[" + label + "]") if label else ""))
    print("%s  code=%s   Python %s" % (base, code, fp["python"]))
    print("=" * 66)

    client = FineBiShareClient(base, code)
    if dump:
        client.dump_dir = dump if os.path.isabs(dump) else os.path.join(HERE, dump)
    try:
        if s1_site(base, fp) and s2_link(client, fp) and s3_widgets(client, fp) \
                and s4_details(client, fp) and s5_data(client, fp):
            title_fixed = [w for w in fp["widgets"] if w["payload_ok"]]
            fp["verdict"] = "兼容（%d 个可导出组件，有 %d 项告警）" % (len(title_fixed), len(_warns))
        else:
            fp["verdict"] = "不兼容（%d 项阻塞）" % len(_fails)
    except Exception:
        traceback.print_exc()
        say(FAIL, "自检过程中抛异常，见上方堆栈")
        fp["verdict"] = "不兼容（异常中断）"

    if compare:
        compare_with(fp, compare)
        # 基线对比可能新增阻塞项（比如内部字段消失），结论要跟着改
        if _fails:
            fp["verdict"] = "不兼容（与基线对比有 %d 项阻塞）" % len(_fails)
        elif _warns and fp["verdict"].startswith("兼容"):
            fp["verdict"] = "兼容（有 %d 项告警）" % len(_warns)

    print()
    print("=" * 66)
    print("结论：%s" % fp["verdict"])
    if _fails:
        for f in _fails:
            print("  阻塞：%s" % f)
    if _warns:
        for w in _warns:
            print("  告警：%s" % w)
    print("=" * 66)

    if save:
        out = save if os.path.isabs(save) else os.path.join(HERE, save)
        with open(out, "w", encoding="utf-8") as f:
            json.dump(fp, f, ensure_ascii=False, indent=2)
        print("指纹已保存：%s（升级后用它 --compare）" % out)
    if dump:
        print("原始报文已留存：%s" % client.dump_dir)
        print("        （分享页 HTML + 每个请求体/响应体，排查结构变化时直接看这些文件）")
    return fp


def main():
    """
    命令行入口。目标链接取自 --base/--code，没给就用 config.xlsx 里第一个启用任务。

    退出码约定：0 = 兼容（可能有 WARN），1 = 不兼容（有 FAIL）。
    """
    ap = argparse.ArgumentParser(description="FineBI 版本兼容性自检")
    ap.add_argument("--base", help="FineBI 站点根，如 http://10.0.0.5:37799")
    ap.add_argument("--code", help="分享链接 code（link/<code>）")
    ap.add_argument("--config", default="config.xlsx", help="配置 xlsx（默认 config.xlsx）")
    ap.add_argument("--save", help="把指纹存成 JSON，如 before.json")
    ap.add_argument("--compare", help="与之前保存的指纹对比，如 before.json")
    ap.add_argument("--all-jobs", action="store_true", help="配置里所有启用任务都查一遍")
    ap.add_argument("--dump", help="把原始 HTML/请求/响应存到指定目录（排查版本差异用）")
    args = ap.parse_args()

    targets = []
    if args.base and args.code:
        targets = [(args.base.rstrip("/"), args.code, "")]
    else:
        cfg = args.config if os.path.isabs(args.config) else os.path.join(HERE, args.config)
        try:
            from collector import load_config
            jobs = [j for j in load_config(cfg) if not j.error and j.enabled]
        except Exception as e:
            print("读配置失败：%s" % e)
            print("改用 --base 与 --code 直接指定。")
            return 2
        if not jobs:
            print("配置里没有启用的任务：%s" % cfg)
            return 2
        if not args.all_jobs:
            jobs = jobs[:1]
        targets = [(j.base, j.code, j.name) for j in jobs]

    results = []
    for base, code, label in targets:
        _fails.clear()
        _warns.clear()
        results.append(run_one(base, code, args.save if len(targets) == 1 else None,
                               args.compare, label,
                               args.dump if len(targets) == 1 else None))
        if len(targets) > 1:
            print()

    bad = [r for r in results if r["verdict"].startswith("不兼容")]
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
