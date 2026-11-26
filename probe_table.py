#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单组件结构探针：把「接口到底返回了什么结构」照原样摊出来，用于校准解析规则。

    python probe_table.py                          # 探 config.xlsx 里第一个任务的全部组件
    python probe_table.py --base http://x:37799 --code nQcq
    python probe_table.py --widget 客户交叉 --out probe/
    python probe_table.py --json                   # 只出 JSON，便于贴给他人

它不写 Excel、不判对错，只回答三个问题：
  1. 这个组件被判成什么表型，依据是什么；
  2. 响应是哪种形态（平铺 / 分组树 / 交叉表视窗），列和行各多少；
  3. 首行原始报文长什么样、对齐后长什么样。
"""

import argparse
import copy
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from finebi_export import (FineBiShareClient, FineBiError, BIG_ROWS,      # noqa: E402
                           KIND_GROUP, KIND_CROSSTAB)
from table_norm import (flatten_header, col_names, detect_kind, row_values,   # noqa: E402
                        describe_filters, build_table, from_response,
                        is_cross_response, is_tree_rows)
from xlsx_io import safe_filename                                           # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def _short(obj, n=160):
    """把对象压成一行 JSON 便于打印，超长截断（原始报文有时几 MB，直接 print 会刷屏）。"""
    s = json.dumps(obj, ensure_ascii=False)
    return s if len(s) <= n else s[:n] + " …"


def _shape_of(data):
    """响应形态：6.1 起三类表各有各的结构。"""
    if is_cross_response(data):
        return "交叉表视窗（top/left 两棵树）"
    if is_tree_rows(data.get("items")):
        return "分组树（children/values）"
    return "平铺（header + items[]）"


def probe_one(client, w, out_dir=None, as_json=False):
    """探一个组件，返回结构摘要 dict。"""
    info = {
        "title": w.title, "widget_id": w.widget_id, "real_id": w.real_id,
        "type": w.type,
        "filters": describe_filters(w.detail or {}),
    }
    if not w.payload:
        info["skipped"] = w.skip_reason or "非表格组件"
        return info

    # 表型先按模型判（6.1 的 type 字段最可靠），再决定用哪种方式取
    kind, why = detect_kind(w.detail or {}, None)
    body = copy.deepcopy(w.payload)
    body["allData"] = True
    body["realData"] = True
    page = 1
    if kind in (KIND_GROUP, KIND_CROSSTAB):
        ts = body.setdefault("settings", {}).setdefault("tableStyle", {})
        ts["rowCounts"] = BIG_ROWS
        ts["colCounts"] = BIG_ROWS
        page = 0               # 6.1 的分组表/交叉表：0 才是「首次取数」

    data = client.fetch_page(page, payload=body)
    shape = _shape_of(data)

    table = from_response(w.title, w.detail or {}, data)
    if table is None:          # 6.0 的平铺结构，走老路
        header = data.get("header") or []
        cols, _cells = flatten_header(header)
        rows = [row_values(it, cols) for it in (data.get("items") or [])]
        table = build_table(w.title, w.detail or {}, header, rows, data.get("row"))

    info.update({
        "kind": table.kind, "why": why, "shape": shape,
        "col_count": len(table.cols),
        "header_depth": table.header_height,
        "columns": table.header,
        "row_count": len(table.rows),
        "row_total": data.get("row"),
        "has_next": data.get("hasNext") or data.get("hasXNext") or data.get("hasYNext"),
        "notes": table.notes,
    })
    raw_items = data.get("items") or []
    if raw_items and isinstance(raw_items[0], list):
        info["items_len_eq_cols"] = (len(raw_items[0]) == len(table.cols))
        info["first_row_raw"] = [_short(c, 60) for c in raw_items[0][:6]]
    info["first_row_norm"] = table.rows[0][:6] if table.rows else []

    if out_dir:
        base = safe_filename(w.title or w.widget_id)
        os.makedirs(out_dir, exist_ok=True)

        def dump(suffix, obj):
            with open(os.path.join(out_dir, base + suffix), "w", encoding="utf-8") as f:
                json.dump(obj, f, ensure_ascii=False, indent=1)

        dump("_detail.json", w.detail)
        dump("_data_p1.json", data)
        dump("_table.json", {"kind": table.kind, "cols": table.cols,
                             "header_cells": table.header_cells,
                             "merges": table.merges,
                             "rows_head": table.rows[:50]})
        info["dumped_to"] = out_dir

    if not as_json:
        print("\n" + "─" * 68)
        print("组件：%s   (widgetId=%s)" % (w.title, w.widget_id))
        print("表型：%s   ← %s" % (table.kind, why))
        print("结构：%s" % shape)
        if info["filters"]:
            print("筛选：%s" % "、".join(info["filters"]))
        print("列数：%d   表头层级：%d   行数：%d   服务端总行数：%s   hasNext=%s"
              % (len(table.cols), table.header_height, len(table.rows),
                 data.get("row"), info["has_next"]))
        for n in table.notes:
            print("   · %s" % n)
        if info.get("items_len_eq_cols") is not None:
            print("行长度 == 列数：%s%s"
                  % (info["items_len_eq_cols"],
                     "" if info["items_len_eq_cols"] else "（稀疏行，必须按 dId 对齐）"))
        if table.header_height > 1:
            print("表头树：")
            for c in table.header_cells:
                print("   row%d col%d  跨%d列/跨%d行  %s"
                      % (c["row"] + 1, c["col"] + 1, c.get("colspan", 1),
                         c.get("rowspan", 1), c["text"]))
        print("列名：%s" % " | ".join(table.header[:12]))
        if table.rows:
            if info.get("first_row_raw"):
                print("首行原始：%s" % " | ".join(str(x) for x in info["first_row_raw"]))
            print("首行对齐：%s" % " | ".join(str(x) for x in table.rows[0][:8]))
    return info


def main():
    """
    命令行入口：对指定组件逐个打印「表型判定 → 响应形态 → 表头树 → 首行对齐」。

    判定不对时先跑它，把输出贴出来就能定位是判型规则还是解析规则的问题；
    --out 会把原始报文落盘。
    """
    ap = argparse.ArgumentParser(description="FineBI 单组件结构探针")
    ap.add_argument("--base"), ap.add_argument("--code")
    ap.add_argument("--config", default="config.xlsx")
    ap.add_argument("--widget", action="append", help="只探指定组件（名称/ID，可多次/逗号分隔）")
    ap.add_argument("--out", help="原始报文输出目录")
    ap.add_argument("--json", action="store_true", help="只输出 JSON 摘要")
    ap.add_argument("--max-widgets", type=int, default=0, help="最多探几个（0=全部）")
    args = ap.parse_args()

    if args.base and args.code:
        base, code = args.base.rstrip("/"), args.code
    else:
        cfg = args.config if os.path.isabs(args.config) else os.path.join(HERE, args.config)
        try:
            from collector import load_config
            jobs = [j for j in load_config(cfg) if not j.error]
        except Exception as e:
            print("读配置失败：%s（可用 --base/--code 直接指定）" % e)
            return 2
        if not jobs:
            print("配置里没有可用任务：%s" % cfg)
            return 2
        base, code = jobs[0].base, jobs[0].code

    wanted = None
    if args.widget:
        wanted = [x.strip() for one in args.widget for x in one.split(",") if x.strip()]

    client = FineBiShareClient(base, code).bootstrap()
    client.fetch_details()
    targets = client.resolve_widgets(wanted) or client.widgets
    if args.max_widgets:
        targets = targets[:args.max_widgets]

    out_dir = None
    if args.out:
        out_dir = args.out if os.path.isabs(args.out) else os.path.join(HERE, args.out)

    if not args.json:
        print("页面：%s  共 %d 个组件，将探 %d 个"
              % (client.report_name or client.report_id, len(client.widgets), len(targets)))

    infos = []
    for w in targets:
        try:
            infos.append(probe_one(client, w, out_dir, args.json))
        except FineBiError as e:
            infos.append({"title": w.title, "error": str(e)})
            if not args.json:
                print("\n组件：%s  探针失败：%s" % (w.title, e))

    if args.json:
        print(json.dumps(infos, ensure_ascii=False, indent=1))

    # 判定兜底：全是「未知」通常意味着结构变了，提示用样本对齐
    kinds = [i.get("kind") for i in infos]
    if not args.json:
        print("\n" + "=" * 68)
        print("汇总：%s" % "、".join("%s=%s" % (i["title"], i.get("kind") or "跳过")
                                     for i in infos))
        if all(k in (None, "未知") for k in kinds):
            print("没有识别出任何表格类型：结构可能变了。"
                  "把 --out 目录里的 *_detail.json 和 *_table.json 发我，我照着改判定规则。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
