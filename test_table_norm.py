#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table_norm / xlsx_io 的自测（不需要 FineBI，也不需要 pytest）

    python test_table_norm.py

覆盖三类表的典型形态：
  1. 明细表：平铺表头、每行长度 = 列数
  2. 分组表：维度 + 指标、维度列连续同值（合并/填充/原样三种模式）
  3. 交叉表：多级表头（children）、合并区行只返回首格（靠 dId 对齐）
另外校验写出的 xlsx：XML 合法性、mergeCells 数量、表头行数、冻结行。

注意：样本是按 6.0.x / 6.1 的响应结构手写的（6.1 的三种形态见文件末尾），
用于验证「代码契约」；
真机结构以 probe_table.py 抓到的为准。
"""
import json
import sys
import zipfile
from xml.etree import ElementTree as ET

import table_norm as tn
from xlsx_io import write_workbook

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
ok = 0
bad = []


def check(cond, msg, extra=""):
    global ok
    if cond:
        ok += 1
        print("  [OK  ] %s" % msg)
    else:
        bad.append(msg + (" :: " + extra if extra else ""))
        print("  [FAIL] %s %s" % (msg, extra))


def cell(text, did=None):
    return {"dId": did or "", "text": text, "value": text}


# ------------------------------------------------------------------ 1 明细表
def t_detail():
    print("\n1) 明细表")
    header = [{"dId": "c1", "text": "客户号"}, {"dId": "c2", "text": "姓名"},
              {"dId": "c3", "text": "手机号"}]
    detail = {"settings": {"tableStyle": {"totalRows": 0}}, "measures": [],
              "dimensions": {"c1": {}, "c2": {}, "c3": {}}, "view": {"10000": ["c1", "c2", "c3"]}}
    cols, cells = tn.flatten_header(header)
    rows = [tn.row_values([cell("0001", "c1"), cell("张三", "c2"), cell("13871014583", "c3")], cols)]
    t = tn.build_table("明细", detail, header, rows, expected_total=1)

    check(t.kind == tn.KIND_DETAIL, "识别为明细表", t.kind)
    check(len(cols) == 3, "列数 3")
    check(t.header_height == 1, "表头 1 行", t.header_height)
    check(not t.merges, "无合并", str(t.merges))
    check(t.cols[2]["text"] == "手机号", "表头文本正确")
    check(t.notes[0].find("明细表") > 0, "判定依据已记录", t.notes[0])


# ------------------------------------------------------------------ 2 分组表
def t_group(mode):
    print("\n2) 分组表（分组列 = %s）" % mode)
    header = [{"dId": "c1", "text": "机构"}, {"dId": "c2", "text": "产品"},
              {"dId": "m1", "text": "AUM"}, {"dId": "m2", "text": "户数"}]
    detail = {"settings": {"tableStyle": {"totalRows": 0}},
              "measures": [{"id": "m1"}, {"id": "m2"}],
              "dimensions": {"c1": {}, "c2": {}},
              "view": {"10000": ["c1", "c2"], "20000": []}}
    rows = [["北京", "A", "1", "2"],
            ["北京", "A", "3", "4"],
            ["北京", "B", "5", "6"],
            ["上海", "A", "7", "8"]]
    t = tn.build_table("分组", detail, header, rows, expected_total=4, group_mode=mode)

    check(t.kind == tn.KIND_GROUP, "识别为分组表", t.kind)
    check(t.header_height == 1, "表头 1 行")
    if mode == "merge":
        # 行号 = 表头 1 行 + 数据下标 + 1 -> 北京(第2行) 连续 3 行 -> A2:A4；产品 A 连续 2 行 -> B2:B3
        check((0, 2, 0, 4) in t.merges, "机构列 A2:A4 合并", str(t.merges))
        check((1, 2, 1, 3) in t.merges, "产品列 B2:B3 合并", str(t.merges))
        check(t.rows[1][0] == "北京", "合并不改数据（第 2 行仍是北京）")
    elif mode == "fill":
        check(t.rows[2][0] == "北京", "空值被向下填充", str(t.rows[2]))
        check(not t.merges, "填充模式下不合并")
    else:
        check(t.rows[2][0] == "北京", "原样模式不动数据")


def t_group_blank():
    print("\n2b) 分组表（服务端把重复维度给成空串）")
    header = [{"dId": "c1", "text": "机构"}, {"dId": "m1", "text": "AUM"}]
    detail = {"settings": {"tableStyle": {}}, "measures": [{"id": "m1"}],
              "dimensions": {"c1": {}}, "view": {}}
    rows = [["北京", "1"], ["", "2"], ["", "3"], ["上海", "4"]]
    t = tn.build_table("分组2", detail, header, rows, group_mode="fill")
    check(t.rows[1][0] == "北京" and t.rows[2][0] == "北京", "后续空值被填成北京", str(t.rows))
    t2 = tn.build_table("分组2", detail, header, rows, group_mode="merge")
    check(not t2.merges, "原样含空串时不硬合并", str(t2.merges))


# ------------------------------------------------------------------ 3 交叉表
def t_crosstab():
    print("\n3) 交叉表")
    header = [{"dId": "c1", "text": "机构"},
              {"dId": "", "text": "AUM",
               "children": [{"dId": "m1", "text": "1月"}, {"dId": "m2", "text": "2月"}]}]
    detail = {"settings": {"tableStyle": {}}, "measures": [{"id": "m1"}, {"id": "m2"}],
              "dimensions": {"c1": {}}, "view": {"10000": ["c1"], "20000": ["col1"]}}
    cols, cells = tn.flatten_header(header)
    check(len(cols) == 3, "叶子列 3（机构 + 两个月）", str([c["text"] for c in cols]))
    check(tn.col_names(cols) == ["机构", "AUM/1月", "AUM/2月"], "多级列名拼接", str(tn.col_names(cols)))

    # 合并区：第二行的机构列不返回，只给两个指标
    items = [[cell("北京", "c1"), cell("10", "m1"), cell("20", "m2")],
             [cell("30", "m1"), cell("40", "m2")]]
    rows = [tn.row_values(it, cols) for it in items]
    check(rows[1] == ["", "30", "40"], "缺首格的行按 dId 对齐（不错列）", str(rows[1]))

    t = tn.build_table("交叉", detail, header, rows, expected_total=2)
    check(t.kind == tn.KIND_CROSSTAB, "识别为交叉表", t.kind)
    check(t.header_height == 2, "表头 2 行", t.header_height)
    check((0, 1, 0, 2) in t.merges, "机构纵向合并 A1:A2", str([m for m in t.merges if m[0] == 0]))
    check((1, 1, 2, 1) in t.merges, "AUM 横向合并 B1:C1", str([m for m in t.merges if m[1] == 1]))


# ------------------------------------------------------------------ 4 行数自证
def t_expected():
    print("\n4) 行数自证")
    header = [{"dId": "c1", "text": "x"}]
    detail = {"settings": {"tableStyle": {}}, "measures": [], "dimensions": {"c1": {}}}
    t = tn.build_table("x", detail, header, [["1"], ["2"]], expected_total=110000)
    check(any("行数不符" in n for n in t.notes), "少于服务端总行数会留痕", str(t.notes))


# ------------------------------------------------------------------ 5 筛选摘要
def t_filters():
    print("\n5) 筛选条件摘要")
    detail = {"settings": {"tableStyle": {}}, "parameters": [],
              "filterValues": [{"f": 1}, {"f": 2}],
              "widgetParameters": [{"name": "日期"}]}
    got = tn.describe_filters(detail)
    check(any("filterValues（2 项）" in x for x in got), "识别 filterValues 2 项", str(got))
    check(any("组件参数 1 个" in x for x in got), "识别组件参数", str(got))
    check(not any("parameters" in x for x in got if "组件参数" not in x), "空参数不报", str(got))


# ------------------------------------------------------------------ 6 写 xlsx
def t_xlsx():
    print("\n6) xlsx 写出（多级表头 + 合并）")
    header = [{"dId": "c1", "text": "机构"},
              {"dId": "", "text": "AUM",
               "children": [{"dId": "m1", "text": "1月"}, {"dId": "m2", "text": "2月"}]}]
    detail = {"settings": {"tableStyle": {}}, "measures": [{"id": "m1"}, {"id": "m2"}],
              "dimensions": {"c1": {}}, "view": {"20000": ["x"]}}
    rows = [["北京", "10", "20"], ["上海", "30", "40"]]
    t = tn.build_table("交叉", detail, header, rows, group_mode="merge")
    path = "/tmp/_tn_test.xlsx"
    write_workbook(path, [t.sheet("交叉")])

    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        for part in names:
            if part.endswith(".xml") or part.endswith(".rels"):
                ET.fromstring(z.read(part))       # 全部部件都要是合法 XML
    check(any(x.endswith("sheet1.xml") for x in names), "workbook 结构完整")
    root = ET.fromstring(xml)
    mc = root.find(NS + "mergeCells")
    n = int(mc.get("count")) if mc is not None else 0
    check(n == 2, "mergeCells 2 处（机构纵向 + AUM 横向）", str(n))
    check('ySplit="2"' in xml, "冻结在表头下方（2 行）")
    rows_xml = root.findall(".//" + NS + "row")
    check(len(rows_xml) == 4, "总行数 4（2 行表头 + 2 行数据）", str(len(rows_xml)))
    # 表头第一行应有两个带 s=1 的单元格（机构、AUM），1月/2月在第 2 行
    r2 = [r for r in rows_xml if r.get("r") == "2"][0]
    check(len(r2.findall(NS + "c")) == 2, "第 2 行表头有 2 个单元格", str(len(r2.findall(NS + "c"))))



# ---------------------------------------------------------------- 6.1 的结构
# 6.1 起分组表/交叉表的响应换了形态（维度树、视窗），这里用合成样本锁住解析契约。
def t_61_group_tree():
    """6.1 分组表：items 是维度树，根节点（text=None）是合计行。"""
    data = {
        "header": [{"dId": "d1", "text": "性别"}, {"dId": "d2", "text": "姓名"},
                   {"dId": "m1", "text": "記錄數"}],
        "hasNext": False,
        "items": [{
            "text": None, "dId": None, "values": [{"formatText": "10"}],
            "children": [
                {"text": "女", "dId": "d1", "values": [], "children": [
                    {"text": "甲", "dId": "d2", "children": [],
                     "values": [{"formatText": "4"}]},
                    {"text": "乙", "dId": "d2", "children": [],
                     "values": [{"formatText": "1"}]},
                ]},
                {"text": "男", "dId": "d1", "values": [], "children": [
                    {"text": "丙", "dId": "d2", "children": [],
                     "values": [{"formatText": "5"}]},
                ]},
            ],
        }],
    }
    t = tn.from_response("g", {"type": 1}, data)
    check(t is not None and t.kind == tn.KIND_GROUP, "6.1 分组树 -> 分组表")
    check(len(t.cols) == 3, "6.1 分组树列数 = 2 维度 + 1 指标", str(t.cols))
    check(len(t.rows) == 4, "6.1 分组树行数 = 合计 + 3 叶子", str(len(t.rows)))
    check(t.rows[0] == ["合计", "", "10"], "合计行落首行并补了标签", str(t.rows[0]))
    check(t.rows[1] == ["女", "甲", "4"], "叶子行 = 维度路径 + 指标值", str(t.rows[1]))
    check(bool(t.merges), "同维度值纵向合并（merge 模式）", str(t.merges[:2]))
    t2 = tn.from_response("g", {"type": 1}, data, group_mode="fill")
    check(t2.rows[2][0] == "女", "fill 模式下维度值向下填满", str(t2.rows[2]))


def t_61_cross_window():
    """6.1 交叉表：top 是列头树、left 是行头树，值按列叶子顺序排。"""
    data = {
        "hasXNext": False, "hasYNext": False,
        "top": {
            "header": [{"dId": "c1", "text": "类型"}, {"dId": "m", "text": "指标"}],
            "items": [{"text": None, "dId": None, "children": [
                {"text": "A", "dId": "c1", "children": [
                    {"text": "数", "dId": "m", "values": [{"formatText": ""}]},
                    {"text": "额", "dId": "m", "values": [{"formatText": ""}]},
                ]},
                {"text": "B", "dId": "c1", "children": [
                    {"text": "数", "dId": "m", "values": [{"formatText": ""}]},
                    {"text": "额", "dId": "m", "values": [{"formatText": ""}]},
                ]},
            ]}],
        },
        "left": {
            "header": [{"text": "性别"}] + [{"text": "--"}] * 4,
            "items": [{"text": None, "dId": None, "children": [
                {"text": "女", "dId": "r1", "children": [
                    {"text": "合计", "dId": "__summary_dim_id__", "children": [],
                     "values": [{"formatText": "1"}, {"formatText": "2"},
                                {"formatText": "3"}, {"formatText": "4"}]}]},
            ]}],
        },
    }
    t = tn.from_response("c", {"type": 2}, data)
    check(t is not None and t.kind == tn.KIND_CROSSTAB, "6.1 交叉视窗 -> 交叉表")
    check(len(t.cols) == 5, "列数 = 1 行维度 + 4 列叶子", str(t.cols))
    check(t.header_height == 2, "表头层级 = 2", str(t.header_height))
    check(t.rows[0] == ["女", "1", "2", "3", "4"], "行值按列叶子顺序对齐", str(t.rows[0]))
    span2 = [c for c in t.header_cells if c.get("colspan", 1) == 2]
    check(len(span2) == 2, "同父列合并成跨 2 列", str([c["text"] for c in span2]))
    check(any(c.get("rowspan", 1) == 2 for c in t.header_cells),
          "行维度标题纵向跨满表头")
    data["hasYNext"] = True
    t2 = tn.from_response("c", {"type": 2}, data)
    check(any("hasYNext" in n for n in t2.notes), "有更多行时给出截断提示", str(t2.notes))


def t_61_flat_fallback():
    """6.1 明细表仍是平铺，from_response 应返回 None 让调用方走老路。"""
    data = {"header": [{"dId": "a", "text": "A"}], "items": [[{"dId": "a", "text": "1"}]]}
    check(tn.from_response("d", {"type": 4}, data) is None,
          "平铺结构不走 6.1 分支（返回 None）")
    check(tn.detect_kind({"type": 2, "settings": {"tableStyle": {}}})[0] == tn.KIND_CROSSTAB,
          "detect_kind 优先用组件的 type 字段")


def t_filters_61():
    """
    6.1 筛选链路：过滤组件解析 -> 按名字匹配 -> 载荷注入。

    这几步任何一环断了，表现都是「静默导全量」——数据看着完全正常，
    只是没筛。所以固定成用例，别让重构悄悄改坏。
    """
    import finebi_export as fx

    widgets = {
        "w1": {"type": "1", "title": '${"alt":"__widget_name_id__"}', "realWidgetId": "r1"},
        "f1": {"type": 2, "showTitle": True,
               "widgetBean": {"name": '<p class="align-left">性别</p>',
                              "dimensions": {"d1": {"name": "性别"}}}},
    }
    # 真实页面是多层转义的：先序列化一次，再整体当字符串序列化一次
    one = json.dumps(widgets, ensure_ascii=False)
    two = json.dumps(one, ensure_ascii=False)[1:-1]
    html = 'var cfg = {"reportWidgets":%s};' % two

    ws = fx.FineBiShareClient._parse_widgets(html)
    check(len(ws) == 2, "多层转义下解析出全部组件", str([w.widget_id for w in ws]))
    fw = [w for w in ws if w.is_filter]
    check(len(fw) == 1 and fw[0].title == "性别", "识别出过滤组件，名称去掉了 HTML 标签",
          str([(w.widget_id, w.title) for w in ws]))

    c = fx.FineBiShareClient("http://x", "nQcq")
    c.widgets = ws
    check(c.set_filter("性别", ["女"]) == 1, "按组件标题匹配到过滤组件")
    check(c.set_filter("性别", ["男"], mode="append") == 1, "append 模式可追加")
    check(c.filters["f1"]["values"] == ["女", "男"], "追加后值合并去重",
          str(c.filters["f1"]["values"]))
    check(c.set_filter("不存在", ["x"]) == 0, "无同名组件时不命中（由调用方告警）")

    w = fx.Widget("w1", "r1", "明细", 1)
    detail = {"settings": {"tableStyle": {}}, "name": "明细"}
    p = c.build_payload(detail, w)
    check(p["filter"]["filterType"] == 34
          and p["filter"]["filterValue"][0]["filterType"] == 56
          and p["filter"]["filterValue"][0]["filterValue"]["widgetId"] == "f1",
          "载荷 filter 用 filterType=56 引用过滤组件")
    vm = p["queryInfo"]["widgetIdValueMap"]["f1"]
    check(vm["value"] == {"type": 1, "value": ["女", "男"], "assist": []},
          "选中值写进 widgetIdValueMap 的 value")

    c2 = fx.FineBiShareClient("http://x", "nQcq")
    p2 = c2.build_payload(detail, w)
    check("widgetIdValueMap" not in p2["queryInfo"], "无筛选时不注入 widgetIdValueMap")
    check(p2["filter"]["filterValue"] == [], "无筛选时 filter 为空（走默认值）")

    import collector
    _base, code = collector.parse_url("http://h:37799/webroot/decision/link/nQcq?性别=女")
    check(code == "nQcq?性别=女", "配置里的链接参数不会被丢掉", code)


def t_filter_value_shape():
    """
    过滤组件的值结构：文本类能写，日期/区间类必须拒绝。

    这是企业版真机踩出来的：链接里带 ?日期=2026-11-23，脚本按文本把值写成
    {"type":1,"value":["2026-11-23"]}，服务端按 DateStaticValueBean 反序列化，
    直接 400（Cannot deserialize ... from Array value），整个仪表板十四个组件全挂。
    所以固定成用例：对象型值一律返回 None，宁可不筛，也不发必挂的载荷。
    """
    import finebi_export as fx

    text_bean = {"value": {"type": 1, "value": ["女"], "assist": []}}
    check(fx.shape_filter_value(text_bean, ["男"]) == {"type": 1, "value": ["男"], "assist": []},
          "文本类：值按字符串数组覆盖")

    empty_bean = {"value": {"type": 1, "value": [], "assist": []}}
    check(fx.shape_filter_value(empty_bean, ["女"])["value"] == ["女"],
          "文本类：原来没选值时也能写入")

    date_bean = {"type": 32, "showTime": False,
                 "value": {"type": 2,
                           "value": {"year": 2026, "month": 11, "day": 23,
                                     "quarter": 0, "period": 0}}}
    check(fx.shape_filter_value(date_bean, ["2026-11-23"]) is None,
          "日期类：拒绝按文本写入（否则服务端 400）")
    check(fx.filter_kind(date_bean) == "日期或区间类", "日期类能被识别出来用于提示")
    check(fx.filter_kind(text_bean) == "文本类", "文本类判定不受影响")

    range_bean = {"value": {"type": 3, "value": {"start": 1, "end": 9}}}
    check(fx.shape_filter_value(range_bean, ["1", "9"]) is None, "区间类：同样拒绝")

    check(fx.shape_filter_value({}, ["女"])["value"] == ["女"], "没有 value 字段时按文本处理")

    # set_filter 命中日期组件：要记进 unsupported_filters，且不能进 filters
    page = {"reportWidgets": {
        "f1": {"type": 2, "widgetBean": {"name": "日期",
                                         "dimensions": {"d1": {"name": "日期"}},
                                         "value": date_bean["value"]}},
    }}
    ws = fx.FineBiShareClient._parse_widgets(json.dumps(page, ensure_ascii=False))
    c = fx.FineBiShareClient("http://x", "nQcq")
    c.widgets = ws
    check(c.set_filter("日期", ["2026-11-23"]) == 1, "名字仍然匹配（由调用方决定是否中止）")
    check(c.unsupported_filters and "日期" in c.unsupported_filters[0],
          "不支持的组件被单独记下", str(c.unsupported_filters))
    check(not c.filters, "不支持的组件不会进入 filters（避免静默导全量）")

    w = fx.Widget("w1", "r1", "明细", 1)
    p = c.build_payload({"settings": {"tableStyle": {}}, "name": "明细"}, w)
    check("filterValues" in p and not (p.get("queryInfo") or {}).get("widgetIdValueMap"),
          "载荷里不会塞入不支持的值（第二道防线）")


def t_csv():
    """write_csv：BOM、中文、逗号/引号/换行转义、None 空串。"""
    print("\nCSV 写出")
    import csv as _csv
    import os
    import tempfile
    from xlsx_io import write_csv

    p = os.path.join(tempfile.mkdtemp(), "t.csv")
    header = ["客户号", "备注"]
    rows = [["0001", "普通"], ["0002", "带,逗号"], ["0003", '带"引号'],
            ["0004", "带\n换行"], ["0005", None], [None, ""]]
    write_csv(p, header, rows)
    raw = open(p, "rb").read()
    check(raw[:3] == b"\xef\xbb\xbf", "UTF-8 BOM 已写入（Excel 中文不乱码）")
    with open(p, "r", encoding="utf-8-sig", newline="") as f:
        got = list(_csv.reader(f))
    check(got[0] == header, "表头原样", str(got[0]))
    check(got[1] == ["0001", "普通"], "中文行", str(got[1]))
    check(got[2][1] == "带,逗号", "逗号被引号包裹", str(got[2]))
    check(got[3][1] == '带"引号', "引号字段读写一致（写入时内部引号翻倍）", str(got[3]))
    check(got[4][1] == "带\n换行", "换行保留在单元格内", str(got[4]))
    check(got[5] == ["0005", ""], "None 写成空串", str(got[5]))
    check(got[6] == ["", ""], "空单元格", str(got[6]))
    check(len(got) == 7, "总行数 = 1 表头 + 6 数据", str(len(got)))


def t_job_format():
    """collector.Job：输出格式解析、文件扩展名跟随格式。"""
    print("\n输出格式（csv/xlsx）")
    import os
    import collector as col

    old_env = os.environ.get("FINEBI_OUTPUT_FORMAT")
    os.environ["FINEBI_OUTPUT_FORMAT"] = "csv"          # 固定环境，测试不受机器配置影响
    try:
        def mk(**kw):
            base = dict(row_no=2, name="对账", url="http://x/webroot/decision/link/nQcq",
                        payload="", outdir="out", filename="", cron="* * * * *",
                        remark="", enabled=True)
            base.update(kw)
            return col.Job(**base)

        j = mk()
        check(j.fmt == "csv", "留空默认 csv", j.fmt)
        check(j.ext == ".csv", "扩展名 .csv", j.ext)
        p = j.target_path(widget="客户清单")
        check(p.endswith(".csv") and "对账_客户清单_" in os.path.basename(p),
              "多文件默认名带 .csv", os.path.basename(p))
        check(not j.target_path().endswith(".csv") or True, "单文件路径生成不抛错")

        j2 = mk(fmt="xlsx")
        check(j2.fmt == "xlsx" and j2.ext == ".xlsx", "显式 xlsx", j2.fmt)
        p2 = j2.target_path(widget="客户清单")
        check(p2.endswith(".xlsx"), "xlsx 扩展名", os.path.basename(p2))

        # 手写文件名带 .xlsx 扩展名，输出格式 csv -> 扩展名被替换成 .csv
        j3 = mk(fmt="", filename="对账单.xlsx")
        p3 = j3.target_path(widget="明细")
        check(p3.endswith("对账单_明细.csv"), "手写扩展名跟随输出格式替换",
              os.path.basename(p3))

        # 单文件模式：csv 多组件会在 run_job 降级（这里只验证路径与 split 判定）
        j4 = mk(mode="单文件")
        check(j4.split is False, "单文件判定", str(j4.split))
        check(j4.fmt == "csv", "单文件也默认 csv", j4.fmt)

        # 分组列在 csv 下的自动行为由 run_job 处理，这里验证配置解析不受影响
        j5 = mk(groups="填充")
        check(j5.group_mode == "fill", "分组列=填充解析", j5.group_mode)
    finally:
        if old_env is None:
            os.environ.pop("FINEBI_OUTPUT_FORMAT", None)
        else:
            os.environ["FINEBI_OUTPUT_FORMAT"] = old_env


def main():
    t_detail()
    for m in ("merge", "fill", "raw"):
        t_group(m)
    t_group_blank()
    t_crosstab()
    t_expected()
    t_filters()
    t_xlsx()
    t_61_group_tree()
    t_61_cross_window()
    t_61_flat_fallback()
    t_filters_61()
    t_filter_value_shape()
    t_csv()
    t_job_format()
    print("\n" + "=" * 60)
    print("通过 %d 项，失败 %d 项" % (ok, len(bad)))
    for b in bad:
        print("  - %s" % b)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
