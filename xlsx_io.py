#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最小 xlsx / csv 读写工具（纯标准库，无需 openpyxl / pandas）

读：支持 sharedStrings、inlineStr、数值、布尔；单 sheet 或多 sheet 按名取
写：首行加粗冻结、列宽自适应、长数字与前导零自动转文本
    支持多级表头（header_cells + merges）与冻结行数自适应，供交叉表/分组表使用
    csv：UTF-8 带 BOM，无行数上限（xlsx 单 sheet 上限 1,048,576 行）

为什么不用 openpyxl：目标机器是内网，装不了第三方包；xlsx 本质是 zip + xml，
这里只实现用得到的那一小部分，400 行以内够用且没有依赖。
"""

import os
import re
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PRNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# 这些列必须按文本写：保留前导零、避免 18 位身份证被 Excel 截断精度
FORCE_TEXT_HINTS = ("身份证", "证件", "客户号", "账号", "帐号", "卡号", "手机", "电话", "号码")


# ==================================================================== 读
def _col_index(ref):
    """A1 -> 0, AB12 -> 27"""
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def _shared_strings(z):
    """读 xl/sharedStrings.xml，按索引返回字符串表。没有这个部件说明全是 inlineStr，返回空表。"""
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall(NS + "si"):
        out.append("".join(t.text or "" for t in si.iter(NS + "t")))
    return out


def _sheet_path(z, sheet_name=None):
    """定位目标 sheet 的部件路径。"""
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return "xl/worksheets/sheet1.xml"

    rel_map = {r.get("Id"): r.get("Target") for r in rels.findall(PRNS + "Relationship")}
    sheets = wb.find(NS + "sheets").findall(NS + "sheet")

    chosen = None
    for s in sheets:
        if sheet_name is None or s.get("name") == sheet_name:
            chosen = s
            break
    if chosen is None:
        raise KeyError("找不到 sheet：%s（可选：%s）"
                       % (sheet_name, [s.get("name") for s in sheets]))

    target = rel_map.get(chosen.get(RNS + "id"), "worksheets/sheet1.xml")
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def read_table(path, sheet_name=None):
    """读一张表 -> (header, rows)，全部元素为字符串。"""
    if not os.path.exists(path):
        raise FileNotFoundError("配置文件不存在：%s" % path)

    with zipfile.ZipFile(path) as z:
        shared = _shared_strings(z)
        part = _sheet_path(z, sheet_name)
        root = ET.fromstring(z.read(part))

    grid = []
    for row in root.iter(NS + "row"):
        cells = {}
        for c in row.findall(NS + "c"):
            ref = c.get("r") or ""
            idx = _col_index(ref) if ref else len(cells)
            t = c.get("t")
            if t == "s":
                v = c.find(NS + "v")
                val = shared[int(v.text)] if v is not None and v.text else ""
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(NS + "t"))
            elif t == "b":
                v = c.find(NS + "v")
                val = "TRUE" if (v is not None and v.text == "1") else "FALSE"
            elif t == "e":
                val = ""
            else:
                v = c.find(NS + "v")
                val = v.text if v is not None and v.text is not None else ""
                # 8.0 -> 8
                if val and re.fullmatch(r"-?\d+\.0+", val):
                    val = val.split(".")[0]
            cells[idx] = (val or "").strip()
        if cells:
            width = max(cells) + 1
            grid.append([cells.get(i, "") for i in range(width)])
        else:
            grid.append([])

    # 去掉尾部空行
    while grid and not any(grid[-1]):
        grid.pop()
    if not grid:
        return [], []
    header = grid[0]
    rows = [r + [""] * (len(header) - len(r)) if len(r) < len(header) else r
            for r in grid[1:]]
    return header, rows


# ==================================================================== 写
def _col_letter(idx):
    """列下标转列字母：0 -> A，27 -> AB（xlsx 的 r 属性要用）。"""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _is_plain_number(text):
    """
    能否安全按「数值」写入。

    两种情况必须按文本：前导零（客户号 0001 会被吃掉）、15 位以上（身份证/账号
    超过 Excel 的 15 位有效数字，后面几位会被改写成 0，是静默的数据错）。
    """
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", str(text)):
        return False
    intpart = str(text).lstrip("-").split(".")[0]
    if len(intpart) > 1 and intpart.startswith("0"):
        return False        # 前导零 -> 文本
    if len(intpart) >= 15:
        return False        # 长数字 -> 文本
    return True


def _cell_xml(ref, col_name, value, style):
    """
    拼一个单元格的 XML。

    写入顺序决定安全：数字 -> 数值单元；文本 -> inlineStr + 文本样式（style），
    并清掉 XML 非法控制字符（否则整本 xlsx 打不开）。
    col_name 参与判断，是因为表头写着「身份证」「客户号」的列一律按文本写。
    """
    if value is None or value == "":
        return '<c r="%s"/>' % ref
    if isinstance(value, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    text = str(value)
    if not any(h in (col_name or "") for h in FORCE_TEXT_HINTS) and _is_plain_number(text):
        return '<c r="%s"><v>%s</v></c>' % (ref, text)
    text = _ILLEGAL.sub("", text)
    return ('<c r="%s" t="inlineStr" s="%d"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, escape(text)))


_BAD_SHEET_CHARS = re.compile(r"[\[\]\:\*\?\/\\]")


def safe_sheet_name(name, used=None, fallback="Sheet"):
    """把任意文本收敛成合法 sheet 名：去非法字符、限 31 字、去重。"""
    s = _BAD_SHEET_CHARS.sub("_", str(name or "")).strip().strip("'") or fallback
    s = s[:31]
    if used is not None:
        base, n = s, 2
        while s.lower() in used:
            suffix = "_%d" % n
            s = base[:31 - len(suffix)] + suffix
            n += 1
        used.add(s.lower())
    return s


# ------------------------------------------------ 文件名（Windows 严格）
_WIN_BAD_CHARS = set('<>:"/\\|?*')
_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL"} | \
                {"COM%d" % i for i in range(1, 10)} | \
                {"LPT%d" % i for i in range(1, 10)}


def safe_filename(name, used=None, fallback="未命名", maxlen=80):
    """
    把组件名收敛成 Windows / macOS 都能安全落盘的文件名主干。
    Windows 比 Unix 严格：额外禁 <>:"/\\|?* 、禁控制字符、结尾不能是空格或点、
    且 CON/PRN/AUX/NUL/COM1-9/LPT1-9 是保留名（哪怕带扩展名也不行）。
    """
    s = "".join("_" if (ch in _WIN_BAD_CHARS or ord(ch) < 32) else ch
                for ch in str(name or ""))
    s = s.strip().rstrip(". ")          # Windows 会静默吃掉结尾的点和空格
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        s = fallback
    if s.split(".")[0].upper() in _WIN_RESERVED:
        s = "_" + s
    s = s[:maxlen]
    if used is not None:
        base, n = s, 2
        while s.lower() in used:
            suffix = "_%d" % n
            s = base[:maxlen - len(suffix)] + suffix
            n += 1
        used.add(s.lower())
    return s


def _sheet_xml(header, rows, header_cells=None, merges=None, header_height=1):
    """
    生成单个 worksheet 的 XML 片段。

    两种表头：
      * header_cells 为空 -> 老行为，header 一行，第 1 行冻结；
      * header_cells 非空 -> 多级表头（含合并），按 header_height 行冻结。
    merges 是 (c1, r1, c2, r2) 的 0-based 列 + 1-based 行；表头与表体的合并都在这里。
    """
    ncol = max(1, len(header) or 0)
    header_cells = header_cells or []
    if header_cells:
        ncol = max(ncol, max(c["col"] + c.get("colspan", 1) for c in header_cells))
        header_height = max(header_height,
                            max(c["row"] + c.get("rowspan", 1) for c in header_cells))
    header_height = max(1, header_height)

    widths = []
    for i in range(ncol):
        longest = len(str(header[i])) if i < len(header) else 8
        for r in rows[:300]:
            if i < len(r) and r[i] is not None:
                longest = max(longest, len(str(r[i])))
        widths.append(min(max(longest * 1.2 + 2, 9), 46))

    cols = "".join('<col min="%d" max="%d" width="%.1f" customWidth="1"/>'
                   % (i + 1, i + 1, w) for i, w in enumerate(widths))

    last_row = header_height + len(rows)
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             '<dimension ref="A1:%s%d"/>' % (_col_letter(ncol - 1), last_row),
             '<sheetViews><sheetView workbookViewId="0">',
             '<pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" state="frozen"/>'
             % (header_height, header_height + 1),
             '</sheetView></sheetViews>',
             '<sheetFormatPr defaultRowHeight="15"/>',
             '<cols>%s</cols>' % cols,
             '<sheetData>']

    if header_cells:
        by_row = {}
        for c in header_cells:
            by_row.setdefault(c["row"], []).append(c)
        for r in range(header_height):
            cells = sorted(by_row.get(r, []), key=lambda x: x["col"])
            xml = "".join(
                _cell_xml("%s%d" % (_col_letter(c["col"]), r + 1), c.get("text", ""),
                          c.get("text", ""), 1) for c in cells)
            parts.append('<row r="%d" ht="20" customHeight="1">%s</row>' % (r + 1, xml))
    else:
        parts.append('<row r="1" ht="20" customHeight="1">%s</row>' % "".join(
            _cell_xml("%s1" % _col_letter(i), header[i] if i < len(header) else "",
                      header[i] if i < len(header) else "", 1) for i in range(ncol)))

    # 被合并覆盖的格子不再写值，否则 Excel 打开会报「修复文档」
    covered = set()
    for (c1, r1, c2, r2) in merges or []:
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (r1, c1):
                    covered.add((rr, cc))

    for ri, row in enumerate(rows, start=header_height + 1):
        cells = []
        for ci in range(ncol):
            if (ri, ci) in covered:
                continue
            name = header[ci] if ci < len(header) else ""
            parts_val = row[ci] if ci < len(row) else None
            cells.append(_cell_xml("%s%d" % (_col_letter(ci), ri), name, parts_val, 0))
        parts.append('<row r="%d">%s</row>' % (ri, "".join(cells)))

    parts.append('</sheetData>')

    if merges:
        refs = "".join('<mergeCell ref="%s%d:%s%d"/>'
                       % (_col_letter(c1), r1, _col_letter(c2), r2)
                       for (c1, r1, c2, r2) in merges)
        parts.append('<mergeCells count="%d">%s</mergeCells>' % (len(merges), refs))

    parts.append('</worksheet>')
    return "".join(parts)


_STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           '<fonts count="2">'
           '<font><sz val="11"/><color theme="1"/><name val="等线"/></font>'
           '<font><b/><sz val="11"/><color theme="1"/><name val="等线"/></font></fonts>'
           '<fills count="3"><fill><patternFill patternType="none"/></fill>'
           '<fill><patternFill patternType="gray125"/></fill>'
           '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF3FA"/><bgColor indexed="64"/>'
           '</patternFill></fill></fills>'
           '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
           '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
           '<cellXfs count="2">'
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
           '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
           '</cellXfs>'
           '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
           '</styleSheet>')

_ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" '
              'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
              'Target="xl/workbook.xml"/></Relationships>')


def write_workbook(path, sheets):
    """
    多 sheet 写出。sheets 每项可以是：
      * (表名, 表头, 数据行) 元组 / 列表 —— 老写法（单行表头）；
      * dict：name / header / rows，另可选 header_cells（多级表头）、
        merges（(c1,r1,c2,r2)，0-based 列 + 1-based 行）、header_height。
    表名会自动清洗（去 []:*?/\\ 、截断 31 字、同名去重）。
    """
    norm = []
    for s in sheets:
        if isinstance(s, dict):
            norm.append({
                "name": s.get("name"),
                "header": s.get("header") or [],
                "rows": s.get("rows") or [],
                "header_cells": s.get("header_cells") or [],
                "merges": s.get("merges") or [],
                "header_height": s.get("header_height") or 1,
            })
        else:
            seq = list(s) + [None, None, None]
            norm.append({"name": seq[0], "header": seq[1] or [], "rows": seq[2] or [],
                         "header_cells": [], "merges": [], "header_height": 1})
    if not norm:
        raise ValueError("至少需要一个 sheet")

    used = set()
    entries = []
    for i, d in enumerate(norm, start=1):
        d["name"] = safe_sheet_name(d["name"] or "Sheet%d" % i, used)
        entries.append(d)
    n = len(entries)

    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>%s</sheets></workbook>'
                % "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(d["name"]), i, i)
                          for i, d in enumerate(entries, start=1)))

    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join('<Relationship Id="rId%d" '
                         'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                         'Target="worksheets/sheet%d.xml"/>' % (i, i)
                         for i in range(1, n + 1))
               + '<Relationship Id="rId%d" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                 'Target="styles.xml"/></Relationships>' % (n + 1))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i
                  for i in range(1, n + 1))
        + '<Override PartName="/xl/styles.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          '</Types>')

    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", _STYLES)
        for i, d in enumerate(entries, start=1):
            z.writestr("xl/worksheets/sheet%d.xml" % i,
                       _sheet_xml(d["header"], d["rows"],
                                  header_cells=d["header_cells"],
                                  merges=d["merges"],
                                  header_height=d["header_height"]))
    return path


def write_xlsx(path, header, rows, sheet_name="Sheet1"):
    """单 sheet 便捷写法（保持旧签名兼容）。"""
    return write_workbook(path, [(sheet_name, header, rows)])


def write_csv(path, header, rows):
    """
    写 UTF-8 CSV（带 BOM，Excel/WPS 双击打开中文不乱码）。

    为什么提供 csv：xlsx 单 sheet 上限 1,048,576 行，明细数据一超就写挂；
    csv 没有行数上限，且任何工具（数据库、Python、Excel）都吃得下。

    注意：csv 没有单元格类型，长数字（身份证/客户号）用 Excel **双击**打开
    仍会被当数值截断精度——这是 Excel 的显示行为，文件本身是完整的文本。
    需要看精度请用「数据 -> 从文本导入」，或该任务改用 xlsx 输出。

    编码用 utf-8-sig：不写 BOM 的话 Excel 会按 ANSI 读，中文全是乱码。
    """
    import csv
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow([str(h) if h is not None else "" for h in header])
            for r in rows:
                w.writerow(["" if v is None else str(v) for v in r])
    except PermissionError:
        raise PermissionError("文件被占用，无法写入：%s（请先关闭正在打开它的 Excel）" % path)
    return path
