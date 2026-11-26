#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格归一化层：把 FineBI 的「明细表 / 分组表 / 交叉表」统一成可写 Excel 的结构。

为什么需要这一层
----------------
三类表在 FineBI 里的**组件模型**和**取数响应**都不一样：

| 类型 | 组件模型特征 | 取数响应特征 | Excel 落法 |
|---|---|---|---|
| 明细表 | `measures` 为空，字段全在 `dimensions` 里按明细展开 | `header` 平铺（dId+text），`items` 每行长度 = 列数 | 单行表头 |
| 分组表 | `measures` 非空；维度做行分组 | `header` 仍平铺，但维度列在视觉上合并 | 维度列纵向合并 |
| 交叉表 | 有列维度（`view['20000']`）或 `header` 带 `children` | `header` 可能是树，叶子才是实际列 | 多级表头 + 合并单元格 |

设计原则
--------
1. **判定不猜死**：类型靠多个信号投票，判不出来就记 `unknown` 并按明细表兜底，
   同时在 `notes` 里写清依据，日志能看见，不会静默出错。
2. **对齐不靠下标**：FineBI 的 cell 自带 `dId`，能对上列就按 `dId` 放，
   对不上才退化成「按位置」。交叉表里的稀疏行（合并区只给首格）靠这个才不会错列。
3. **合并不改数据**：分组列合并/填充只影响呈现，`rows` 里的原始值不动；
   需要「原样」时把 group_mode 设成 raw，导出的就是接口给的东西。
4. **行数可自证**：取数响应里的 `row` 是服务端算出的总行数，导出后拿它比对，
   少了就是截断（totalRows 没生效 / 翻页失效），当场能看出来。
"""

import re

# ------------------------------------------------------------------ 常量
KIND_DETAIL = "明细表"
KIND_GROUP = "分组表"
KIND_CROSSTAB = "交叉表"
KIND_UNKNOWN = "未知"

GROUP_MODES = ("merge", "fill", "raw")          # 合并 / 向下填充 / 原样

# 多级表头里「子节点」可能用的键名，不同版本/不同表类型见过不同写法，都认
_CHILD_KEYS = ("children", "childHeader", "subHeader", "subHeaders", "headers", "child")
_TEXT_KEYS = ("text", "name", "title", "caption", "label", "fieldName")

# ---------------------------------------------------------------- 单元格取值
# BI 会把手机号这类数字列按「千分位」显示（13,871,014,583），
# value 侧则是科学计数（1.3871014583E10）。落 Excel 前要还原成纯数字串。
_THOUSANDS = re.compile(r"^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$")
_SCI = re.compile(r"^-?\d(?:\.\d+)?[eE][+]?\d+$")


def cell_text(cell):
    """从 cell 里取最干净的字符串值。"""
    if not isinstance(cell, dict):
        return "" if cell is None else str(cell)
    # formatText 是服务端算好的显示值（6.1 的树节点只有它），优先用它；
    # 6.0 的明细 cell 没有 formatText 时退回 text/value，行为不变。
    t = cell.get("formatText")
    if t in (None, ""):
        t = cell.get("text")
    v = cell.get("value")
    s = str(t) if t not in (None, "") else ("" if v is None else str(v))
    if _THOUSANDS.match(s):                 # 13,871,014,583 -> 13871014583
        s = s.replace(",", "")
    elif _SCI.match(s) and v is not None:   # 1.3871014583E10 -> 13871014583
        try:
            from decimal import Decimal
            s = format(Decimal(str(v)).to_integral_value(), "f")
        except Exception:
            pass
    return s


# ------------------------------------------------------------------ 表类型
def detect_kind(detail, header=None):
    """
    判定表格类型，返回 (kind, 依据说明)。

    信号（按优先级）：
      1. 没有 settings.tableStyle  -> 不是表格组件（图表），由调用方跳过
      2. 组件模型里的 type 字段（6.1 直接给：1 分组 / 2 交叉 / 4 明细）
      3. 有列维度（view['20000']）或多级表头 -> 交叉表
      4. measures 非空 -> 分组表
      5. 其余 -> 明细表
    注意：第 3、4 条是从 6.0.x 的模型反推的，只在拿不到 type 时兜底；
    判错时用配置列「表样式」或 --kind 强制指定。
    """
    settings = (detail or {}).get("settings") or {}
    if "tableStyle" not in settings:
        return None, "没有 settings.tableStyle，不是表格组件"

    # 6.1 起组件模型自带 type 字段，比任何反推都可靠
    t = (detail or {}).get("type")
    if t in (1, 2, 4) or (isinstance(t, str) and t.isdigit() and int(t) in (1, 2, 4)):
        t = int(t)
        return {4: KIND_DETAIL, 2: KIND_CROSSTAB, 1: KIND_GROUP}[t], "组件模型 type=%d" % t

    view = detail.get("view") or {}
    col_dims = [x for x in (view.get("20000") or []) if x]
    if col_dims:
        return KIND_CROSSTAB, "view['20000'] 有 %d 个列维度" % len(col_dims)

    if _has_nested_header(header):
        return KIND_CROSSTAB, "表头是多级结构（含 children）"

    measures = [m for m in (detail.get("measures") or []) if m]
    if measures:
        return KIND_GROUP, "measures 非空（%d 个指标）" % len(measures)

    return KIND_DETAIL, "无列维度、无指标，按明细表处理"


def _has_nested_header(header):
    """表头里有没有子节点（有就是多级表头，交叉表常见）。键名见过多种写法，见 _CHILD_KEYS。"""
    for h in header or []:
        if isinstance(h, dict):
            for k in _CHILD_KEYS:
                if isinstance(h.get(k), list) and h[k]:
                    return True
    return False


# ------------------------------------------------------------------ 表头
def _node_text(node):
    """取表头节点的显示文本。text 常为 None（分组树的根节点就是这样），按候选键依次退。"""
    for k in _TEXT_KEYS:
        v = node.get(k)
        if isinstance(v, str) and v.strip():
            return v.strip()
    return ""


def _node_children(node):
    """取子节点列表；没有子节点返回 None（叶子节点与「子节点为空」要区分开）。"""
    for k in _CHILD_KEYS:
        v = node.get(k)
        if isinstance(v, list) and v:
            return v
    return None


def flatten_header(header):
    """
    把（可能多级的）表头压成列定义。

    返回 (cols, header_cells)：
      cols         = [{'dId','text','path':[一级,二级...]}]   # 叶子列，顺序 = Excel 列顺序
      header_cells = [{'row','col','rowspan','colspan','text'}]  # 供 Excel 画多级表头
    单级表头时 header_cells 只有一行，行为和老版本一致。
    """
    cols, cells = [], []

    def walk(nodes, level, start, parent_path):
        width = 0
        for node in nodes:
            if not isinstance(node, dict):
                node = {"text": str(node)}
            text = _node_text(node)
            kids = _node_children(node)
            if kids:
                sub = walk(kids, level + 1, start + width, parent_path + [text])
                cells.append({"row": level, "col": start + width, "colspan": sub,
                              "rowspan": 1, "text": text, "leaf": False})
                width += sub
            else:
                cols.append({"dId": str(node.get("dId") or node.get("id") or ""),
                             "text": text, "path": parent_path + [text]})
                cells.append({"row": level, "col": start + width, "colspan": 1,
                              "rowspan": 1, "text": text, "leaf": True})
                width += 1
        return width

    if header:
        walk(header, 0, 0, [])

    if not cells:                       # 空表头兜底，避免下游除零
        return [], [{"row": 0, "col": 0, "colspan": 1, "rowspan": 1, "text": "", "leaf": True}]

    depth = max(c["row"] for c in cells) + 1
    # 跨行规则：
    #   * 有子节点的表头（如「AUM」跨两个月）只占自己这一层，向下由子节点填 -> rowspan 1
    #   * 叶子节点如果不在最底层（如「机构」旁边是多级指标），要纵跨到最后一层
    # 这两条写反了的表现是：合并区错位一行，Excel 里看着「少一行表头」。
    for c in cells:
        c["rowspan"] = 1 if not c["leaf"] else (depth - c["row"])
    return cols, cells


def col_names(cols):
    """叶子列的显示名：单级用 text，多级用「父/子」拼接，便于识别与列宽估算。"""
    out = []
    for c in cols:
        if isinstance(c, str):          # 6.1 解析出来的列名已经是拼好的字符串
            out.append(c)
            continue
        path = [p for p in (c.get("path") or []) if p]
        out.append(path[-1] if len(path) <= 1 else "/".join(path))
    return out


# ------------------------------------------------------------------ 行对齐
def row_values(item, cols, did_map=None):
    """
    把一行 cell 数组对齐到 cols。优先按 cell['dId'] 落位，对不上再按位置。

    为什么不能只按位置：交叉表/分组表里合并区的行只返回首个格，
    后面几列直接缺省，纯按下标会整体左移，越到右边越错。
    """
    n = len(cols)
    if did_map is None:
        did_map = did_index(cols)

    if len(item) == n:                  # 长度刚好，按位置最快最稳
        return [cell_text(c) for c in item]

    vals = [""] * n
    used = {}
    for pos, cell in enumerate(item):
        did = cell.get("dId") if isinstance(cell, dict) else None
        idx = None
        if did and did in did_map:
            loc = did_map[did]
            k = used.get(did, 0)
            idx = loc[k] if k < len(loc) else loc[-1]     # 同名列（交叉表重复指标）顺序消费
            used[did] = k + 1
        elif pos < n:
            idx = pos
        if idx is not None:
            vals[idx] = cell_text(cell)
    return vals


def did_index(cols):
    """
    dId -> 列下标列表。

    一个 dId 对应多列是正常的（交叉表里同一指标按列维度展开成多列），
    所以值是 list 而不是 int；行对齐时按顺序取用（见 row_values）。
    """
    m = {}
    for i, c in enumerate(cols):
        d = c.get("dId")
        if d:
            m.setdefault(d, []).append(i)
    return m


# ------------------------------------------------------------------ 维度列
def measure_ids(detail):
    """组件模型里所有「指标」字段的 id（measures / widgetMeasures 两种写法都收）。"""
    out = set()
    for key in ("measures", "widgetMeasures"):
        for m in (detail or {}).get(key) or []:
            if isinstance(m, dict):
                for k in ("id", "fieldId", "dId", "widgetFieldId"):
                    if m.get(k):
                        out.add(str(m[k]))
            elif isinstance(m, (str, int)):
                out.add(str(m))
    return out


def dimension_cols(cols, detail):
    """
    找出「行维度」列（合并/填充只作用在这些列上）。
    只取开头连续的一段：分组表的维度总在指标左边，交叉表更复杂，
    保守一点宁可少合并，也不要错把指标列合并掉。
    """
    mids = measure_ids(detail)
    out = []
    for i, c in enumerate(cols):
        d = c.get("dId") or ""
        if d and d in mids:
            break
        out.append(i)
    return out or ([0] if cols else [])


def _runs(rows, col):
    """连续相同的非空值区间 -> [(起始行, 结束行)]，用于纵向合并。"""
    spans, i, n = [], 0, len(rows)
    while i < n:
        v = rows[i][col] if col < len(rows[i]) else ""
        if v == "":
            i += 1
            continue
        j = i + 1
        while j < n and col < len(rows[j]) and rows[j][col] == v:
            j += 1
        if j - i >= 2:
            spans.append((i, j - 1))
        i = j
    return spans


def apply_group_mode(rows, dim_cols, mode, header_height):
    """
    处理分组列的呈现方式，返回 (rows, merges)。
      merge：只在 Excel 里合并单元格，数据不动（最接近 FineBI 页面）
      fill ：把同值向下填满（做透视/对数更友好，但会改变单元格内容）
      raw  ：完全原样（接口给什么就什么）
    """
    if mode not in GROUP_MODES or not dim_cols or not rows:
        return rows, []

    if mode == "raw":
        return rows, []

    if mode == "fill":
        rows = [list(r) for r in rows]
        for c in dim_cols:
            last = ""
            for r in rows:
                if c < len(r):
                    if r[c] == "":
                        r[c] = last
                    else:
                        last = r[c]
        return rows, []

    merges = []
    for c in dim_cols:
        for a, b in _runs(rows, c):
            # 行号：表头占 header_height 行，数据从 header_height+1 开始（1-based）
            merges.append((c, a + header_height + 1, c, b + header_height + 1))
    return rows, merges


# ------------------------------------------------------------------ 筛选摘要
def describe_filters(detail):
    """
    把组件模型里跟「筛选/过滤/参数」有关的配置摘出来，供日志核对。
    不在这里做解释——FineBI 各版本的过滤器结构差异很大，硬解释容易说错，
    所以只报告「哪些键有值、几条」，原始片段留给 --dump 出来的报文。
    """
    out = []

    def scan(obj, path="", depth=0):
        if not isinstance(obj, dict) or depth > 2:
            return
        for k, v in obj.items():
            low = str(k).lower()
            if any(t in low for t in ("filter", "condition", "parameter", "injection")):
                if v in (None, "", [], {}):
                    continue
                n = len(v) if isinstance(v, (list, dict)) else 1
                out.append("%s%s（%d 项）" % (path, k, n))
            elif isinstance(v, dict):
                scan(v, path + str(k) + ".", depth + 1)

    scan(detail)
    # 组件级参数（日期控件等）单独点名，这类不显式带上会取到错区间
    params = (detail or {}).get("widgetParameters") or (detail or {}).get("parameters")
    if params:
        out.append("组件参数 %d 个" % (len(params) if hasattr(params, "__len__") else 1))
    return out


# ------------------------------------------------------------------ 成品
class Table(object):
    """一个可写 Excel 的表格。"""

    def __init__(self, name, kind, cols, header_cells, rows, merges=None,
                 notes=None, expected_total=None, filter_summary=None):
        """
        cols          列定义（含 dId，行对齐靠它）
        header_cells  表头单元格（带 row/col/colspan/rowspan），多级表头就靠它落位
        expected_total 服务端报的总行数，用于自证「没被截断」
        filter_summary 生效的筛选描述，写进注释行便于回溯这份数据是什么口径
        """
        self.name = name
        self.kind = kind
        self.cols = cols
        self.header_cells = header_cells
        self.rows = rows
        self.merges = merges or []
        self.notes = notes or []
        self.expected_total = expected_total
        self.filter_summary = filter_summary or []

    @property
    def header(self):
        """扁平列名（日志、列宽、兼容旧接口用）。"""
        return col_names(self.cols)

    @property
    def header_height(self):
        """表头占几行（写 Excel 时用来决定冻结几行）。"""
        return max([c["row"] + c.get("rowspan", 1) for c in self.header_cells] or [1])

    def sheet(self, name=None):
        """转成 write_workbook 认识的 sheet dict。"""
        return {
            "name": name or self.name,
            "header": self.header,
            "header_cells": self.header_cells,
            "rows": self.rows,
            "merges": self.merges,
            "header_height": self.header_height,
        }

    def __repr__(self):
        """日志里一眼看清是哪张表、什么类型、多大。"""
        return "<Table %r %s %dx%d>" % (self.name, self.kind, len(self.rows), len(self.cols))


def build_table(name, detail, header, rows, expected_total=None,
                kind_override=None, group_mode="merge"):
    """
    组装 Table：判类型 -> 扁平表头 -> 处理分组列 -> 记录告警。

    header 是接口原样的表头数组；rows 应已用 row_values() 对齐过列。
    """
    kind, why = detect_kind(detail, header)
    if kind is None:
        kind = KIND_UNKNOWN
    if kind_override:                      # 人工强制指定，覆盖自动判定
        notes_kind = "自动判定为 %s（%s），已按配置强制为 %s" % (kind, why, kind_override)
        kind = kind_override
    else:
        notes_kind = "判定为 %s（%s）" % (kind, why)

    cols, cells = flatten_header(header)
    notes = [notes_kind]
    header_height = max([c["row"] + c.get("rowspan", 1) for c in cells] or [1])

    # 行数自证：接口给了总行数，和实际拿到的对比，少了就是截断
    if isinstance(expected_total, int) and expected_total >= 0 and rows is not None:
        if expected_total != len(rows):
            notes.append("行数不符：服务端报 %d 行，实际拿到 %d 行"
                         % (expected_total, len(rows)))

    if kind in (KIND_GROUP, KIND_CROSSTAB):
        dims = dimension_cols(cols, detail)
        rows, body_merges = apply_group_mode(rows, dims, group_mode, header_height)
        if group_mode == "fill" and dims:
            notes.append("分组列已向下填充（%d 列）" % len(dims))
        elif group_mode == "merge" and body_merges:
            notes.append("分组列已合并单元格（%d 处）" % len(body_merges))
    else:
        body_merges = []

    # 表头自身的合并：跨列/跨行的都要写进 mergeCells
    head_merges = []
    for c in cells:
        cs, rs = c.get("colspan", 1), c.get("rowspan", 1)
        if cs > 1 or rs > 1:
            r1, c1 = c["row"] + 1, c["col"]
            head_merges.append((c1, r1, c1 + cs - 1, r1 + rs - 1))

    return Table(name, kind, cols, cells, rows,
                 merges=head_merges + body_merges,
                 notes=notes, expected_total=expected_total,
                 filter_summary=describe_filters(detail))


# ---------------------------------------------------------------- 6.1 响应结构
# 6.1 起取数响应按表型分化出三种形态（6.0 只有第一种）：
#   明细表 type=4  header[] + items[[cell, cell...]]            平铺
#   分组表 type=1  header[] + items[{text, children[], values[]}] 维度是嵌套树
#   交叉表 type=2  top{header,items} + left{header,items}        视窗式，值在 left 叶子上
# 三者都要能落成同一张 Excel，所以统一收敛到 Table。

def _tree_leaves(nodes, path=()):
    """
    把嵌套节点树摊成 [(路径元组, 叶子节点)]，顺序即 Excel 里的列/行顺序。

    根节点的 text 通常是 None（它是整张表的合计行，不是维度值），
    这种空节点不能进路径，否则会多出一个空维度层。
    """
    out = []
    for n in nodes or []:
        if not isinstance(n, dict):
            continue
        txt = n.get("text")
        p = path if txt in (None, "") else path + (txt,)
        kids = n.get("children") or []
        if kids:
            out.extend(_tree_leaves(kids, p))
        else:
            out.append((p, n))
    return out


def _tree_depth(nodes, d=1):
    """树有几层（决定表头高度，也用于算跨行合并的 rowspan）。"""
    best = d
    for n in nodes or []:
        kids = (n or {}).get("children") or []
        if kids:
            best = max(best, _tree_depth(kids, d + 1))
    return best


def _header_merges(cells):
    """表头里跨行/跨列的格子要写成 mergeCells。"""
    out = []
    for c in cells:
        cs, rs = c.get("colspan", 1), c.get("rowspan", 1)
        if cs > 1 or rs > 1:
            r1, c1 = c["row"] + 1, c.get("col", 0)
            out.append((c1, r1, c1 + cs - 1, r1 + rs - 1))
    return out


def _flat_header_cells(cols):
    """单行表头：每个列名一格，rowspan/colspan 都是 1（明细表、6.0 平铺响应走这里）。"""
    return [{"row": 0, "col": i, "text": c, "colspan": 1, "rowspan": 1, "leaf": True}
            for i, c in enumerate(cols)]


def is_tree_rows(items):
    """items 的元素是 dict（带 children/values）而不是 cell 列表 -> 6.1 树形行。"""
    return bool(items) and isinstance(items[0], dict)


def is_cross_response(data):
    """6.1 交叉表的响应把行列头拆成 top / left 两块。"""
    return isinstance(data, dict) and ("top" in data or "left" in data)


def from_group_tree(name, detail, data, group_mode="merge"):
    """
    解析 6.1 分组表：items 是一棵维度树，叶子节点的 values 才是指标值。
    根节点的 text 为 None，代表「合计」行（6.1 不给标签，这里补一个）。
    """
    hdr = [(h or {}).get("text", "") for h in (data.get("header") or [])]
    items = data.get("items") or []
    # 维度层数按叶子路径算（根节点是合计行，不算一层）
    depth = max([len(p) for p, _ in _tree_leaves(items)] or [1]) or 1
    notes = ["按 6.1 分组树解析（维度 %d 层）" % depth]

    rows, has_total = [], False

    def walk(node, path):
        nonlocal has_total
        txt = node.get("text")
        cur = list(path) if txt in (None, "") else list(path) + [txt]
        vals = node.get("values") or []
        if vals:                             # 有值就落一行（合计行也落，与 BI 页面一致）
            pad = cur + [""] * max(0, depth - len(cur))
            if not cur:                      # 根合计行：接口不给标签，补一个
                pad[0] = "合计"
                has_total = True
            rows.append(pad[:depth] + [cell_text(v) for v in vals])
        for kid in (node.get("children") or []):
            walk(kid, cur)

    for n in items:
        walk(n, [])

    nmeas = max([len(r) - depth for r in rows] or [0])
    if len(hdr) >= depth + nmeas and nmeas:
        cols = hdr[:depth + nmeas]
    else:
        cols = (hdr[:depth] or ["维度%d" % (i + 1) for i in range(depth)])
        cols += (hdr[depth:depth + nmeas] or ["值%d" % (i + 1) for i in range(nmeas)])

    if has_total:
        notes.append("含合计行（接口未给标签，已补「合计」）")
    cells = _flat_header_cells(cols)
    rows2, body_merges = apply_group_mode(rows, list(range(depth)), group_mode, 1)
    if group_mode == "merge" and body_merges:
        notes.append("分组列已合并单元格（%d 处）" % len(body_merges))
    elif group_mode == "fill" and depth:
        notes.append("分组列已向下填充（%d 列）" % depth)

    return Table(name, KIND_GROUP, cols, cells, rows2,
                 merges=_header_merges(cells) + body_merges,
                 notes=notes, filter_summary=describe_filters(detail))


def from_cross_window(name, detail, data, group_mode="merge"):
    """
    解析 6.1 交叉表：top 是列头树、left 是行头树，
    行叶子的 values 按「列叶子顺序」一一对应，长度应等于列叶子数。
    """
    top = data.get("top") or {}
    left = data.get("left") or {}
    col_leaves = _tree_leaves(top.get("items") or [])
    row_leaves = _tree_leaves(left.get("items") or [])
    notes = ["按 6.1 交叉表视窗解析（列叶子 %d，行叶子 %d）"
             % (len(col_leaves), len(row_leaves))]

    if not col_leaves or not row_leaves:
        notes.append("列头或行头为空，可能是视窗未展开（调大 colCounts/rowCounts 后重试）")

    lhdr = [(h or {}).get("text", "") for h in (left.get("header") or [])]
    col_depth = max([len(p) for p, _ in col_leaves] or [1])
    row_depth = max([len(p) for p, _ in row_leaves] or [1])
    # left.header 的长度 = 行维度数 + 列叶子数，用它能反推行维度数
    row_dims = len(lhdr) - len(col_leaves) if len(lhdr) > len(col_leaves) else row_depth
    row_dims = max(1, min(row_dims, row_depth)) if row_depth else 1
    dim_names = (lhdr[:row_dims] or ["维度%d" % (i + 1) for i in range(row_dims)])

    # 列名：叶子路径拼接，重名时带上层级避免重复
    cols = dim_names[:]
    for p, node in col_leaves:
        cols.append(" / ".join([x for x in p if x]) or "值")

    header_rows = max(col_depth, row_dims, 1)
    cells = []
    for i in range(row_dims):
        r = min(i, header_rows - 1)
        span = (header_rows - r) if i == row_dims - 1 else 1
        cells.append({"row": r, "col": i, "text": dim_names[i],
                      "colspan": 1, "rowspan": span, "leaf": False})

    # 列头：逐层把「连续且同名」的列叶子合成一个跨列单元格
    for lvl in range(col_depth):
        i = 0
        while i < len(col_leaves):
            j = i + 1
            while (j < len(col_leaves)
                   and col_leaves[j][0][lvl:lvl + 1] == col_leaves[i][0][lvl:lvl + 1]):
                j += 1
            name_txt = col_leaves[i][0][lvl] if lvl < len(col_leaves[i][0]) else ""
            last = (lvl == col_depth - 1)
            cells.append({"row": lvl, "col": row_dims + i,
                          "text": name_txt or "",
                          "colspan": j - i,
                          "rowspan": (header_rows - lvl) if last else 1,
                          "leaf": last})
            i = j

    ncol_val = len(col_leaves)
    rows = []
    for p, node in row_leaves:
        head = list(p) + [""] * max(0, row_dims - len(p))
        vals = [cell_text(v) for v in (node.get("values") or [])]
        vals += [""] * (ncol_val - len(vals))
        rows.append(head[:row_dims] + vals)

    if data.get("hasXNext") or data.get("hasYNext"):
        notes.append("服务端标记还有更多列/行（hasXNext/hasYNext），当前只是可视区域")

    rows2, body_merges = apply_group_mode(rows, list(range(row_dims)), group_mode, header_rows)
    if group_mode == "merge" and body_merges:
        notes.append("行维度已合并单元格（%d 处）" % len(body_merges))

    return Table(name, KIND_CROSSTAB, cols, cells, rows2,
                 merges=_header_merges(cells) + body_merges,
                 notes=notes, filter_summary=describe_filters(detail))


def from_response(name, detail, data, group_mode="merge", kind_override=None):
    """
    统一入口：把取数响应落成 Table（兼容 6.0 的平铺结构与 6.1 的三种形态）。
    data 必须是单次请求的 data 部分；明细表的分页由调用方负责拼好。
    """
    if is_cross_response(data):
        t = from_cross_window(name, detail, data, group_mode)
    elif is_tree_rows(data.get("items")):
        t = from_group_tree(name, detail, data, group_mode)
    else:
        t = None
    if t is None:
        return None
    if kind_override:
        t.notes.insert(0, "自动识别为 %s，已按配置强制为 %s" % (t.kind, kind_override))
        t.kind = kind_override
    return t
