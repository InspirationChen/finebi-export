#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cron 表达式解析与匹配（纯标准库）

支持：
  标准 5 段   ：分 时 日 月 周         例  30 8 * * *
  Quartz 6/7 段：秒 分 时 日 月 周 (年) 例  0 30 8 * * ?      —— 秒位只允许 0 或 *
  宏          ：@hourly @daily @midnight @weekly @monthly @yearly @annually
  语法        ：*  */n  a-b  a-b/n  a,b,c  （周：0/7=周日，1=周一）

注意：调度粒度是「分钟」，秒位只接受 0 或 *。
"""

import datetime as _dt
import re
from datetime import date, datetime, timedelta

MACROS = {
    "@yearly": "0 0 1 1 *",
    "@annually": "0 0 1 1 *",
    "@monthly": "0 0 1 * *",
    "@weekly": "0 0 * * 0",
    "@daily": "0 0 * * *",
    "@midnight": "0 0 * * *",
    "@hourly": "0 * * * *",
}

_RANGES = {"minute": (0, 59), "hour": (0, 23), "dom": (1, 31),
           "month": (1, 12), "dow": (0, 7)}      # dow 的 7 视作周日

# Quartz 风格的名字写法
DOW_NAMES = {"SUN": 0, "MON": 1, "TUE": 2, "WED": 3, "THU": 4, "FRI": 5, "SAT": 6}
MONTH_NAMES = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
               "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}


class CronError(ValueError):
    """cron 写法错误。消息里带上出错那一段，配置写错时能直接定位到改哪格。"""
    pass


def _subst_names(text, names):
    """把 MON / JAN 这类缩写换成数字。"""
    def rep(m):
        word = m.group(0).upper()
        if word not in names:
            raise CronError("无法识别的名称：%r（可用：%s）"
                            % (m.group(0), "/".join(sorted(names))))
        return str(names[word])
    return re.sub(r"[A-Za-z]{3}", rep, text)


def _parse_field(expr, name, lo, hi):
    """解析单个字段 -> set"""
    expr = expr.strip()
    if expr in ("", "*", "?"):
        return set(range(lo, hi + 1))

    out = set()
    for piece in expr.split(","):
        piece = piece.strip()
        if not piece:
            raise CronError("%s 字段有空项：%r" % (name, expr))
        step = 1
        if "/" in piece:
            piece, s = piece.split("/", 1)
            if not s.isdigit() or int(s) <= 0:
                raise CronError("%s 字段步长非法：%r" % (name, expr))
            step = int(s)
        if piece in ("*", "?"):
            start, end = lo, hi
        elif "-" in piece:
            a, b = piece.split("-", 1)
            if not (a.isdigit() and b.isdigit()):
                raise CronError("%s 字段区间非法：%r" % (name, expr))
            start, end = int(a), int(b)
        elif piece.isdigit():
            start = end = int(piece)
        else:
            raise CronError("%s 字段无法解析：%r" % (name, expr))

        if start < lo or end > hi or start > end:
            raise CronError("%s 字段越界（%d-%d）：%r" % (name, lo, hi, expr))
        out.update(range(start, end + 1, step))
    return out


class CronExpr:
    """一条 cron 表达式。"""

    __slots__ = ("raw", "minute", "hour", "dom", "month", "dow",
                 "_dom_star", "_dow_star", "_hours", "_minutes")

    def __init__(self, expr):
        """
        解析表达式，允许三种写法：标准 5 段、Quartz 6/7 段、@宏。

        ? 一律当 * 处理（Quartz 里两者都是「不限制」）；秒位只认 0 或 *，
        其它值直接报错而不是降级——静默降级会让用户以为精确到秒跑了，
        实际只在整分钟跑，很难查。
        """
        self.raw = (expr or "").strip()
        if not self.raw:
            raise CronError("cron 表达式为空")

        text = MACROS.get(self.raw.lower(), self.raw).replace("?", "*")
        parts = text.split()
        if len(parts) == 7:            # 去掉年份字段
            parts = parts[:6]
        if len(parts) == 6:            # Quartz 秒位
            if parts[0] not in ("0", "*"):
                raise CronError("仅支持分钟级调度，秒位必须是 0 或 *：%r" % self.raw)
            parts = parts[1:]
        if len(parts) != 5:
            raise CronError("cron 需要 5 段（分 时 日 月 周），实际 %d 段：%r"
                            % (len(parts), self.raw))

        self.minute = _parse_field(parts[0], "分钟", 0, 59)
        self.hour = _parse_field(parts[1], "小时", 0, 23)
        self.dom = _parse_field(parts[2], "日", 1, 31)
        self.month = _parse_field(_subst_names(parts[3], MONTH_NAMES), "月", 1, 12)
        dow = _parse_field(_subst_names(parts[4], DOW_NAMES), "周", 0, 7)
        if 7 in dow:
            dow.discard(7)
            dow.add(0)
        self.dow = dow

        self._dom_star = parts[2].strip() in ("*", "?")
        self._dow_star = parts[4].strip() in ("*", "?")
        self._hours = sorted(self.hour)
        self._minutes = sorted(self.minute)

    # ------------------------------------------------------------ 匹配
    def _day_ok(self, d):
        """某一天是否命中。日与周同时被限定时按「或」（标准 cron 语义，与 Quartz 一致）。"""
        if d.month not in self.month:
            return False
        dom_ok = d.day in self.dom
        dow_ok = (d.isoweekday() % 7) in self.dow     # 周一=1 … 周日=0
        # 标准 cron 语义：日与周都被限定时取「或」
        if not self._dom_star and not self._dow_star:
            return dom_ok or dow_ok
        return dom_ok and dow_ok

    def matches(self, when):
        """这个时刻是否命中表达式（粒度到分钟，秒/毫秒忽略）。"""
        return (when.minute in self.minute
                and when.hour in self.hour
                and self._day_ok(when.date()))

    def next_after(self, when=None, horizon_days=400):
        """下一次触发时间（不含 when 本身），找不到返回 None。"""
        when = (when or datetime.now()).replace(second=0, microsecond=0)
        d = when.date()
        for _ in range(horizon_days):
            if self._day_ok(d):
                for h in self._hours:
                    if h < when.hour and d == when.date():
                        continue
                    for m in self._minutes:
                        cand = datetime(d.year, d.month, d.day, h, m)
                        if cand > when:
                            return cand
            d += timedelta(days=1)
        return None

    def describe(self):
        """人话描述。目前就返回原文（日志里够用），留个口子方便以后改成「每天 8:30」。"""
        return self.raw

    def __repr__(self):
        """调试用；日志里正常走 describe()。"""
        return "CronExpr(%r)" % self.raw


def parse(expr):
    """便捷函数：字符串 -> CronExpr（等价于 CronExpr(expr)）。"""
    return CronExpr(expr)
