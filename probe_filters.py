#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
过滤组件结构探针：把每个过滤组件的值结构原样打出来。

为什么需要它：文本类过滤组件的值就是字符串数组，直接写没问题；日期类、数值区间类
的值是对象（日期是 {year, month, day, ...}），按文本写会被服务端反序列化拒掉，
报的是一串 HttpMessageNotReadableException，看起来像接口坏了。要支持这些组件，
得先知道它们真实的结构——就是这个脚本的输出。

用法：
    python probe_filters.py --base http://10.0.0.5:37799 --code nQcq
    python probe_filters.py --config config.xlsx          # 取配置里第一个启用任务
    python probe_filters.py --base ... --code ... --out f.json

注意：链接里不要带 ?参数，否则会被当成筛选值，脚本会直接报「不支持」而拿不到结构。
"""

import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import finebi_export as fx   # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="打印 FineBI 过滤组件的值结构")
    ap.add_argument("--base", help="FineBI 站点根，如 http://10.0.0.5:37799")
    ap.add_argument("--code", help="分享链接 code（link/<code>），不要带 ? 参数")
    ap.add_argument("--config", default="config.xlsx", help="配置 xlsx（用于取第一个启用任务）")
    ap.add_argument("--out", help="把结果另存为 JSON，便于回传")
    args = ap.parse_args()

    base, code = args.base, args.code
    if not (base and code):
        path = args.config if os.path.isabs(args.config) else os.path.join(HERE, args.config)
        from collector import load_config
        jobs = [j for j in load_config(path) if not j.error and j.enabled]
        if not jobs:
            print("配置里没有可用任务，请用 --base/--code 指定")
            return 2
        base, code = jobs[0].base, jobs[0].code
        print("用配置任务「%s」：%s  code=%s" % (jobs[0].name, base, code))

    # 参数会把 bootstrap 里的筛选解析带起来，这里只想要结构，先剥掉
    code = code.split("?")[0]

    client = fx.FineBiShareClient(base, code).bootstrap()
    fws = client.filter_widgets()
    print("\n页面「%s」共 %d 个过滤组件" % (client.report_name or client.report_id, len(fws)))
    if not fws:
        print("没有过滤组件——这个仪表板的筛选不在组件层面，URL 传参无从生效")
        return 0

    dumped = []
    for i, w in enumerate(fws, 1):
        bean = w.bean or {}
        kind = fx.filter_kind(bean)
        support = "可直接传文本" if fx.shape_filter_value(bean, ["x"]) is not None else "对象型值，需适配"
        print("\n%d) %s" % (i, w.title or "(未命名)"))
        print("   字段      : %s" % (w.field or "?"))
        print("   widgetId  : %s" % w.widget_id)
        print("   值结构    : %s（%s）" % (kind, support))
        value = bean.get("value")
        text = json.dumps(value, ensure_ascii=False, indent=2) if value is not None else "null（未选值）"
        for line in text.splitlines():
            print("   " + line)
        dumped.append({"title": w.title, "field": w.field, "widgetId": w.widget_id,
                       "kind": kind, "label": copy_label(bean), "value": value})

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(dumped, f, ensure_ascii=False, indent=2)
        print("\n已写入 %s" % args.out)

    bad = [d["title"] for d in dumped if d["kind"] != "文本类"]
    if bad:
        print("\n以下组件的值不是字符串数组，链接参数里传它们会失败：%s" % "、".join(bad))
        print("把上面这几段 value 结构发出来，即可适配。")
    return 0


def copy_label(bean):
    """bean 里能标识组件类型的线索（类名/type 数字），报错时定位用。"""
    cur = (bean or {}).get("value")
    if isinstance(cur, dict):
        return {"type": cur.get("type"), "keys": sorted(cur.keys())}
    return None


if __name__ == "__main__":
    sys.exit(main())
