#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
部署自检：在目标机器上第一件事就跑这个，逐项确认环境是否可用。

    python check_env.py                    # 只查本地环境
    python check_env.py --net              # 额外测 FineBI 连通性
    python check_env.py --config config.xlsx --net

退出码：0 = 全部通过，1 = 有阻塞项。
"""

import argparse
import os
import platform
import socket
import sys
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/158.0.0.0 Safari/537.36"

OK, WARN, BAD = "[ OK ]", "[WARN]", "[FAIL]"
problems = []


def say(level, msg, hint=""):
    """打一行检查结果；BAD 会被记进 problems，决定最终退出码。"""
    print("%s %s" % (level, msg))
    if hint:
        print("        %s" % hint)
    if level == BAD:
        problems.append(msg)


def check_python():
    """
    Python 版本 / 解释器路径 / 控制台编码。

    分档处理：< 3.7 判阻塞（语法都过不去）；< 3.11 只告警（能跑但没实测）；
    另外专门识别 Microsoft Store 版 Python——它在计划任务里常常找不到解释器。
    """
    v = sys.version_info
    print("=" * 66)
    print("Python %d.%d.%d  |  %s  |  %s" %
          (v.major, v.minor, v.micro, platform.system(), sys.executable))
    print("=" * 66)
    # 目标环境是 3.11（Windows）；3.7 以上语法都能跑，但只在 3.11 / 3.13 上实测过
    if v < (3, 7):
        say(BAD, "Python 版本过低：%s，本工具需要 3.7+，实测环境为 3.11"
            % sys.version.split()[0],
            "请升级到 Python 3.11")
    elif v < (3, 11):
        say(WARN, "Python %s 低于实测版本 3.11" % sys.version.split()[0],
            "语法上可运行，但未做完整验证；建议统一到 3.11")
    else:
        say(OK, "Python 版本 %s（兼容）" % sys.version.split()[0])

    # Windows 上 Microsoft Store 版 python 会有这个路径特征，容易踩坑
    if platform.system() == "Windows" and "WindowsApps" in sys.executable:
        say(WARN, "当前是 Microsoft Store 版 Python",
            "计划任务/服务里调用它容易找不到，建议装官方安装包或写绝对路径")

    # 中文输出能力
    enc = getattr(sys.stdout, "encoding", "") or ""
    try:
        "中文测试".encode(enc)
        say(OK, "控制台编码 %s，可以输出中文" % enc)
    except Exception:
        say(WARN, "控制台编码 %s 可能输不出中文" % enc,
             "真控制台下走 Windows Unicode API 不受影响；若通过重定向看日志仍乱码，把重定向文件用记事本转码查看")


def check_stdlib():
    """确认要用的标准库都在——内网装不了包，所以这工具必须只用标准库。"""
    missing = []
    for m in ("urllib.request", "zipfile", "json", "re", "logging.handlers",
              "xml.etree.ElementTree", "decimal"):
        try:
            __import__(m)
        except Exception:
            missing.append(m)
    if missing:
        say(BAD, "标准库缺失：%s" % ", ".join(missing))
    else:
        say(OK, "标准库齐备，无需 pip install 任何包")


def check_files():
    """程序文件是否齐——最常见的事故是只拷了 .py 没拷全，或拷到了错误的目录。"""
    need = ["collector.py", "finebi_export.py", "table_norm.py", "xlsx_io.py",
            "cron_expr.py", "preflight.py"]
    miss = [f for f in need if not os.path.exists(os.path.join(HERE, f))]
    if miss:
        say(BAD, "缺少程序文件：%s" % ", ".join(miss), "确认整个目录都拷贝过去了")
    else:
        say(OK, "程序文件齐备（%d 个）" % len(need))


def check_config(path):
    """配置表能否解析：必需列在不在、每行任务校验能否通过。返回可用任务列表（失败返回 None）。"""
    full = path if os.path.isabs(path) else os.path.join(HERE, path)
    if not os.path.exists(full):
        say(BAD, "配置文件不存在：%s" % full)
        return None
    try:
        sys.path.insert(0, HERE)
        from xlsx_io import read_table
        header, rows = read_table(full)
        say(OK, "配置可读：%d 列 × %d 行数据" % (len(header), len(rows)))
    except Exception as e:
        say(BAD, "配置读取失败：%s" % e)
        return None

    try:
        from collector import load_config
        jobs = load_config(full)
        good = [j for j in jobs if not j.error]
        say(OK if good else BAD, "解析出 %d 个有效任务" % len(good))
        for j in jobs:
            if j.error:
                say(BAD, "  第 %d 行 %s：%s" % (j.row_no, j.name, j.error))
            else:
                mode = "每组件一个文件" if j.split else "合并多 sheet"
                print("        - %-18s %-16s %-14s %s  表样式=%s 分组列=%s"
                      % (j.name, j.cron_raw, mode, "启用" if j.enabled else "停用",
                         j.kind or "自动", j.group_mode))
    except Exception as e:
        say(BAD, "任务解析失败：%s" % e)
    return jobs


def check_writable(jobs):
    """输出目录是否可写。计划任务用普通用户跑，往 Program Files 写必然失败，提前查明比跑到一半报错好。"""
    dirs = set()
    if jobs:
        for j in jobs:
            if not j.error:
                dirs.add(j.outdir)
    if not dirs:
        dirs.add(os.path.join(HERE, "out"))
    for d in sorted(dirs):
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, ".write_test.tmp")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(probe)
            say(OK, "输出目录可写：%s" % d)
        except Exception as e:
            say(BAD, "输出目录不可写：%s（%s）" % (d, e),
                "常见原因：目录被占用、无权限、onedrive 同步锁")


def check_net(jobs):
    """FineBI 站点连通性（--net 才跑）：先测端口，再看登录页能否拉到版本号。"""
    if not jobs:
        return
    seen = set()
    for j in jobs:
        if j.error or not j.base:
            continue
        if j.base in seen:
            continue
        seen.add(j.base)
        host = j.base.split("//")[-1].split("/")[0]
        try:
            hp = host.rsplit(":", 1)
            sock = socket.create_connection((hp[0], int(hp[1]) if len(hp) > 1 else 80), timeout=5)
            sock.close()
            say(OK, "端口可达：%s" % host)
        except Exception as e:
            say(BAD, "连不上 %s：%s" % (host, e), "检查网络/VPN/防火墙")
            continue

        url = "%s/webroot/decision/link/%s" % (j.base, j.code)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=15) as resp:
                html = resp.read().decode("utf-8", "replace")
            if "sessionId" in html:
                say(OK, "分享链接可访问，且能取到会话令牌")
            else:
                say(BAD, "链接返回了页面但拿不到 sessionId", "链接可能已失效，或需要登录后才能访问")
        except urllib.error.HTTPError as e:
            say(BAD, "链接返回 HTTP %s" % e.code, "404=链接失效；401/403=需要鉴权")
        except Exception as e:
            say(BAD, "访问链接失败：%s" % e)


def main():
    """命令行入口：本地项必查，--net 才测连通性。退出码 0 = 无阻塞项，1 = 有。"""
    ap = argparse.ArgumentParser(description="FineBI 采集器部署自检")
    ap.add_argument("--config", default="config.xlsx", help="配置 xlsx（默认同目录 config.xlsx）")
    ap.add_argument("--net", action="store_true", help="额外测试 FineBI 连通性")
    args = ap.parse_args()

    check_python()
    check_stdlib()
    check_files()
    jobs = check_config(args.config)
    check_writable(jobs)
    if args.net:
        check_net(jobs)
    else:
        say(WARN, "未测试网络连通性（加 --net 可测）")

    print()
    if problems:
        print("有 %d 项阻塞问题，见上方 [FAIL]" % len(problems))
        return 1
    print("全部通过，可以开跑：")
    print("    python preflight.py                    # 先验 BI 版本与取数链路（几十秒）")
    print("    python collector.py --list             # 再看任务和下次执行时间")
    print("    python collector.py --once --json      # 立刻跑一轮")
    return 0


if __name__ == "__main__":
    sys.exit(main())
