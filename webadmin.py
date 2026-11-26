#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FineBI 采集管理页面：浏览器里改配置、点按钮跑任务、看日志、下载产出。

启动（在项目目录）：
    python webadmin.py                 # 默认 http://127.0.0.1:8787
    python webadmin.py --port 9000     # 换端口
    python webadmin.py --host 0.0.0.0  # 允许局域网其他机器访问（默认只允许本机）

只依赖标准库 + 项目内既有模块，不需要额外安装任何东西。

接口一览（前端 admin.html 调用，也可用 curl）：
    GET  /api/jobs     读 config.xlsx 里的任务列表
    POST /api/save     保存任务列表（整表覆盖写回 config.xlsx）
    POST /api/run      跑任务：{"job": "任务名"} 单跑，{} 全部（启用任务）
    GET  /api/status   当前运行状态 + 实时日志 + 结束后的结果汇总
    GET  /api/files    out/ 下按日期文件夹列出产出
    GET  /download?f=  下载某个产出文件（仅限 out/ 目录内）
    GET  /api/log      logs/collector.log 末尾（常驻模式的日志也能看）
    POST /api/check    跑 check_env.py 部署自检

设计要点：
  * 跑任务用 subprocess 起 collector.py --once --json，进程隔离：页面崩了/关了
    不影响采集本身；结束后的 JSON 汇总从输出里解析出来给前端展示
  * 同一时刻只允许一个采集进程（FineBI 并发许可有限，叠跑只会互相挤占）
  * /download 做了目录穿越防护，只允许 out/ 下的文件
  * **内置 cron 自动调度**：管理台开着时，后台线程每 20 秒检查一次 config 里的
    cron 表达式，到点自动跑对应任务（与手动跑共用同一进程槽，不会叠跑）。
    所以「开着 webadmin.bat = cron 生效」，不用再单独跑 collector.py 常驻模式。
    两者只能留一个，避免双重触发；实在要单跑常驻模式，设环境变量
    FINEBI_ADMIN_NO_SCHED=1 可关掉这里的调度器
"""

import argparse
import datetime as _dt
import json
import os
import re
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import collector
from xlsx_io import read_table, write_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ADMIN_HTML = os.path.join(HERE, "admin.html")

# 保存配置时的列顺序（新表按这个建；读旧表时按表头匹配，列序无关）
CONFIG_COLUMNS = ["任务名", "启用", "分享链接", "组件", "输出方式", "输出格式", "输出目录",
                  "文件名", "cron表达式", "备注", "表样式", "分组列"]
# Job 属性 -> 配置列名的对应关系（读写 config.xlsx 都走它）
JOB_FIELDS = [
    ("name", "任务名"), ("enabled", "启用"), ("url", "分享链接"),
    ("widgets", "组件"), ("mode", "输出方式"), ("fmt_raw", "输出格式"),
    ("outdir", "输出目录"),
    ("filename", "文件名"), ("cron_raw", "cron表达式"), ("remark", "备注"),
    ("kind_raw", "表样式"), ("groups_raw", "分组列"),
]


class RunState(object):
    """一次采集运行的全部状态。同一时刻最多一个进程，state 加锁保护。"""

    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None            # 当前 subprocess.Popen
        self.kind = ""              # "run" / "check"
        self.label = ""             # 展示用说明
        self.started = None
        self.finished = None
        self.exitcode = None
        self.lines = []             # 输出按行存（含时间戳由子进程日志自带）
        self.report = None          # 结束后从 --json 输出解析出的结果汇总

    def snapshot(self, tail=500):
        with self.lock:
            return {
                "running": self.proc is not None and self.proc.poll() is None,
                "kind": self.kind, "label": self.label,
                "started": self.started, "finished": self.finished,
                "exitcode": self.exitcode,
                "lines": self.lines[-tail:],
                "report": self.report,
                "scheduler": _sched_info(),
            }


STATE = RunState()
CONFIG_PATH = os.path.join(HERE, "config.xlsx")


def _win_no_window():
    """Windows 下隐藏子进程控制台窗口；其他平台返回 0（无副作用）。"""
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _spawn(args):
    """
    起子进程并把 stdout/stderr 按行收进 STATE.lines，结束后解析 --json 汇总。
    返回线程对象（调用方不必 join，进程结束时会自动收尾）。
    """
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"       # Windows 控制台默认 GBK，强制 UTF-8
    env["PYTHONUNBUFFERED"] = "1"

    def worker():
        try:
            proc = subprocess.Popen(
                args, cwd=HERE, env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
                creationflags=_win_no_window())
        except Exception as e:
            with STATE.lock:
                STATE.lines.append("[启动失败] %s" % e)
                STATE.exitcode = -1
                STATE.finished = _now()
            return
        with STATE.lock:
            STATE.proc = proc
        for line in proc.stdout:
            with STATE.lock:
                STATE.lines.append(line.rstrip("\n"))
        proc.wait()
        with STATE.lock:
            STATE.exitcode = proc.returncode
            STATE.finished = _now()
            STATE.report = _parse_json_report()

    th = threading.Thread(target=worker, daemon=True)
    th.start()
    return th


def _parse_json_report():
    """从输出末尾解析 collector --json 的汇总（JSON 在输出最后，indent 多行）。"""
    text = "\n".join(STATE.lines)
    i = text.rfind("\n{")
    if i < 0:
        return None
    try:
        return json.loads(text[i + 1:])
    except Exception:
        return None


def _now():
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _job_dict(j):
    """Job 对象 -> 前端要的 dict（显示用信息 + 编辑用原始值）。"""
    return {
        "name": j.name, "enabled": j.enabled, "url": j.url,
        "widgets": j.widgets_raw, "mode": j.mode, "fmt": j.fmt_raw, "outdir": j.outdir,
        "filename": j.filename or "", "cron": j.cron_raw, "remark": j.remark or "",
        "kind": j.kind_raw, "groups": j.groups_raw,
        "nextRun": j.next_run().strftime("%Y-%m-%d %H:%M") if j.next_run() else "",
        "error": j.error or "",
    }


def api_jobs():
    """读配置。load_config 会打日志，这里把根 logger 压到 WARNING 免得混进页面输出。"""
    import logging
    logging.getLogger().setLevel(logging.WARNING)
    jobs = collector.load_config(CONFIG_PATH)
    if isinstance(jobs, tuple):
        jobs = jobs[0]
    return {"ok": True, "columns": CONFIG_COLUMNS,
            "jobs": [_job_dict(j) for j in jobs]}


def api_save(rows):
    """
    整表写回 config.xlsx。rows 是前端表格的每行 dict。
    Job 上的 *_raw 属性是「用户在配置里写的原文」（未解析），保存时以原文为准。
    """
    if not isinstance(rows, list):
        raise ValueError("body 应是 {rows: [...]}")
    header = list(CONFIG_COLUMNS)
    out_rows = []
    for r in rows:
        vals = {col: "" for col in header}
        vals["任务名"] = str(r.get("name") or "").strip()
        vals["启用"] = "是" if r.get("enabled") in (True, "true", "True", "是", 1) else \
                       ("" if r.get("enabled") in ("", None) else "否")
        vals["分享链接"] = str(r.get("url") or "").strip()
        vals["组件"] = str(r.get("widgets") or "").strip()
        vals["输出方式"] = str(r.get("mode") or "").strip()
        fmt = str(r.get("fmt") or "").strip().lower()
        if fmt in ("xlsx", "excel", "xls"):
            fmt = "xlsx"
        elif fmt in ("", "csv"):
            fmt = ""
        else:
            fmt = ""
        vals["输出格式"] = fmt
        vals["输出目录"] = str(r.get("outdir") or "out").strip()
        vals["文件名"] = str(r.get("filename") or "").strip()
        vals["cron表达式"] = str(r.get("cron") or "").strip()
        vals["备注"] = str(r.get("remark") or "").strip()
        vals["表样式"] = str(r.get("kind") or "").strip()
        vals["分组列"] = str(r.get("groups") or "").strip()
        out_rows.append([vals[c] for c in header])
    if not out_rows:
        raise ValueError("至少要保留一行任务")
    write_workbook(CONFIG_PATH, [("任务", header, out_rows)])
    return api_jobs()


def api_files():
    """out/ 下按文件夹分组列产出。文件夹名就是日期（20261125），新的排前面。"""
    root = os.path.join(HERE, "out")
    groups = {}
    if os.path.isdir(root):
        for dirpath, _dirs, files in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            if rel == ".":
                rel = ""
            for fn in files:
                if fn.startswith("."):
                    continue
                p = os.path.join(dirpath, fn)
                st = os.stat(p)
                groups.setdefault(rel, []).append({
                    "name": fn, "path": os.path.relpath(p, root),
                    "size": st.st_size,
                    "mtime": _dt.datetime.fromtimestamp(st.st_mtime).strftime("%m-%d %H:%M"),
                })
    out = [{"folder": k, "files": sorted(v, key=lambda x: x["name"])}
           for k, v in sorted(groups.items(), key=lambda kv: kv[0], reverse=True)]
    return {"ok": True, "groups": out}


def api_download(qs):
    """下载 out/ 下的文件。路径拼出来必须在 out/ 里，否则拒绝（防目录穿越）。"""
    rel = qs.get("f", [""])[0].lstrip("/")
    root = os.path.abspath(os.path.join(HERE, "out"))
    full = os.path.abspath(os.path.join(root, rel))
    if not full.startswith(root + os.sep) or not os.path.isfile(full):
        raise PermissionError("只能下载 out/ 目录下的文件")
    ctype = "application/octet-stream"
    if full.endswith(".html"):
        ctype = "text/html; charset=utf-8"
    elif full.endswith(".csv"):
        ctype = "text/csv; charset=utf-8"
    with open(full, "rb") as f:
        data = f.read()
    return 200, ctype, data, os.path.basename(full)


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):    # 静默默认的访问日志
        pass

    # ------------------------------------------------------------ 工具
    def _send(self, code, ctype, body, extra_headers=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra_headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, obj, code=200):
        self._send(code, "application/json; charset=utf-8", json.dumps(obj, ensure_ascii=False))

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return {}
        return json.loads(self.rfile.read(n).decode("utf-8"))

    # ------------------------------------------------------------ 路由
    def do_GET(self):
        path, _, qs = self.path.partition("?")
        from urllib.parse import parse_qs, unquote
        qs = parse_qs(qs)
        path = unquote(path)
        try:
            if path == "/" or path == "/index.html":
                with open(ADMIN_HTML, "rb") as f:
                    self._send(200, "text/html; charset=utf-8", f.read())
            elif path == "/api/jobs":
                self._json(api_jobs())
            elif path == "/api/status":
                self._json({"ok": True, **STATE.snapshot()})
            elif path == "/api/files":
                self._json(api_files())
            elif path == "/api/log":
                self._json(api_log())
            elif path == "/api/check":
                self._json(api_check())        # GET 也放行：前端 api() 无参数时发的是 GET
            elif path == "/api/stop":
                self._json(api_stop())         # 同上，终止按钮的兜底
            elif path == "/download":
                code, ctype, data, name = api_download(qs)
                self._send(code, ctype, data,
                           {"Content-Disposition":
                            "attachment; filename*=UTF-8''%s"
                            % _quote_filename(name)})
            else:
                self._json({"ok": False, "error": "未知路径 %s" % path}, 404)
        except Exception as e:
            self._json({"ok": False, "error": str(e)}, 500)

    def do_POST(self):
        from urllib.parse import unquote
        path = unquote(self.path)
        try:
            if path == "/api/save":
                self._json(api_save(self._body().get("rows") or []))
            elif path == "/api/run":
                self._json(api_run(self._body()))
            elif path == "/api/stop":
                self._json(api_stop())
            elif path == "/api/check":
                self._json(api_check())
            else:
                self._json({"ok": False, "error": "未知路径 %s" % path}, 404)
        except Exception as e:
            self._json({"ok": False, "error": str(e)}, 500)


import urllib.parse

def _quote_filename(s):
    """下载头里的中文文件名按 RFC 5987 编码。"""
    return urllib.parse.quote(s)


def api_run(body):
    """跑任务。body.job 为空 = 全部启用任务；否则单跑同名任务。"""
    with STATE.lock:
        if STATE.proc is not None and STATE.proc.poll() is None:
            return {"ok": False, "error": "已有任务在跑（%s），等它结束再试" % STATE.label}
        args = [sys.executable, os.path.join(HERE, "collector.py"),
                "--config", CONFIG_PATH, "--once", "--json"]
        name = str(body.get("job") or "").strip()
        if name:
            args += ["--job", name]
        if body.get("verbose"):
            args.append("--verbose")
        if body.get("dryRun"):
            args.append("--dry-run")
        STATE.kind = "run"
        STATE.label = name or "全部启用任务"
        STATE.started, STATE.finished = _now(), None
        STATE.exitcode, STATE.report = None, None
        STATE.lines = ["$ python collector.py --once %s" %
                       ("--job " + name if name else "").strip()]
    _spawn(args)
    return {"ok": True, "label": STATE.label}


def api_check():
    """部署自检（check_env.py），复用同一套进程状态。"""
    with STATE.lock:
        if STATE.proc is not None and STATE.proc.poll() is None:
            return {"ok": False, "error": "已有任务在跑（%s）" % STATE.label}
        STATE.kind, STATE.label = "check", "环境自检"
        STATE.started, STATE.finished = _now(), None
        STATE.exitcode, STATE.report = None, None
        STATE.lines = ["$ python check_env.py"]
    _spawn([sys.executable, os.path.join(HERE, "check_env.py")])
    return {"ok": True}


def api_stop():
    with STATE.lock:
        if STATE.proc is None or STATE.proc.poll() is not None:
            return {"ok": False, "error": "当前没有在跑的任务"}
        STATE.proc.kill()
        return {"ok": True, "note": "已终止（正在写一半的文件可能不完整）"}


def api_log():
    """常驻模式日志末尾 200 行。"""
    p = os.path.join(HERE, "logs", "collector.log")
    lines = []
    if os.path.isfile(p):
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()[-200:]
    return {"ok": True, "lines": [ln.rstrip("\n") for ln in lines]}


# ============================================================
# cron 自动调度：管理台开着时，到点自动按 config 里的 cron 跑任务
# ============================================================
SCHED_TICK = 20   # 秒。cron 的粒度是「分钟」，每 20 秒醒一次足够准

SCHED = {
    "lock": threading.Lock(),
    "queue": [],      # 到点待跑的任务名（一次只放一个进程，剩下的排队）
    "fired": set(),   # 已触发过的 (任务名, "2026-11-25 16:45")，防同一分钟重复跑
    "recent": [],     # 最近几次自动触发记录（页面展示用）
    "started": False,
    "stop": False,
}


def _sched_info():
    """给 /api/status 的调度器概况（队列 + 最近自动触发记录）。"""
    with SCHED["lock"]:
        return {"on": SCHED["started"], "queue": list(SCHED["queue"]),
                "recent": list(SCHED["recent"])}


def _sched_tick():
    """找出这一分钟到点的启用任务排进队列。同一分钟同一任务只排一次。"""
    now = _dt.datetime.now().replace(second=0, microsecond=0)
    stamp = now.strftime("%Y-%m-%d %H:%M")
    import logging
    logging.getLogger().setLevel(logging.WARNING)   # load_config 的日志别混进页面输出
    try:
        jobs = collector.load_config(CONFIG_PATH)
        if isinstance(jobs, tuple):
            jobs = jobs[0]
    except Exception:
        return
    with SCHED["lock"]:
        for j in jobs:
            if not j.enabled or not j.cron_raw:
                continue
            key = (j.name, stamp)
            if key in SCHED["fired"]:
                continue
            try:
                due = collector.CronExpr(j.cron_raw).matches(now)
            except Exception:
                continue      # cron 写错不在此处报：load_config 已写进 j.error，页面可见
            if due:
                SCHED["fired"].add(key)
                SCHED["queue"].append(j.name)
        if len(SCHED["fired"]) > 3000:            # 防膨胀：只保留最近的触发记录
            SCHED["fired"] = set(sorted(SCHED["fired"])[-1500:])


def _auto_run(job_name):
    """到点跑一个任务（与手动跑共用同一进程槽）。返回 False = 进程槽被占。"""
    with STATE.lock:
        if STATE.proc is not None and STATE.proc.poll() is None:
            return False
        args = [sys.executable, os.path.join(HERE, "collector.py"),
                "--config", CONFIG_PATH, "--once", "--json", "--job", job_name]
        STATE.kind = "run"
        STATE.label = "%s（cron 自动）" % job_name
        STATE.started, STATE.finished = _now(), None
        STATE.exitcode, STATE.report = None, None
        STATE.lines = ["$ python collector.py --once --job %s   [cron 自动触发]"
                       % job_name]
    _spawn(args)
    return True


def _wait_idle():
    """等当前采集进程结束（与手动跑互斥）。

    注意不能只看 proc.poll()：进程一退出 poll() 就返回，但 worker 线程
    还要把 exitcode/report 写进状态——只等 poll 会读到 None 的退出码。
    所以额外等「这次运行还没有 finished 标记」为止。
    """
    while not SCHED["stop"]:
        with STATE.lock:
            alive = STATE.proc is not None and STATE.proc.poll() is None
            pending = STATE.started is not None and STATE.finished is None
        if not (alive or pending):
            return
        time.sleep(1)


def _sched_note(job_name, exitcode):
    """记一笔自动触发结果，页面「调度器」一行能看到最近几次。"""
    with SCHED["lock"]:
        SCHED["recent"].insert(0, {"time": _now(), "job": job_name,
                                   "exitcode": exitcode})
        del SCHED["recent"][8:]


def scheduler_loop():
    """后台调度主循环：每 SCHED_TICK 秒查一次到点任务并逐个执行。"""
    while not SCHED["stop"]:
        try:
            _sched_tick()
            job = None
            with SCHED["lock"]:
                if SCHED["queue"]:
                    job = SCHED["queue"].pop(0)
            if job is not None:
                if _auto_run(job):
                    _wait_idle()
                    with STATE.lock:
                        rc = STATE.exitcode
                    _sched_note(job, rc)
                    print("[%s] cron 自动跑完：%s（退出码 %s）" % (_now(), job, rc))
                else:
                    # 手动任务正占着进程槽：放回队头，下一轮再试
                    with SCHED["lock"]:
                        SCHED["queue"].insert(0, job)
        except Exception as e:
            print("[%s] 调度器异常：%r" % (_now(), e))
        time.sleep(SCHED_TICK)


def main():
    global CONFIG_PATH
    ap = argparse.ArgumentParser(description="FineBI 采集管理页面")
    ap.add_argument("--host", default="127.0.0.1",
                    help="监听地址（默认 127.0.0.1 只允许本机；0.0.0.0 允许局域网）")
    ap.add_argument("--port", type=int,
                    default=int(os.environ.get("FINEBI_ADMIN_PORT", "8787")))
    ap.add_argument("--config", default=CONFIG_PATH, help="配置 Excel 路径")
    ap.add_argument("--no-browser", action="store_true", help="启动后不自动开浏览器")
    args = ap.parse_args()
    CONFIG_PATH = args.config if os.path.isabs(args.config) \
        else os.path.join(HERE, args.config)

    url = "http://%s:%d" % (args.host if args.host != "0.0.0.0" else "127.0.0.1",
                            args.port)
    print("管理页面：%s" % url)
    print("配置文件：%s" % CONFIG_PATH)
    # cron 自动调度：默认开。注意与 collector.py 常驻模式二选一，避免双重触发；
    # 确要单跑常驻模式时设 FINEBI_ADMIN_NO_SCHED=1 关掉这里的调度器。
    if os.environ.get("FINEBI_ADMIN_NO_SCHED") == "1":
        print("cron 自动调度：关（FINEBI_ADMIN_NO_SCHED=1）")
    else:
        SCHED["started"] = True
        threading.Thread(target=scheduler_loop, daemon=True).start()
        print("cron 自动调度：开（每 %d 秒检查一次 cron 表达式，到点自动跑）" % SCHED_TICK)
    print("Ctrl+C 退出")
    if not args.no_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n退出")
        SCHED["stop"] = True
        with STATE.lock:
            if STATE.proc and STATE.proc.poll() is None:
                STATE.proc.kill()


if __name__ == "__main__":
    main()
