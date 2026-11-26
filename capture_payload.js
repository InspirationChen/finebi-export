#!/usr/bin/env node
/**
 * 抓一次 FineBI 分享页发出的取数请求体（payload）。
 *
 * 用途：页面上的筛选条件（过滤组件、URL 参数）是前端 JS 算好后塞进请求体的，
 * 格式没法靠猜。这个脚本用 headless Chrome 打开链接，把 /widget/data 的
 * 请求体原样抓下来，供 Python 端复用。
 *
 * 用法：
 *   node capture_payload.js "<分享链接，可带 ?参数=值>" [输出文件] [等待毫秒]
 *   node capture_payload.js "http://localhost:37799/webroot/decision/link/nQcq?性别=女" cap.json 25000
 */
const { spawn } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const PORT = 9223;
const target = process.argv[2];
const outFile = process.argv[3] || "captured_payloads.json";
const waitMs = parseInt(process.argv[4] || "25000", 10);

if (!target) {
  console.error("用法: node capture_payload.js <url> [out.json] [waitMs]");
  process.exit(2);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitDevtools(maxMs = 15000) {
  const t0 = Date.now();
  while (Date.now() - t0 < maxMs) {
    try {
      const r = await fetch(`http://127.0.0.1:${PORT}/json/version`);
      if (r.ok) return await r.json();
    } catch (_) { /* 还没起来 */ }
    await sleep(300);
  }
  throw new Error("DevTools 端口没起来");
}

function cdp(ws) {
  let id = 0;
  const pending = new Map();
  const listeners = [];
  ws.addEventListener("message", (ev) => {
    const msg = JSON.parse(ev.data);
    if (msg.id && pending.has(msg.id)) {
      const { resolve, reject } = pending.get(msg.id);
      pending.delete(msg.id);
      msg.error ? reject(new Error(JSON.stringify(msg.error))) : resolve(msg.result);
    } else if (msg.method) {
      for (const fn of listeners) fn(msg);
    }
  });
  return {
    send(method, params = {}) {
      const mid = ++id;
      return new Promise((resolve, reject) => {
        pending.set(mid, { resolve, reject });
        ws.send(JSON.stringify({ id: mid, method, params }));
      });
    },
    on(fn) { listeners.push(fn); },
    close() { try { ws.close(); } catch (_) {} },
  };
}

(async () => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "finebi-cdp-"));
  const chrome = spawn(CHROME, [
    "--headless=new",
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${profile}`,
    "--no-first-run", "--no-default-browser-check",
    "--disable-gpu", "--disable-extensions", "--no-sandbox",
    "--window-size=1600,1000",
    "about:blank",
  ], { stdio: "ignore" });

  const captured = [];
  const seenUrls = new Set();
  let client;

  try {
    await waitDevtools();
    const res = await fetch(`http://127.0.0.1:${PORT}/json/new?url=about:blank`, { method: "PUT" });
    const page = await res.json();

    const ws = new WebSocket(page.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
      ws.addEventListener("open", resolve, { once: true });
      ws.addEventListener("error", reject, { once: true });
    });
    client = cdp(ws);

    client.on((msg) => {
      if (msg.method !== "Network.requestWillBeSent") return;
      const req = msg.params.request;
      if (!/\/widget\/data|\/widgets\/data/.test(req.url)) return;
      if (req.method !== "POST" || !req.postData) return;
      if (seenUrls.has(req.url)) return;   // 同一组件只留第一次（= 页面首次取数）
      seenUrls.add(req.url);
      let body = null, parsed = true;
      try { body = JSON.parse(req.postData); } catch (_) { parsed = false; }
      captured.push({
        url: req.url,
        widget: (body && (body.name || body.title)) || "",
        type: (body && body.type) || null,
        parsed,
        payload: parsed ? body : req.postData,
      });
      console.error(`  [抓到] ${(body && (body.title || body.name)) || "?"}  type=${body && body.type}`);
    });

    await client.send("Network.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: target });

    console.error(`等待 ${waitMs}ms 让页面完成取数…`);
    await sleep(waitMs);
  } catch (e) {
    console.error("抓取失败:", e.message);
  } finally {
    try { client && client.close(); } catch (_) {}
    chrome.kill("SIGTERM");
    await sleep(400);
    try { fs.rmSync(profile, { recursive: true, force: true }); } catch (_) {}
  }

  fs.writeFileSync(outFile, JSON.stringify(captured, null, 1), "utf8");
  console.error(`\n共抓到 ${captured.length} 个取数请求体 -> ${outFile}`);
  process.exit(captured.length ? 0 : 1);
})();
