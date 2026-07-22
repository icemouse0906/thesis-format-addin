import { spawn } from "node:child_process";
import http from "node:http";
import https from "node:https";
import fs from "node:fs";

const CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
// 使用真实 Word 常见的 $ 分隔格式
const URL =
  "https://localhost:3000/taskpane/index.html#_host_Info=Word$Mac$16.00$zh-CN$$$";

const userData = `/tmp/thesis-chrome-profile-${Date.now()}`;
fs.mkdirSync(userData, { recursive: true });

const chrome = spawn(
  CHROME,
  [
    "--headless=new",
    "--disable-gpu",
    "--ignore-certificate-errors",
    "--remote-debugging-port=9223",
    `--user-data-dir=${userData}`,
    URL,
  ],
  { stdio: "ignore" }
);

function get(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith("https") ? https : http;
    lib
      .get(url, { rejectUnauthorized: false }, (res) => {
        let data = "";
        res.on("data", (c) => (data += c));
        res.on("end", () => resolve({ status: res.statusCode, data }));
      })
      .on("error", reject);
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  await sleep(2500);
  const { default: WebSocket } = await import("ws");
  const list = await get("http://127.0.0.1:9223/json/list");
  const pages = JSON.parse(list.data);
  const page = pages.find((p) => p.url.includes("localhost:3000")) || pages[0];
  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((r, j) => {
    ws.once("open", r);
    ws.once("error", j);
  });

  let id = 0;
  const send = (method, params = {}) => {
    const thisId = ++id;
    return new Promise((resolve, reject) => {
      const onMsg = (buf) => {
        const msg = JSON.parse(buf.toString());
        if (msg.id === thisId) {
          ws.off("message", onMsg);
          if (msg.error) reject(new Error(JSON.stringify(msg.error)));
          else resolve(msg.result);
        }
      };
      ws.on("message", onMsg);
      ws.send(JSON.stringify({ id: thisId, method, params }));
    });
  };

  const logs = [];
  const failed = [];
  ws.on("message", (buf) => {
    const msg = JSON.parse(buf.toString());
    if (msg.method === "Network.responseReceived") {
      const u = msg.params.response.url;
      const s = msg.params.response.status;
      if (u.includes("office-js") || u.includes("word-")) {
        logs.push(`${s} ${u.replace("https://localhost:3000", "")}`);
      }
      if (s >= 400) failed.push(`${s} ${u}`);
    }
    if (msg.method === "Runtime.consoleAPICalled") {
      const t = msg.params.args?.map((a) => a.value ?? a.description).join(" ");
      if (t) logs.push(`console:${t}`.slice(0, 200));
    }
  });

  await send("Network.enable");
  await send("Runtime.enable");
  await send("Page.reload", { ignoreCache: true });
  await sleep(6000);

  const result = await send("Runtime.evaluate", {
    expression: `(() => JSON.stringify({
      hasOffice: !!window.Office,
      hasWord: !!window.Word,
      hasWordRun: !!(window.Word && window.Word.run),
      scripts: [...document.scripts].map(s => s.src).filter(s => s.includes('office') || s.includes('word')).map(s => s.replace(location.origin,'')),
      diag: document.querySelector('.diag')?.textContent || '',
      hostInfo: (sessionStorage.getItem('hostInfoValue')||'').slice(0,80)
    }))()`,
    returnByValue: true,
  });

  console.log("EVAL", result.result.value);
  console.log("NET", JSON.stringify(logs.slice(0, 30), null, 2));
  console.log("FAIL", failed.slice(0, 20));
  ws.close();
  chrome.kill("SIGKILL");
}

main().catch((e) => {
  console.error(e);
  try {
    chrome.kill("SIGKILL");
  } catch {}
  process.exit(1);
});
