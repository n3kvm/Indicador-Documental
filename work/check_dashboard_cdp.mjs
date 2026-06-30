import http from "node:http";
import { pathToFileURL } from "node:url";

const dashboardPath = process.argv[2] || "C:/Users/DELL/Documents/Codex/2026-06-03/files-mentioned-by-the-user-05/outputs/dashboard_soportes_mayo_2026.html";

function requestJson(url, method = "GET") {
  return new Promise((resolve, reject) => {
    const req = http.request(url, { method }, (res) => {
      let data = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          resolve(JSON.parse(data));
        } catch (err) {
          reject(new Error(data.slice(0, 500)));
        }
      });
    });
    req.on("error", reject);
    req.end();
  });
}

async function send(ws, method, params = {}) {
  const id = ++send.nextId;
  ws.send(JSON.stringify({ id, method, params }));
  return await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error(`Timeout waiting for ${method}`)), 30000);
    function onMessage(event) {
      const raw = event.data ?? event;
      const msg = JSON.parse(typeof raw === "string" ? raw : raw.toString());
      if (msg.id === id) {
        clearTimeout(timeout);
        ws.removeEventListener("message", onMessage);
        if (msg.error) reject(new Error(JSON.stringify(msg.error)));
        else resolve(msg.result);
      }
    }
    ws.addEventListener("message", onMessage);
  });
}
send.nextId = 0;

const url = pathToFileURL(dashboardPath).href;
const tab = await requestJson(`http://127.0.0.1:9222/json/new?${encodeURIComponent(url)}`, "PUT");
const ws = new WebSocket(tab.webSocketDebuggerUrl);
await new Promise((resolve, reject) => {
  ws.addEventListener("open", resolve, { once: true });
  ws.addEventListener("error", reject, { once: true });
});
await send(ws, "Runtime.enable");
await new Promise((resolve) => setTimeout(resolve, 1500));
const result = await send(ws, "Runtime.evaluate", {
  expression: `({
    title: document.title,
    files: document.getElementById("kpiFiles")?.textContent || document.getElementById("kFiles")?.textContent,
    sites: document.getElementById("kpiSites")?.textContent || document.querySelectorAll("#rowsSites tr").length,
    typed: document.getElementById("kpiTyped")?.textContent || document.getElementById("kChecks")?.textContent,
    missing: document.getElementById("kpiMissing")?.textContent || document.getElementById("kMissing")?.textContent,
    hours: document.getElementById("kHours")?.textContent,
    rows: document.querySelectorAll("#siteRows tr").length || document.querySelectorAll("#rowsSites tr").length
  })`,
  returnByValue: true,
});
console.log(JSON.stringify(result.result.value, null, 2));
ws.close();
