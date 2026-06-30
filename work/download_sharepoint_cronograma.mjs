import fs from "node:fs/promises";
import http from "node:http";
import https from "node:https";

const siteOrigin = "https://comfandisa.sharepoint.com";
const sitePath = "/teams/Z3-MantenimientoLocativo";
const anio = process.env.SHAREPOINT_ANIO || "2026";
const carpetaMes = process.env.SHAREPOINT_CARPETA_MES || "05-MAYO";
const mesNombre = process.env.SHAREPOINT_MES_NOMBRE || "MAYO";
const folderUrl = `/teams/Z3-MantenimientoLocativo/Documentos compartidos/4. Mantenimiento Preventivo/FACILITY BRILLASEO/01-CONTROL MTTO PREVE-CRON 2022, 2023, 2024, 2025/${anio}/${carpetaMes}`;
const outPath = "work/cronograma_sharepoint.xlsx";
const metaPath = "work/cronograma_sharepoint_meta.json";

function getJsonHttp(url) {
  return new Promise((resolve, reject) => {
    http.get(url, (res) => {
      let data = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error(`Invalid JSON from ${url}: ${data.slice(0, 500)}`)); }
      });
    }).on("error", reject);
  });
}

function requestBuffer(url, headers) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode)) {
        res.resume();
        requestBuffer(new URL(res.headers.location, url).href, headers).then(resolve, reject);
        return;
      }
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        const body = Buffer.concat(chunks);
        if (res.statusCode < 200 || res.statusCode >= 300) {
          if (res.statusCode === 401 || res.statusCode === 403) {
            reject(new Error(`HTTP ${res.statusCode}: SharePoint no autorizo la consulta del cronograma. Inicia sesion en la ventana de Edge que abre el script con una cuenta que tenga acceso a ${siteOrigin}${folderUrl}. Detalle: ${body.toString("utf8", 0, 500)}`));
            return;
          }
          reject(new Error(`HTTP ${res.statusCode}: ${body.toString("utf8", 0, 1000)}`));
          return;
        }
        resolve(body);
      });
    }).on("error", reject);
  });
}

function getJsonHttps(url, headers) {
  return requestBuffer(url, headers).then((buf) => JSON.parse(buf.toString("utf8")));
}

async function send(ws, method, params = {}) {
  const id = ++send.nextId;
  ws.send(JSON.stringify({ id, method, params }));
  return await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error(`Timeout waiting for ${method}`)), 60000);
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

function quoteFolder(path) {
  return `'${encodeURIComponent(path.replace(/'/g, "''"))}'`;
}

async function getCookiesFromEdge() {
  const tabs = await getJsonHttp("http://127.0.0.1:9222/json");
  const page = tabs.find((t) => t.type === "page" && t.webSocketDebuggerUrl);
  if (!page) throw new Error("No hay pestañas controlables en Edge.");
  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  await send(ws, "Network.enable");
  const result = await send(ws, "Network.getCookies", { urls: [siteOrigin] });
  ws.close();
  const cookieHeader = (result.cookies || []).map((c) => `${c.name}=${c.value}`).join("; ");
  if (!cookieHeader) throw new Error("No encontré cookies de SharePoint en Edge.");
  return cookieHeader;
}

const cookieHeader = await getCookiesFromEdge();
const headers = { Cookie: cookieHeader, Accept: "application/json;odata=nometadata" };
const filesUrl = `${siteOrigin}${sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)/Files?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length&$top=5000&@u=${quoteFolder(folderUrl)}`;
const files = (await getJsonHttps(filesUrl, headers)).value || [];

const candidates = files
  .filter((f) => /\.xlsx$/i.test(f.Name))
  .filter((f) => {
    const name = f.Name.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    return name.includes("cronograma general mtto prog") && name.includes(mesNombre.toLowerCase()) && name.includes(String(anio));
  })
  .sort((a, b) => new Date(b.TimeLastModified) - new Date(a.TimeLastModified));

if (!candidates.length) {
  throw new Error(`No encontré un Excel de cronograma en SharePoint. Archivos Excel disponibles: ${files.filter(f => /\.xlsx$/i.test(f.Name)).map(f => f.Name).join(" | ")}`);
}

const selected = candidates[0];
const downloadUrl = siteOrigin + encodeURI(selected.ServerRelativeUrl);
const bytes = await requestBuffer(downloadUrl, { Cookie: cookieHeader, Accept: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*" });
await fs.writeFile(outPath, bytes);
await fs.writeFile(metaPath, JSON.stringify({ selected, downloadedBytes: bytes.length, outPath }, null, 2), "utf8");

console.log(JSON.stringify({
  ok: true,
  name: selected.Name,
  modified: selected.TimeLastModified,
  bytes: bytes.length,
  outPath,
}, null, 2));
