import fs from "node:fs/promises";
import http from "node:http";
import https from "node:https";

const siteOrigin = "https://comfandisa.sharepoint.com";
const sitePath = "/teams/Z3-MantenimientoLocativo";
const anio = process.env.SHAREPOINT_ANIO || "2026";
const carpetaMes = process.env.SHAREPOINT_CARPETA_MES || "05-MAYO";
const baseFolder = `/teams/Z3-MantenimientoLocativo/Documentos compartidos/4. Mantenimiento Preventivo/FACILITY BRILLASEO/01-CONTROL MTTO PREVE-CRON 2022, 2023, 2024, 2025/${anio}/${carpetaMes}`;
const folderUrl = `${baseFolder}/SOPORTE DE VISITAS`;
const outPath = "work/sharepoint_files.json";

function getJsonHttp(url) {
  return new Promise((resolve, reject) => {
    http.get(url, (res) => {
      let data = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          resolve(JSON.parse(data));
        } catch {
          reject(new Error(`Invalid JSON from ${url}: ${data.slice(0, 500)}`));
        }
      });
    }).on("error", reject);
  });
}

function getJsonHttps(url, headers) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers }, (res) => {
      let data = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`HTTP ${res.statusCode} ${res.statusMessage}: ${data.slice(0, 1000)}`));
          return;
        }
        try {
          resolve(JSON.parse(data));
        } catch {
          reject(new Error(`Invalid JSON from ${url}: ${data.slice(0, 500)}`));
        }
      });
    }).on("error", reject);
  });
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
  if (!page) throw new Error("No hay pestañas controlables en Edge. Abre Edge desde el script o usa -NoAbrirEdge solo si ya está abierto con puerto 9222.");

  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  await send(ws, "Network.enable");
  const result = await send(ws, "Network.getCookies", { urls: [siteOrigin] });
  ws.close();

  const cookies = result.cookies || [];
  const cookieHeader = cookies.map((c) => `${c.name}=${c.value}`).join("; ");
  if (!cookieHeader) {
    throw new Error("No encontré cookies de SharePoint en Edge. Inicia sesión en SharePoint en la ventana abierta por el script y vuelve a ejecutar.");
  }
  return cookieHeader;
}

async function readFolder(path, headers, depth = 0) {
  const base = `${siteOrigin}${sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)`;
  const filesUrl = `${base}/Files?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length&$top=5000&@u=${quoteFolder(path)}`;
  const foldersUrl = `${base}/Folders?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified&$top=5000&@u=${quoteFolder(path)}`;

  const files = (await getJsonHttps(filesUrl, headers)).value || [];
  const folders = ((await getJsonHttps(foldersUrl, headers)).value || []).filter((f) => f.Name !== "Forms");

  const nested = [];
  if (depth < 8) {
    for (const folder of folders) {
      const child = await readFolder(folder.ServerRelativeUrl, headers, depth + 1);
      nested.push(...child.files);
    }
  }
  return {
    folders,
    files: [...files.map((f) => ({ ...f, ParentFolder: path })), ...nested],
  };
}

const cookieHeader = await getCookiesFromEdge();
const headers = {
  Cookie: cookieHeader,
  Accept: "application/json;odata=nometadata",
};

const root = await readFolder(folderUrl, headers);
const output = {
  ok: true,
  count: root.files.length,
  rootFolders: root.folders,
  files: root.files,
};

await fs.writeFile(outPath, JSON.stringify(output, null, 2), "utf8");
console.log(JSON.stringify({
  ok: output.ok,
  count: output.count,
  rootFolders: output.rootFolders.map((f) => f.Name),
}, null, 2));
