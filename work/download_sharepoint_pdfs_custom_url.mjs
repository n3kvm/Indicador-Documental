import fs from "node:fs/promises";
import path from "node:path";
import http from "node:http";
import https from "node:https";
import { URL } from "node:url";

const listPath = process.env.SHAREPOINT_FILES_PATH || "work/sharepoint_files.json";
const manifestPath = process.env.SHAREPOINT_PDF_MANIFEST_PATH || "work/sharepoint_pdf_manifest.json";
const outDir = process.env.SHAREPOINT_PDF_CACHE_DIR || "work/sharepoint_pdfs";
const defaultOrigin = process.env.SHAREPOINT_ORIGIN || "https://comfandisa-my.sharepoint.com";

function safeName(name) {
  return String(name || "archivo.pdf").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim() || "archivo.pdf";
}

async function exists(file) {
  try { await fs.access(file); return true; } catch { return false; }
}

async function readJson(file, fallback) {
  try { return JSON.parse(await fs.readFile(file, "utf8")); } catch { return fallback; }
}

function requestJson(url) {
  return new Promise((resolve, reject) => {
    http.get(url, (res) => {
      let body = "";
      res.setEncoding("utf8");
      res.on("data", (chunk) => body += chunk);
      res.on("end", () => {
        if (res.statusCode < 200 || res.statusCode >= 300) return reject(new Error(`HTTP ${res.statusCode}: ${body.slice(0, 500)}`));
        try { resolve(JSON.parse(body)); } catch (err) { reject(err); }
      });
    }).on("error", reject);
  });
}

function requestBuffer(url, headers = {}, redirects = 0) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const client = parsed.protocol === "http:" ? http : https;
    const req = client.get(parsed, { headers }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location && redirects < 8) {
        res.resume();
        const next = new URL(res.headers.location, parsed).toString();
        return resolve(requestBuffer(next, headers, redirects + 1));
      }
      const chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        const body = Buffer.concat(chunks);
        if (res.statusCode < 200 || res.statusCode >= 300) {
          return reject(new Error(`HTTP ${res.statusCode}: ${body.toString("utf8", 0, 700)}`));
        }
        resolve(body);
      });
    });
    req.on("error", reject);
  });
}

async function getCookieHeader(origin) {
  const version = await requestJson("http://127.0.0.1:9222/json/version");
  if (!version.webSocketDebuggerUrl) throw new Error("Edge no expuso WebSocket de depuracion en 9222.");
  const ws = new WebSocket(version.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  let id = 1;
  function send(method, params) {
    const msgId = id++;
    ws.send(JSON.stringify({ id: msgId, method, params }));
    return new Promise((resolve, reject) => {
      const onMessage = (event) => {
        const data = JSON.parse(event.data.toString());
        if (data.id !== msgId) return;
        ws.removeEventListener("message", onMessage);
        if (data.error) reject(new Error(data.error.message || JSON.stringify(data.error)));
        else resolve(data.result || {});
      };
      ws.addEventListener("message", onMessage);
    });
  }
  const result = await send("Network.getCookies", { urls: [origin] });
  ws.close();
  const cookieHeader = (result.cookies || []).map((c) => `${c.name}=${c.value}`).join("; ");
  if (!cookieHeader) throw new Error("No encontre cookies de SharePoint en Edge. Inicia sesion en la ventana abierta y vuelve a ejecutar.");
  return cookieHeader;
}

function absoluteDownloadUrl(file, origin) {
  if (file.ServerRelativeUrl) return `${origin}${encodeURI(file.ServerRelativeUrl)}`;
  if (file.webUrl) return file.webUrl;
  if (file.Url) return file.Url;
  throw new Error(`El archivo ${file.Name || "sin nombre"} no tiene URL descargable.`);
}

async function main() {
  await fs.mkdir(outDir, { recursive: true });
  const source = await readJson(listPath, { files: [] });
  const files = Array.isArray(source.files) ? source.files.filter((f) => /\.pdf$/i.test(f.Name || f.name || "")) : [];
  const origin = source.siteOrigin || source.SiteOrigin || process.env.SHAREPOINT_SITE_ORIGIN || defaultOrigin;
  const cachePath = path.join(outDir, "_cache_manifest.json");
  const cache = await readJson(cachePath, {});
  const manifest = [];

  if (!files.length) {
    console.warn("No hay PDFs en el listado. Se genera manifiesto vacio para que el dashboard muestre faltantes.");
    await fs.writeFile(manifestPath, JSON.stringify([], null, 2), "utf8");
    return;
  }

  const cookieHeader = await getCookieHeader(origin);
  for (let index = 0; index < files.length; index++) {
    const file = files[index];
    const name = file.Name || file.name || `archivo_${index + 1}.pdf`;
    const localPath = path.join(outDir, safeName(name));
    const size = Number(file.Length || file.Size || 0);
    const modified = file.TimeLastModified || file.lastModifiedDateTime || "";
    const cacheKey = file.ServerRelativeUrl || file.webUrl || name;
    const cached = cache[cacheKey] || {};
    let skipped = false;
    let downloadedBytes = 0;
    try {
      if (await exists(localPath)) {
        const stat = await fs.stat(localPath);
        if ((!size || stat.size === size) && cached.modified === modified) {
          skipped = true;
          downloadedBytes = stat.size;
        }
      }
      if (!skipped) {
        const url = absoluteDownloadUrl(file, file.SiteOrigin || origin);
        const body = await requestBuffer(url, { Cookie: cookieHeader, Accept: "application/pdf,*/*" });
        await fs.writeFile(localPath, body);
        downloadedBytes = body.length;
        cache[cacheKey] = { modified, size: downloadedBytes, localPath, downloadedAt: new Date().toISOString() };
      }
      console.log(`${index + 1}/${files.length} ${skipped ? "cache" : "descargado"} ${name} ${downloadedBytes}`);
      manifest.push({ ...file, Name: name, localPath, downloadedBytes, skipped, text_error: "" });
    } catch (err) {
      console.error(`${index + 1}/${files.length} error ${name}: ${err.message}`);
      manifest.push({ ...file, Name: name, localPath, downloadedBytes: 0, skipped: false, text_error: err.message });
    }
  }
  await fs.writeFile(cachePath, JSON.stringify(cache, null, 2), "utf8");
  await fs.writeFile(manifestPath, JSON.stringify(manifest, null, 2), "utf8");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
