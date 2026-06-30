import fs from "node:fs/promises";
import http from "node:http";
import https from "node:https";

const inputUrl = process.env.SHAREPOINT_SOPORTES_URL || "";
const outPath = "work/sharepoint_files.json";
const debugPath = "work/sharepoint_supports_debug.json";
const anio = process.env.SHAREPOINT_ANIO || "2026";
const carpetaMes = process.env.SHAREPOINT_CARPETA_MES || "05-MAYO";
const mesNombre = process.env.SHAREPOINT_MES_NOMBRE || "MAYO";

function failWithHelp(err) {
  console.error("");
  console.error("ERROR AL LISTAR SOPORTES");
  console.error(String(err?.message || err));
  if (String(err?.message || err).includes("ECONNREFUSED") || String(err?.message || err).includes("9222")) {
    console.error("");
    console.error("No pude conectarme a Edge en modo controlable.");
    console.error("Deja abierta la ventana de Edge que abre el backend y completa el inicio de sesion si SharePoint lo pide.");
  }
  console.error("");
  console.error("Verifica:");
  console.error(`- Año digitado: ${anio}`);
  console.error(`- Mes digitado: ${carpetaMes}`);
  console.error("- La URL de soportes puede ser la carpeta padre, la carpeta del año o la carpeta del mes.");
  console.error("- Si pegaste la carpeta padre, debe existir una subcarpeta del año y luego una del mes.");
  console.error("- Dentro del mes deben existir PDFs, directamente o en subcarpetas.");
  console.error("");
  process.exit(1);
}

process.on("uncaughtException", failWithHelp);
process.on("unhandledRejection", failWithHelp);

if (!inputUrl.trim()) {
  throw new Error("Falta SHAREPOINT_SOPORTES_URL. Pega la URL nueva de la carpeta de soportes.");
}

function parseSharePointFolder(rawUrl) {
  const url = new URL(rawUrl);
  const origin = url.origin;
  let folderPath = "";

  const id = url.searchParams.get("id");
  if (id) {
    folderPath = decodeURIComponent(id);
  } else if (url.pathname.includes("/:f:/r/")) {
    folderPath = decodeURIComponent(url.pathname.replace(/^\/:f:\/r/i, ""));
  } else {
    folderPath = decodeURIComponent(url.pathname);
  }

  if (!folderPath.startsWith("/")) folderPath = `/${folderPath}`;

  const parts = folderPath.split("/").filter(Boolean);
  if (parts.length < 2) {
    throw new Error(`No pude detectar el sitio de SharePoint desde la ruta: ${folderPath}`);
  }
  const sitePath = `/${parts[0]}/${parts[1]}`;
  return { origin, sitePath, folderPath };
}

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

function normalizeName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function isYearFolder(name) {
  return normalizeName(name) === normalizeName(anio);
}

function isMonthFolder(name) {
  const normalized = normalizeName(name);
  const expected = normalizeName(carpetaMes);
  const month = normalizeName(mesNombre);
  const monthNumberPadded = carpetaMes.split("-")[0];
  const monthNumber = monthNumberPadded.replace(/^0+/, "");
  return (
    normalized === expected ||
    normalized.includes(expected) ||
    normalized.includes(month) ||
    normalized.startsWith(`${monthNumberPadded} `) ||
    normalized.startsWith(`${monthNumberPadded}-`) ||
    normalized.startsWith(`${monthNumber} `) ||
    normalized.startsWith(`${monthNumber}-`)
  );
}

function isSupportFolder(name) {
  const normalized = normalizeName(name);
  return (
    normalized === "soporte de visitas" ||
    normalized === "soportes de visita" ||
    normalized === "soportes de visitas" ||
    normalized === "soporte visitas" ||
    normalized === "soportes visita" ||
    normalized === "soportes visitas" ||
    normalized.includes("soporte de visita") ||
    normalized.includes("soportes de visita") ||
    normalized.includes("soportes")
  );
}

async function getCookiesFromEdge(origin) {
  const tabs = await getJsonHttp("http://127.0.0.1:9222/json");
  const page = tabs.find((t) => t.type === "page" && t.webSocketDebuggerUrl);
  if (!page) throw new Error("No hay pestanas controlables en Edge. Abre Edge desde el script.");

  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  await send(ws, "Network.enable");
  const result = await send(ws, "Network.getCookies", { urls: [origin] });
  ws.close();

  const cookieHeader = (result.cookies || []).map((c) => `${c.name}=${c.value}`).join("; ");
  if (!cookieHeader) {
    throw new Error("No encontre cookies del nuevo SharePoint en Edge. Inicia sesion en la ventana abierta y vuelve a ejecutar.");
  }
  return cookieHeader;
}

async function readFolder(ctx, path, headers, depth = 0) {
  const folderInfo = {
    path,
    depth,
    files: [],
    folders: [],
    errors: [],
  };
  const base = `${ctx.origin}${ctx.sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)`;
  const filesUrl = `${base}/Files?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length&$top=5000&@u=${quoteFolder(path)}`;
  const foldersUrl = `${base}/Folders?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified&$top=5000&@u=${quoteFolder(path)}`;

  let files = [];
  let folders = [];
  try {
    files = (await getJsonHttps(filesUrl, headers)).value || [];
    folderInfo.files = files.map((f) => f.Name);
  } catch (err) {
    folderInfo.errors.push(`Files: ${err.message}`);
  }
  try {
    folders = ((await getJsonHttps(foldersUrl, headers)).value || []).filter((f) => f.Name !== "Forms");
    folderInfo.folders = folders.map((f) => f.Name);
  } catch (err) {
    folderInfo.errors.push(`Folders: ${err.message}`);
  }

  const nested = [];
  const nestedFolders = [];
  const debugFolders = [folderInfo];
  if (depth < 8) {
    for (const folder of folders) {
      const child = await readFolder(ctx, folder.ServerRelativeUrl, headers, depth + 1);
      nested.push(...child.files);
      nestedFolders.push(...child.allFolders);
      debugFolders.push(...child.debugFolders);
    }
  }
  return {
    folders,
    allFolders: [...folders.map((f) => ({ ...f, ParentFolder: path })), ...nestedFolders],
    debugFolders,
    files: [...files.map((f) => ({ ...f, ParentFolder: path, SiteOrigin: ctx.origin })), ...nested],
  };
}

async function listFolders(ctx, path, headers) {
  const base = `${ctx.origin}${ctx.sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)`;
  const foldersUrl = `${base}/Folders?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified&$top=5000&@u=${quoteFolder(path)}`;
  return ((await getJsonHttps(foldersUrl, headers)).value || []).filter((f) => f.Name !== "Forms");
}

async function scopedFolderForPeriod(ctx, headers) {
  let path = ctx.folderPath;
  let folders = await listFolders(ctx, path, headers);
  const originalFolders = folders.map((f) => f.Name);

  const yearFolder = folders.find((f) => isYearFolder(f.Name));
  if (yearFolder) {
    path = yearFolder.ServerRelativeUrl;
    folders = await listFolders(ctx, path, headers);
  } else if (!normalizeName(path).split(" ").includes(normalizeName(anio))) {
    throw new Error(`No encontre la carpeta del año ${anio}. Subcarpetas encontradas en la ruta pegada: ${originalFolders.join(" | ") || "ninguna"}. Ruta revisada: ${path}`);
  }

  const monthFolder = folders.find((f) => isMonthFolder(f.Name));
  if (monthFolder) {
    path = monthFolder.ServerRelativeUrl;
    folders = await listFolders(ctx, path, headers);
  } else if (!isMonthFolder(path.split("/").pop() || "")) {
    throw new Error(`No encontre la carpeta del mes ${carpetaMes}. Subcarpetas encontradas en ${path}: ${folders.map((f) => f.Name).join(" | ") || "ninguna"}.`);
  } else {
    folders = await listFolders(ctx, path, headers);
  }

  const supportFolder = folders.find((f) => isSupportFolder(f.Name));
  if (supportFolder) {
    path = supportFolder.ServerRelativeUrl;
  }

  return path;
}

const ctx = parseSharePointFolder(inputUrl);
const cookieHeader = await getCookiesFromEdge(ctx.origin);
const headers = {
  Cookie: cookieHeader,
  Accept: "application/json;odata=nometadata",
};

const scopedFolderPath = await scopedFolderForPeriod(ctx, headers);
const root = await readFolder(ctx, scopedFolderPath, headers);
const pdfFiles = root.files.filter((f) => /\.pdf$/i.test(f.Name || ""));
const ignoredNonPdf = root.files.length - pdfFiles.length;
await fs.writeFile(debugPath, JSON.stringify({
  ok: true,
  sourceUrl: inputUrl,
  requestedPeriod: `${anio}/${carpetaMes}`,
  inputFolderPath: ctx.folderPath,
  scopedFolderPath,
  totalFilesRead: root.files.length,
  pdfFilesRead: pdfFiles.length,
  ignoredNonPdf,
  folders: root.debugFolders,
}, null, 2), "utf8");
if (!pdfFiles.length) {
  console.warn(`No encontre PDFs para ${anio}/${carpetaMes}. Se generara el dashboard con soportes en cero/faltantes. Diagnostico guardado en ${debugPath}`);
}
const output = {
  ok: true,
  sourceUrl: inputUrl,
  siteOrigin: ctx.origin,
  sitePath: ctx.sitePath,
  folderPath: scopedFolderPath,
  inputFolderPath: ctx.folderPath,
  requestedPeriod: `${anio}/${carpetaMes}`,
  count: pdfFiles.length,
  ignoredNonPdf,
  rootFolders: root.folders,
  allFolders: root.allFolders,
  debugFolders: root.debugFolders,
  files: pdfFiles,
};

await fs.writeFile(outPath, JSON.stringify(output, null, 2), "utf8");
console.log(JSON.stringify({
  ok: output.ok,
  sourceUrl: output.sourceUrl,
  requestedPeriod: output.requestedPeriod,
  inputFolderPath: output.inputFolderPath,
  folderPath: output.folderPath,
  count: output.count,
  ignoredNonPdf: output.ignoredNonPdf,
  rootFolders: output.rootFolders.map((f) => f.Name),
  foldersRead: output.allFolders.map((f) => `${f.ParentFolder?.split("/").pop() || ""}/${f.Name}`),
  debugPath,
}, null, 2));
