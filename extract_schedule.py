import fs from "node:fs/promises";
import http from "node:http";
import https from "node:https";

const inputUrl = process.env.SHAREPOINT_CRONOGRAMA_URL || "";
const anio = process.env.SHAREPOINT_ANIO || "2026";
const mesNombre = process.env.SHAREPOINT_MES_NOMBRE || "MAYO";
const carpetaMes = process.env.SHAREPOINT_CARPETA_MES || "05-MAYO";
const outPath = "work/cronograma_sharepoint.xlsx";
const metaPath = "work/cronograma_sharepoint_meta.json";

function failWithHelp(err) {
  console.error("");
  console.error("ERROR AL DESCARGAR EL CRONOGRAMA");
  console.error(String(err?.message || err));
  console.error("");
  console.error("Verifica la URL pegada:");
  console.error("- Puede ser la URL de la carpeta donde esta el Excel del cronograma.");
  console.error("- Tambien sirve el enlace directo del archivo .xlsx si empieza como /:x:/r/...");
  console.error("- No pegues la URL interna de Excel Online que contiene _layouts/15/Doc.aspx.");
  console.error("");
  process.exit(1);
}

process.on("uncaughtException", failWithHelp);
process.on("unhandledRejection", failWithHelp);

if (!inputUrl.trim()) {
  throw new Error("Falta SHAREPOINT_CRONOGRAMA_URL. Pega la URL del archivo o carpeta donde esta el cronograma.");
}

function parseSharePointPath(rawUrl) {
  const url = new URL(rawUrl);
  if (url.pathname.toLowerCase().includes("/_layouts/15/doc.aspx")) {
    throw new Error("La URL pegada es de Excel Online (_layouts/15/Doc.aspx). Cierra la vista del Excel y copia el vinculo desde SharePoint sobre el archivo o pega la URL de la carpeta donde esta el cronograma.");
  }
  const origin = url.origin;
  let itemPath = "";

  const id = url.searchParams.get("id");
  if (id) {
    itemPath = decodeURIComponent(id);
  } else if (url.pathname.includes("/:x:/r/") || url.pathname.includes("/:f:/r/")) {
    itemPath = decodeURIComponent(url.pathname.replace(/^\/:[xf]:\/r/i, ""));
  } else {
    itemPath = decodeURIComponent(url.pathname);
  }

  if (!itemPath.startsWith("/")) itemPath = `/${itemPath}`;
  const parts = itemPath.split("/").filter(Boolean);
  if (parts.length < 2) {
    throw new Error(`No pude detectar el sitio de SharePoint desde la ruta: ${itemPath}`);
  }
  const sitePath = `/${parts[0]}/${parts[1]}`;
  return { origin, sitePath, itemPath };
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

async function getCookiesFromEdge(origin) {
  const tabs = await getJsonHttp("http://127.0.0.1:9222/json");
  const page = tabs.find((t) => t.type === "page" && t.webSocketDebuggerUrl);
  if (!page) throw new Error("No hay pestanas controlables en Edge.");

  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  await send(ws, "Network.enable");
  const result = await send(ws, "Network.getCookies", { urls: [origin] });
  ws.close();

  const cookieHeader = (result.cookies || []).map((c) => `${c.name}=${c.value}`).join("; ");
  if (!cookieHeader) throw new Error(`No encontre cookies de SharePoint para ${origin}.`);
  return cookieHeader;
}

function normalizeName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function isMonthFolder(name) {
  const normalized = normalizeName(name);
  const expected = normalizeName(carpetaMes);
  const month = normalizeName(mesNombre);
  const monthNumber = carpetaMes.split("-")[0].replace(/^0+/, "");
  return (
    normalized === expected ||
    normalized.includes(expected) ||
    normalized.includes(month) ||
    normalized.startsWith(`${monthNumber}-`) ||
    normalized.startsWith(`${monthNumber} `) ||
    normalized.startsWith(carpetaMes.split("-")[0])
  );
}

async function listFilesInFolder(ctx, headers, folderPath) {
  const filesUrl = `${ctx.origin}${ctx.sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)/Files?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length&$top=5000&@u=${quoteFolder(folderPath)}`;
  let files = [];
  try {
    files = (await getJsonHttps(filesUrl, headers)).value || [];
  } catch (err) {
    throw new Error(`No pude leer la carpeta indicada para buscar el cronograma. Revisa que sea una carpeta de SharePoint y que tengas acceso. Ruta detectada: ${folderPath}. Detalle: ${err.message}`);
  }
  return files;
}

async function listFoldersInFolder(ctx, headers, folderPath) {
  const foldersUrl = `${ctx.origin}${ctx.sitePath}/_api/web/GetFolderByServerRelativeUrl(@u)/Folders?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified&$top=5000&@u=${quoteFolder(folderPath)}`;
  try {
    return ((await getJsonHttps(foldersUrl, headers)).value || []).filter((f) => f.Name !== "Forms");
  } catch (err) {
    throw new Error(`No pude leer subcarpetas para ubicar el mes ${carpetaMes}. Ruta detectada: ${folderPath}. Detalle: ${err.message}`);
  }
}

function pickCronograma(files) {
  const excelFiles = files.filter((f) => /\.xlsx$/i.test(f.Name));
  const preferred = excelFiles
    .filter((f) => {
      const name = normalizeName(f.Name);
      return name.includes("cronograma") && name.includes(normalizeName(mesNombre)) && name.includes(String(anio));
    })
    .sort((a, b) => new Date(b.TimeLastModified) - new Date(a.TimeLastModified));

  if (preferred.length) return preferred[0];

  const cronogramas = excelFiles
    .filter((f) => normalizeName(f.Name).includes("cronograma"))
    .sort((a, b) => new Date(b.TimeLastModified) - new Date(a.TimeLastModified));
  if (cronogramas.length) return cronogramas[0];
  if (excelFiles.length === 1) return excelFiles[0];
  return null;
}

async function findCronogramaInFolder(ctx, headers) {
  let searchedPath = ctx.itemPath;
  let files = await listFilesInFolder(ctx, headers, searchedPath);
  let selected = pickCronograma(files);
  if (selected) {
    selected.detectedFolder = searchedPath;
    return selected;
  }

  const folders = await listFoldersInFolder(ctx, headers, searchedPath);
  const yearFolder = folders.find((f) => normalizeName(f.Name) === normalizeName(anio));
  if (yearFolder) {
    searchedPath = yearFolder.ServerRelativeUrl;
    files = await listFilesInFolder(ctx, headers, searchedPath);
    selected = pickCronograma(files);
    if (selected) {
      selected.detectedFolder = searchedPath;
      selected.detectedYearFolder = yearFolder.Name;
      return selected;
    }

    const monthFolders = await listFoldersInFolder(ctx, headers, searchedPath);
    const nestedMonthFolder = monthFolders.find((f) => isMonthFolder(f.Name));
    if (nestedMonthFolder) {
      searchedPath = nestedMonthFolder.ServerRelativeUrl;
      files = await listFilesInFolder(ctx, headers, searchedPath);
      selected = pickCronograma(files);
      if (selected) {
        selected.detectedFolder = searchedPath;
        selected.detectedYearFolder = yearFolder.Name;
        selected.detectedMonthFolder = nestedMonthFolder.Name;
        return selected;
      }
    }
  }

  const monthFolder = folders.find((f) => isMonthFolder(f.Name));
  if (monthFolder) {
    searchedPath = monthFolder.ServerRelativeUrl;
    files = await listFilesInFolder(ctx, headers, searchedPath);
    selected = pickCronograma(files);
    if (selected) {
      selected.detectedFolder = searchedPath;
      selected.detectedMonthFolder = monthFolder.Name;
      return selected;
    }
  }

  const excelFiles = files.filter((f) => /\.xlsx$/i.test(f.Name));
  const folderNames = folders.map((f) => f.Name).join(" | ");

  throw new Error(`No encontre un Excel de cronograma. Si pegaste la carpeta del anio, confirme que exista una subcarpeta del mes ${carpetaMes}. Subcarpetas encontradas: ${folderNames}. Excels encontrados en la carpeta revisada: ${excelFiles.map((f) => f.Name).join(" | ")}`);
}

const ctx = parseSharePointPath(inputUrl);
const cookieHeader = await getCookiesFromEdge(ctx.origin);
const headers = {
  Cookie: cookieHeader,
  Accept: "application/json;odata=nometadata",
};

let selected = null;
if (/\.xlsx$/i.test(ctx.itemPath)) {
  selected = {
    Name: ctx.itemPath.split("/").pop(),
    ServerRelativeUrl: ctx.itemPath,
    TimeLastModified: "",
    source: "archivo indicado",
  };
} else {
  selected = await findCronogramaInFolder(ctx, headers);
  selected.source = "carpeta indicada";
}

const downloadUrl = ctx.origin + encodeURI(selected.ServerRelativeUrl);
const bytes = await requestBuffer(downloadUrl, {
  Cookie: cookieHeader,
  Accept: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
});

await fs.writeFile(outPath, bytes);
await fs.writeFile(metaPath, JSON.stringify({
  sourceUrl: inputUrl,
  selected,
  downloadedBytes: bytes.length,
  outPath,
}, null, 2), "utf8");

console.log(JSON.stringify({
  ok: true,
  source: selected.source,
  name: selected.Name,
  detectedFolder: selected.detectedFolder,
  detectedYearFolder: selected.detectedYearFolder || "",
  detectedMonthFolder: selected.detectedMonthFolder || "",
  modified: selected.TimeLastModified,
  bytes: bytes.length,
  outPath,
}, null, 2));
