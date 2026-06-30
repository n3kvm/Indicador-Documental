import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { pathToFileURL } from "node:url";

const inputUrl = process.env.SHAREPOINT_SOPORTES_URL || "";
const anio = process.env.SHAREPOINT_ANIO || "2026";
const carpetaMes = process.env.SHAREPOINT_CARPETA_MES || "05-MAYO";
const mesNombre = process.env.SHAREPOINT_MES_NOMBRE || "MAYO";
const outPath = process.env.SHAREPOINT_FILES_PATH || "work/sharepoint_files.json";
const manifestPath = process.env.SHAREPOINT_PDF_MANIFEST_PATH || "work/sharepoint_pdf_manifest.json";
const resultPath = process.env.LOCAL_SYNC_SUPPORTS_RESULT_PATH || "work/local_synced_supports_result.json";
const maxDepth = Number(process.env.SHAREPOINT_LOCAL_SEARCH_DEPTH || 8);

function normalize(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[_-]+/g, " ")
    .replace(/[^A-Za-z0-9]+/g, " ")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

function monthMatches(name) {
  const n = normalize(name);
  const cm = normalize(carpetaMes);
  const mn = normalize(mesNombre);
  const num = carpetaMes.split("-")[0].replace(/^0+/, "") || carpetaMes.split("-")[0];
  return n === cm || n.includes(cm) || (n.includes(mn) && new RegExp(`(^| )0?${num}( |$|-)`).test(n));
}

function yearMatches(name) {
  return normalize(name) === normalize(anio) || normalize(name).includes(normalize(anio));
}

function supportMatches(name) {
  const n = normalize(name);
  return n === "soporte de visitas" || n === "soportes de visita" || n.includes("soporte de visita") || n.includes("soportes de visita");
}

function parseFolderSegments(rawUrl) {
  if (!rawUrl || /^[A-Za-z]:[\\/]/.test(rawUrl) || rawUrl.startsWith("\\\\")) {
    return { localPath: rawUrl, segments: [] };
  }
  let folderPath = rawUrl;
  try {
    const url = new URL(rawUrl);
    const id = url.searchParams.get("id");
    if (id) folderPath = decodeURIComponent(id);
    else if (url.pathname.includes("/:f:/r/")) folderPath = decodeURIComponent(url.pathname.replace(/^\/:f:\/r/i, ""));
    else folderPath = decodeURIComponent(url.pathname);
  } catch {}
  return { localPath: "", segments: folderPath.split(/[\\/]+/).filter(Boolean).map((s) => s.trim()).filter(Boolean) };
}

async function exists(target) {
  try { await fs.access(target); return true; } catch { return false; }
}

async function statSafe(target) {
  try { return await fs.stat(target); } catch { return null; }
}

async function readDirSafe(target) {
  try { return await fs.readdir(target, { withFileTypes: true }); } catch { return []; }
}

async function candidateRoots() {
  const roots = new Set();
  const envKeys = ["OneDriveCommercial", "OneDrive", "USERPROFILE", "HOMEDRIVE"];
  for (const key of envKeys) {
    const value = process.env[key];
    if (value) roots.add(value);
  }
  const home = os.homedir();
  roots.add(home);
  for (const base of [home, "D:\\", "C:\\"]) {
    const entries = await readDirSafe(base);
    for (const entry of entries) {
      if (entry.isDirectory() && /OneDrive|COMFANDI|BRILLASEO/i.test(entry.name)) {
        roots.add(path.join(base, entry.name));
      }
    }
  }
  const extra = (process.env.SHAREPOINT_LOCAL_ROOTS || "").split(";").map((x) => x.trim()).filter(Boolean);
  for (const root of extra) roots.add(root);
  const out = [];
  for (const root of roots) {
    if (root && await exists(root)) out.push(path.resolve(root));
  }
  return [...new Set(out)];
}

function scorePath(folder, segments) {
  const full = normalize(folder);
  let score = 0;
  const meaningful = segments.filter((s) => !/^(personal|teams|sites|documents|documentos compartidos|shared documents)$/i.test(normalize(s)));
  for (const segment of meaningful) {
    const ns = normalize(segment);
    if (!ns) continue;
    if (full.includes(ns)) score += Math.min(8, ns.length / 4);
  }
  if (full.includes(normalize(anio))) score += 10;
  if (monthMatches(path.basename(folder))) score += 18;
  if (supportMatches(path.basename(folder))) score += 25;
  return score;
}

async function findFoldersByName(root, matcher, limit = 25) {
  const found = [];
  const queue = [{ dir: root, depth: 0 }];
  while (queue.length && found.length < limit) {
    const { dir, depth } = queue.shift();
    const entries = await readDirSafe(dir);
    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const full = path.join(dir, entry.name);
      if (matcher(entry.name, full)) found.push(full);
      if (depth < maxDepth && !/^\.|node_modules|AppData|Windows|Program Files/i.test(entry.name)) {
        queue.push({ dir: full, depth: depth + 1 });
      }
    }
  }
  return found;
}

async function descendForPeriod(base) {
  let current = base;
  let entries = await readDirSafe(current);
  const year = entries.find((e) => e.isDirectory() && yearMatches(e.name));
  if (year) current = path.join(current, year.name);

  entries = await readDirSafe(current);
  const month = entries.find((e) => e.isDirectory() && monthMatches(e.name));
  if (month) current = path.join(current, month.name);

  entries = await readDirSafe(current);
  const support = entries.find((e) => e.isDirectory() && supportMatches(e.name));
  if (support) current = path.join(current, support.name);
  return current;
}

async function collectPdfs(root, maxPdfDepth = 5) {
  const out = [];
  const queue = [{ dir: root, depth: 0 }];
  while (queue.length) {
    const { dir, depth } = queue.shift();
    const entries = await readDirSafe(dir);
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory() && depth < maxPdfDepth) queue.push({ dir: full, depth: depth + 1 });
      else if (entry.isFile() && /\.pdf$/i.test(entry.name)) out.push(full);
    }
  }
  return out;
}

async function bestLocalFolder() {
  const parsed = parseFolderSegments(inputUrl);
  if (parsed.localPath && await exists(parsed.localPath)) return await descendForPeriod(path.resolve(parsed.localPath));
  const segments = parsed.segments;
  const lastSegments = segments.slice(-5).map(normalize).filter(Boolean);
  const roots = await candidateRoots();
  const candidates = [];
  for (const root of roots) {
    const matches = await findFoldersByName(root, (name, full) => {
      const n = normalize(name);
      return lastSegments.includes(n) || yearMatches(name) || monthMatches(name) || supportMatches(name);
    });
    for (const match of matches) {
      const scoped = await descendForPeriod(match);
      const pdfs = await collectPdfs(scoped, 4);
      candidates.push({ base: match, scoped, pdfs, score: scorePath(scoped, segments) + Math.min(pdfs.length, 20) });
    }
  }
  candidates.sort((a, b) => b.score - a.score || b.pdfs.length - a.pdfs.length);
  return candidates[0]?.scoped || "";
}

function fileRow(file, scopedRoot) {
  const rel = path.relative(scopedRoot, file).replace(/\\/g, "/");
  const parent = path.dirname(rel) === "." ? scopedRoot : path.join(scopedRoot, path.dirname(rel));
  return fs.stat(file).then((st) => ({
    Name: path.basename(file),
    ServerRelativeUrl: "",
    ParentFolder: parent,
    SiteOrigin: "file://local-sync",
    TimeCreated: st.birthtime.toISOString(),
    TimeLastModified: st.mtime.toISOString(),
    Length: st.size,
    LocalPath: file,
    localPath: file,
    Source: "OneDrive sincronizado",
    WebUrl: pathToFileURL(file).toString(),
  }));
}

async function writeNotFound(message, checkedFolder = "") {
  const result = { ok: true, localFound: false, count: 0, checkedFolder, message };
  await fs.writeFile(resultPath, JSON.stringify(result, null, 2), "utf8");
  console.log(JSON.stringify(result, null, 2));
}

async function main() {
  const scoped = await bestLocalFolder();
  if (!scoped) return await writeNotFound("No encontre carpeta local sincronizada compatible con la URL.");
  const pdfs = await collectPdfs(scoped, 5);
  if (!pdfs.length) return await writeNotFound("Encontre carpeta local, pero no contiene PDFs para el periodo.", scoped);

  const files = [];
  for (const pdf of pdfs) files.push(await fileRow(pdf, scoped));
  const output = {
    ok: true,
    localSync: true,
    sourceUrl: inputUrl,
    siteOrigin: "file://local-sync",
    sitePath: "",
    folderPath: scoped,
    inputFolderPath: scoped,
    requestedPeriod: `${anio}/${carpetaMes}`,
    count: files.length,
    ignoredNonPdf: 0,
    rootFolders: [],
    allFolders: [],
    debugFolders: [{ path: scoped, files: files.map((f) => f.Name), folders: [], errors: [] }],
    files,
  };
  await fs.writeFile(outPath, JSON.stringify(output, null, 2), "utf8");
  await fs.writeFile(manifestPath, JSON.stringify(files.map((f) => ({ ...f, downloadedBytes: f.Length, skipped: true })), null, 2), "utf8");
  const result = { ok: true, localFound: true, count: files.length, checkedFolder: scoped, message: "Usando PDFs sincronizados localmente. No se descargaran soportes desde SharePoint." };
  await fs.writeFile(resultPath, JSON.stringify(result, null, 2), "utf8");
  console.log(JSON.stringify(result, null, 2));
}

main().catch(async (err) => {
  await writeNotFound(`Error buscando carpeta local: ${err.message}`);
});
