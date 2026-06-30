import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

const root = process.cwd();
const outputsDir = path.join(root, "outputs");
const assetsDir = path.join(outputsDir, "assets");
const port = Number(process.env.DASHBOARD_SERVER_PORT || 8787);
const host = String(process.env.DASHBOARD_SERVER_HOST || "127.0.0.1");
const defaultYear = String(process.env.SHAREPOINT_ANIO || "2026");
const defaultMonth = String(process.env.SHAREPOINT_MES || "06").padStart(2, "0");
const refreshMode = String(process.env.DASHBOARD_REFRESH_MODE || "main").toLowerCase();
const nombreRuta = String(process.env.DASHBOARD_NOMBRE_RUTA || "RUTA_NUEVA");
const soportesUrl = process.env.SHAREPOINT_SOPORTES_URL || "";
const cronogramaUrl = process.env.SHAREPOINT_CRONOGRAMA_URL || "";
const months = {
  "01": "ENERO",
  "02": "FEBRERO",
  "03": "MARZO",
  "04": "ABRIL",
  "05": "MAYO",
  "06": "JUNIO",
  "07": "JULIO",
  "08": "AGOSTO",
  "09": "SEPTIEMBRE",
  "10": "OCTUBRE",
  "11": "NOVIEMBRE",
  "12": "DICIEMBRE",
};

let running = false;
let lastRun = { ok: true, message: "Sin ejecuciones recientes." };

function appendProgress(text) {
  const clean = String(text || "");
  const previous = lastRun.tail || "";
  lastRun.tail = (previous + clean).slice(-6000);
  const lines = clean.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const step = [...lines].reverse().find((line) => line.startsWith("==>"));
  const progress = [...lines].reverse().find((line) => /^\d+\/\d+\s+/.test(line));
  if (step) {
    lastRun.message = step.replace(/^==>\s*/, "");
    lastRun.currentStep = lastRun.message;
  } else if (progress) {
    lastRun.message = progress;
    lastRun.currentStep = progress;
  }
}

function lanUrls() {
  const urls = [];
  const nets = os.networkInterfaces();
  for (const entries of Object.values(nets)) {
    for (const net of entries || []) {
      if (net.family === "IPv4" && !net.internal) {
        urls.push(`http://${net.address}:${port}/`);
      }
    }
  }
  return urls;
}

function json(res, status, body) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  });
  res.end(JSON.stringify(body, null, 2));
}

function html(res, content) {
  res.writeHead(200, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "*",
  });
  res.end(content);
}

function serveAsset(res, assetName) {
  const safe = path.basename(assetName || "");
  const full = path.join(assetsDir, safe);
  if (!fs.existsSync(full)) {
    return json(res, 404, { ok: false, message: "Asset no encontrado." });
  }
  const ext = path.extname(safe).toLowerCase();
  const types = { ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp" };
  res.writeHead(200, {
    "Content-Type": types[ext] || "application/octet-stream",
    "Cache-Control": "public, max-age=3600",
  });
  res.end(fs.readFileSync(full));
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[ch]);
}

function appPage(reqHost) {
  const dashboardSrc = `/dashboard?anio=${encodeURIComponent(defaultYear)}&mes=${encodeURIComponent(defaultMonth)}`;
  return `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard mantenimiento</title>
<style>
:root{--bg:#f4f7fb;--panel:#fff;--line:#d7e0ea;--ink:#17202a;--muted:#607086;--blue:#071936;--green:#7ee000;--red:#a33939}
*{box-sizing:border-box} body{margin:0;background:var(--bg);font-family:Arial,Helvetica,sans-serif;color:var(--ink);font-size:14px;position:relative;min-height:100vh}
body::before{content:"Realizado por Bryan Martinez";position:fixed;left:50%;top:56%;transform:translate(-50%,-50%) rotate(-22deg);font-size:clamp(34px,6vw,86px);font-weight:900;letter-spacing:.04em;color:#071936;opacity:.045;white-space:nowrap;pointer-events:none;z-index:0}
header{background:#fff;border-bottom:1px solid var(--line);color:var(--ink);padding:0;position:sticky;top:0;z-index:2;box-shadow:0 6px 18px rgba(20,31,45,.08)}
.appbar{min-height:76px;display:flex;align-items:center;justify-content:space-between;gap:20px;padding:12px 24px;border-left:8px solid #7ee000}
.app-title{display:flex;align-items:center;gap:14px;min-width:0}
.app-mark{width:50px;height:50px;border-radius:8px;background:#071936;display:grid;place-items:center;color:#7ee000;font-weight:900;font-size:22px;flex:0 0 auto}
.app-copy{min-width:0}.eyebrow{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.12em;color:#607086;margin-bottom:3px}
h1{font-size:21px;margin:0;line-height:1.15;color:#071936}.accent{color:#679d00}.sub{color:#607086;font-size:13px;max-width:780px;line-height:1.35;margin-top:3px}
.maker-mark{font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#607086;border:1px solid var(--line);border-radius:999px;padding:9px 14px;background:rgba(255,255,255,.78);white-space:nowrap}
main{padding:16px 22px;display:grid;gap:14px;position:relative;z-index:1}.panel{background:rgba(255,255,255,.94);border:1px solid var(--line);border-radius:8px;padding:14px;box-shadow:0 8px 22px rgba(20,31,45,.07);backdrop-filter:blur(3px)}
.grid{display:grid;grid-template-columns:110px 110px 170px 1fr 1fr auto;gap:10px;align-items:end}
label{display:grid;gap:5px;font-weight:700;color:#33485f;font-size:12px} input{height:36px;border:1px solid var(--line);border-radius:6px;padding:0 9px;font:inherit;min-width:0}
button{height:36px;border:1px solid var(--blue);border-radius:6px;background:var(--blue);color:#fff;font:inherit;font-weight:700;padding:0 14px;cursor:pointer;white-space:nowrap}
button:disabled{opacity:.65;cursor:wait}.status{margin-top:10px;color:var(--muted);font-size:13px}.status.ok{color:var(--green)}.status.bad{color:var(--red)}
.log{display:none;margin-top:10px;max-height:170px;overflow:auto;background:#0f1720;color:#d7f7ff;border-radius:6px;padding:10px;font:12px Consolas,monospace;white-space:pre-wrap}
.log.active{display:block}
.links{display:flex;gap:10px;flex-wrap:wrap;margin-top:8px}.links a{color:#155f9e;text-decoration:none}.links a:hover{text-decoration:underline}
iframe{width:100%;height:calc(100vh - 228px);min-height:520px;border:1px solid var(--line);border-radius:8px;background:#fff}
@media(max-width:1100px){.grid{grid-template-columns:1fr 1fr}.grid label:nth-child(n+4){grid-column:1/-1}button{width:max-content}.appbar{align-items:flex-start;flex-direction:column}.maker-mark{align-self:flex-start}h1{font-size:19px}}
</style>
</head>
<body>
<header>
  <div class="appbar">
    <div class="app-title">
      <div class="app-mark">ID</div>
      <div class="app-copy">
        <div class="eyebrow">Panel de ejecucion</div>
        <h1>Centro de control <span class="accent">documental</span></h1>
        <div class="sub">Servidor ${escapeHtml(reqHost || "")}. Consulta SharePoint con la sesión del equipo servidor y genera el reporte documental.</div>
      </div>
    </div>
    <div class="maker-mark">Realizado por Bryan Martinez</div>
  </div>
</header>
<main>
  <section class="panel">
    <form id="runForm" class="grid">
      <label>Año <input name="anio" value="${escapeHtml(defaultYear)}" inputmode="numeric" required></label>
      <label>Mes <input name="mes" value="${escapeHtml(defaultMonth)}" inputmode="numeric" required></label>
      <label>Nombre corto <input name="nombreRuta" value="${escapeHtml(nombreRuta || "RED")}"></label>
      <label>URL soportes <input name="soportesUrl" value="${escapeHtml(soportesUrl)}" placeholder="Pega la URL de la carpeta de soportes PDF" required></label>
      <label>URL cronograma <input name="cronogramaUrl" value="${escapeHtml(cronogramaUrl)}" placeholder="Pega la carpeta o archivo del cronograma" required></label>
      <button id="runBtn" type="submit">Generar / actualizar</button>
    </form>
    <div id="status" class="status">${escapeHtml(lastRun.message || "Sin ejecuciones recientes.")}</div>
    <pre id="processLog" class="log"></pre>
    <div class="links">
      <a href="/status" target="_blank" rel="noopener">Ver estado técnico</a>
      <a id="openDashboard" href="${dashboardSrc}" target="dashboardFrame">Abrir último dashboard</a>
    </div>
  </section>
  <iframe id="dashboardFrame" name="dashboardFrame" src="${dashboardSrc}"></iframe>
</main>
<script>
const form = document.getElementById("runForm");
const btn = document.getElementById("runBtn");
const statusEl = document.getElementById("status");
const logEl = document.getElementById("processLog");
const frame = document.getElementById("dashboardFrame");
const openDashboard = document.getElementById("openDashboard");
let pollTimer = null;
function setStatus(text, cls="") {
  statusEl.className = "status " + cls;
  statusEl.textContent = text;
}
function setLog(text) {
  if (!text) return;
  logEl.classList.add("active");
  logEl.textContent = text;
  logEl.scrollTop = logEl.scrollHeight;
}
async function pollStatus() {
  try {
    const res = await fetch("/status?t=" + Date.now());
    const data = await res.json();
    if (data.lastRun) {
      const prefix = data.running ? "En proceso: " : "";
      setStatus(prefix + (data.lastRun.message || "Consultando estado..."), data.lastRun.ok && !data.running ? "ok" : "");
      setLog(data.lastRun.tail || "");
    }
  } catch {
    // El resultado final del POST mostrara el error si lo hay.
  }
}
function startPolling() {
  clearInterval(pollTimer);
  pollStatus();
  pollTimer = setInterval(pollStatus, 1500);
}
function stopPolling() {
  clearInterval(pollTimer);
  pollTimer = null;
}
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const payload = Object.fromEntries(new FormData(form).entries());
  payload.mes = String(payload.mes || "").padStart(2, "0");
  btn.disabled = true;
  logEl.textContent = "";
  logEl.classList.add("active");
  setStatus("Ejecutando consulta en SharePoint. Esto puede tardar varios minutos...");
  startPolling();
  try {
    const res = await fetch("/refresh", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.message || data.tail || "Error al generar dashboard.");
    const url = (data.dashboardUrl || "/dashboard") + (data.dashboardUrl?.includes("?") ? "&" : "?") + "t=" + Date.now();
    openDashboard.href = url;
    frame.src = url;
    setStatus(data.message || "Dashboard actualizado.", "ok");
    setLog(data.tail || logEl.textContent);
    alert("Dashboard generado correctamente.");
  } catch (err) {
    setStatus("No se pudo generar el dashboard: " + err.message, "bad");
    alert("No se pudo generar el dashboard. Revisa el detalle en pantalla.");
  } finally {
    stopPolling();
    pollStatus();
    btn.disabled = false;
  }
});
</script>
</body>
</html>`;
}

function readBody(req) {
  return new Promise((resolve) => {
    let body = "";
    req.setEncoding("utf8");
    req.on("data", (chunk) => (body += chunk));
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        resolve({});
      }
    });
  });
}

function findDashboard(year = defaultYear, month = defaultMonth) {
  const monthText = months[String(month).padStart(2, "0")] || "";
  const prefix = `dashboard_soportes_pilares_horas_${year}_${String(month).padStart(2, "0")}_${monthText}`;
  const files = fs.existsSync(outputsDir)
    ? fs.readdirSync(outputsDir)
        .filter((name) => name.startsWith(prefix) && name.toLowerCase().endsWith(".html"))
        .map((name) => {
          const full = path.join(outputsDir, name);
          return { name, full, mtime: fs.statSync(full).mtimeMs };
        })
        .sort((a, b) => b.mtime - a.mtime)
    : [];
  if (files.length) return files[0].full;
  return path.join(outputsDir, `${prefix}.html`);
}

function runPowerShell(args) {
  return new Promise((resolve) => {
    const ps1 = path.join(
      outputsDir,
      refreshMode === "custom"
        ? "generar_dashboard_mantenimiento_ruta_nueva.ps1"
        : "generar_dashboard_mantenimiento.ps1"
    );
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, ...args],
      { cwd: root, windowsHide: false }
    );
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (data) => {
      const text = data.toString();
      stdout += text;
      appendProgress(text);
      process.stdout.write(text);
    });
    child.stderr.on("data", (data) => {
      const text = data.toString();
      stderr += text;
      appendProgress(text);
      process.stderr.write(text);
    });
    child.on("close", (code) => resolve({ code, stdout, stderr }));
  });
}

async function refresh(payload) {
  if (running) {
    return { ok: false, status: 409, message: "Ya hay una actualizacion en curso." };
  }
  running = true;
  const year = String(payload.anio || defaultYear);
  const month = String(payload.mes || defaultMonth).padStart(2, "0");
  const effectiveNombreRuta = String(payload.nombreRuta || nombreRuta || "RED");
  const effectiveSoportesUrl = String(payload.soportesUrl || soportesUrl || "");
  const effectiveCronogramaUrl = String(payload.cronogramaUrl || cronogramaUrl || "");
  lastRun = { ok: false, running: true, message: `Actualizando ${year}-${month}...` };
  try {
    const args = ["-Anio", year, "-Mes", String(Number(month))];
    if (refreshMode === "custom") {
      if (!effectiveSoportesUrl || !effectiveCronogramaUrl) {
        throw new Error("Digite la URL de soportes y la URL del cronograma.");
      }
      args.push("-NombreRuta", effectiveNombreRuta, "-SoportesUrl", effectiveSoportesUrl, "-CronogramaUrl", effectiveCronogramaUrl);
    }
    const result = await runPowerShell(args);
    const dashboardFile = findDashboard(year, month);
    const ok = result.code === 0 && fs.existsSync(dashboardFile);
    lastRun = {
      ok,
      running: false,
      code: result.code,
      dashboardFile,
      message: ok ? "Dashboard actualizado." : "El proceso termino con error.",
      tail: (result.stderr || result.stdout).slice(-4000),
    };
    return {
      ok,
      status: ok ? 200 : 500,
      message: lastRun.message,
      dashboardFile,
      dashboardUrl: `/dashboard?anio=${encodeURIComponent(year)}&mes=${encodeURIComponent(month)}`,
      code: result.code,
      tail: lastRun.tail,
    };
  } finally {
    running = false;
  }
}

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") return json(res, 204, {});
  const url = new URL(req.url, `http://${req.headers.host || `127.0.0.1:${port}`}`);
  if (url.pathname.startsWith("/assets/")) return serveAsset(res, decodeURIComponent(url.pathname.slice("/assets/".length)));
  if (url.pathname === "/status") return json(res, 200, { running, lastRun });
  if (url.pathname === "/refresh" && req.method === "POST") {
    const payload = await readBody(req);
    const result = await refresh(payload);
    return json(res, result.status, result);
  }
  if (url.pathname === "/" || url.pathname === "/app") {
    return html(res, appPage(req.headers.host));
  }
  if (url.pathname !== "/dashboard") {
    return json(res, 404, { ok: false, message: "Ruta no encontrada." });
  }
  const year = url.searchParams.get("anio") || defaultYear;
  const month = url.searchParams.get("mes") || defaultMonth;
  const dashboardFile = findDashboard(year, month);
  if (!fs.existsSync(dashboardFile)) {
    return html(
      res,
      `<!doctype html><html lang="es"><meta charset="utf-8"><body style="font-family:Arial;padding:24px"><h2>Sin dashboard generado</h2><p>No encontré dashboard para ${escapeHtml(year)}-${escapeHtml(month)}.</p><p>Digite año, mes y rutas en el formulario superior y pulse <strong>Generar / actualizar</strong>.</p><p style="color:#607086">${escapeHtml(dashboardFile)}</p></body></html>`
    );
  }
  return html(res, fs.readFileSync(dashboardFile, "utf8"));
});

server.listen(port, host, () => {
  const localUrl = `http://127.0.0.1:${port}/`;
  const urls = host === "0.0.0.0" ? lanUrls() : [];
  console.log("============================================================");
  console.log("Dashboard interactivo listo");
  console.log(`URL local: ${localUrl}`);
  for (const url of urls) console.log(`URL red:   ${url}`);
  console.log(`Modo de refresco: ${refreshMode === "custom" ? "ruta nueva" : "ruta original"}`);
  console.log("Deja esta ventana abierta para que el boton Refrescar funcione.");
  console.log("============================================================");
  if (process.env.DASHBOARD_OPEN_BROWSER !== "0") {
    spawn("cmd.exe", ["/c", "start", "", localUrl], { detached: true, stdio: "ignore" }).unref();
  }
});




