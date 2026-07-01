param(
  [Parameter(Mandatory=$true)]
  [string]$SoportesUrl,
  [string]$CronogramaUrl = "",
  [string]$NombreRuta = "ruta_nueva",
  [string]$CronogramaPath = "",
  [int]$Anio = (Get-Date).Year,
  [int]$Mes = (Get-Date).Month,
  [switch]$NoAbrirEdge
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$WorkDir = Join-Path $ProjectRoot "work"
$OutputsDir = Join-Path $ProjectRoot "outputs"
$Python = "C:\Users\DELL\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$Node = "C:\Users\DELL\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"

$EdgeProfile = Join-Path $WorkDir "edge-sharepoint-profile-ruta-nueva"
$CronogramaCopia = Join-Path $WorkDir "cronograma_mayo_2026_copia.xlsx"
$CronogramaSharePoint = Join-Path $WorkDir "cronograma_sharepoint.xlsx"
$Meses = @{
  1 = "ENERO"; 2 = "FEBRERO"; 3 = "MARZO"; 4 = "ABRIL"; 5 = "MAYO"; 6 = "JUNIO";
  7 = "JULIO"; 8 = "AGOSTO"; 9 = "SEPTIEMBRE"; 10 = "OCTUBRE"; 11 = "NOVIEMBRE"; 12 = "DICIEMBRE"
}
if (-not $Meses.ContainsKey($Mes)) { throw "Mes invalido: $Mes. Usa un numero de 1 a 12." }
$MesNumero = "{0:D2}" -f $Mes
$MesNombre = $Meses[$Mes]
$CarpetaMes = "$MesNumero-$MesNombre"
$NombreSeguro = ($NombreRuta -replace '[^A-Za-z0-9_-]+', '_').Trim('_')
if ([string]::IsNullOrWhiteSpace($NombreSeguro)) { $NombreSeguro = "ruta_nueva" }

$env:SHAREPOINT_ANIO = [string]$Anio
$env:SHAREPOINT_MES = $MesNumero
$env:SHAREPOINT_MES_NOMBRE = $MesNombre
$env:SHAREPOINT_CARPETA_MES = $CarpetaMes
$env:SHAREPOINT_SOPORTES_URL = $SoportesUrl
$env:SHAREPOINT_CRONOGRAMA_URL = $CronogramaUrl
$env:SHAREPOINT_PDF_CACHE_DIR = Join-Path $WorkDir "sharepoint_pdfs\$Anio\$CarpetaMes\$NombreSeguro"
$DashboardFileName = "dashboard_soportes_pilares_horas_${Anio}_${MesNumero}_${MesNombre}_${NombreSeguro}.html"
$Dashboard = Join-Path $OutputsDir $DashboardFileName
$env:DASHBOARD_OUTPUT_PATH = $Dashboard
$env:APP_UPLOAD_VALIDATION_PATH = Join-Path $WorkDir "app_upload_validation.json"
if ([string]::IsNullOrWhiteSpace($env:APP_AUTOGESTIONES_CSV_PATH)) {
  $env:APP_AUTOGESTIONES_CSV_PATH = "D:\Datos\OneDrive - BRILLASEO SAS\Descargas\Master Autogestiones.csv"
}

function Write-Step($Message) {
  Write-Host ""
  Write-Host "==> $Message" -ForegroundColor Cyan
}

function Invoke-Step($File, $Arguments) {
  Write-Host "Ejecutando: $File $Arguments" -ForegroundColor DarkGray
  Push-Location $ProjectRoot
  try {
    & $File @Arguments
    $exitCode = $LASTEXITCODE
  } finally {
    Pop-Location
  }
  if ($exitCode -ne 0) {
    $debugFile = Join-Path $WorkDir "sharepoint_supports_debug.json"
    if (Test-Path -LiteralPath $debugFile) {
      Write-Host ""
      Write-Host "Diagnostico de carpetas/archivos SharePoint:" -ForegroundColor Yellow
      try {
        $debug = Get-Content -LiteralPath $debugFile -Raw | ConvertFrom-Json
        Write-Host "Periodo: $($debug.requestedPeriod)" -ForegroundColor Yellow
        Write-Host "Ruta revisada: $($debug.scopedFolderPath)" -ForegroundColor Yellow
        Write-Host "Archivos totales leidos: $($debug.totalFilesRead)" -ForegroundColor Yellow
        Write-Host "PDFs leidos: $($debug.pdfFilesRead)" -ForegroundColor Yellow
        $debug.folders | Select-Object -First 30 | ForEach-Object {
          $files = if ($_.files) { ($_.files -join " | ") } else { "sin archivos" }
          $folders = if ($_.folders) { ($_.folders -join " | ") } else { "sin subcarpetas" }
          Write-Host "- $($_.path)" -ForegroundColor DarkYellow
          Write-Host "  Subcarpetas: $folders" -ForegroundColor DarkYellow
          Write-Host "  Archivos: $files" -ForegroundColor DarkYellow
          if ($_.errors) { Write-Host "  Errores: $($_.errors -join ' | ')" -ForegroundColor Red }
        }
      } catch {
        Write-Host "No se pudo leer el diagnostico JSON: $debugFile" -ForegroundColor Yellow
      }
    }
    throw "Fallo el paso: $File $Arguments"
  }
}

function Wait-EdgeReady {
  param([int]$TimeoutSeconds = 60)
  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    try {
      $tabs = Invoke-RestMethod -Uri "http://127.0.0.1:9222/json" -TimeoutSec 5
      $anyPage = $tabs | Where-Object { $_.type -eq "page" } | Select-Object -First 1
      if ($anyPage) {
        Write-Host "Edge controlable detectado. Se usaran cookies de la sesion para leer SharePoint." -ForegroundColor Green
        return
      }
    } catch {
      Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 2
  }
  throw "No se detecto Edge en el puerto 9222. Ejecuta sin -NoAbrirEdge o abre Edge desde el script."
}

function Ensure-EdgeReady {
  param([string]$Url)
  try {
    Wait-EdgeReady -TimeoutSeconds 8
    return
  } catch {
    Write-Host "Edge controlable no esta disponible. Reabriendo sesion de SharePoint..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $EdgeProfile | Out-Null
    Start-Process -FilePath "msedge.exe" -ArgumentList @(
      "--new-window",
      "--remote-debugging-port=9222",
      "--remote-allow-origins=*",
      "--user-data-dir=$EdgeProfile",
      $Url
    )
    Wait-EdgeReady -TimeoutSeconds 90
  }
}

New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
New-Item -ItemType Directory -Force -Path $OutputsDir | Out-Null

if (!(Test-Path -LiteralPath $Python)) { throw "No encuentro Python en $Python" }
if (!(Test-Path -LiteralPath $Node)) { throw "No encuentro Node en $Node" }

if (-not $NoAbrirEdge) {
  Write-Step "Abriendo Edge con la nueva ruta de soportes"
  New-Item -ItemType Directory -Force -Path $EdgeProfile | Out-Null
  Start-Process -FilePath "msedge.exe" -ArgumentList @(
    "--new-window",
    "--remote-debugging-port=9222",
    "--remote-allow-origins=*",
    "--user-data-dir=$EdgeProfile",
    $SoportesUrl
  )
  Wait-EdgeReady -TimeoutSeconds 90
} else {
  Write-Step "Usando Edge ya abierto en puerto 9222"
  Wait-EdgeReady -TimeoutSeconds 20
}

if ([string]::IsNullOrWhiteSpace($CronogramaPath)) {
  if (-not [string]::IsNullOrWhiteSpace($CronogramaUrl)) {
    Write-Step "Descargando cronograma desde la ruta indicada"
    Invoke-Step $Node @("work\download_sharepoint_cronograma_custom_url.mjs")
  } else {
    Write-Step "Descargando cronograma desde SharePoint original"
    Invoke-Step $Node @("work\download_sharepoint_cronograma.mjs")
  }
  $CronogramaPath = $CronogramaSharePoint
} else {
  Write-Step "Usando cronograma local indicado por parametro"
}

if (!(Test-Path -LiteralPath $CronogramaPath)) { throw "No encuentro el cronograma: $CronogramaPath" }

Write-Step "Copiando cronograma para evitar bloqueos de Excel/OneDrive"
if ((Resolve-Path -LiteralPath $CronogramaPath).Path -ne (Resolve-Path -LiteralPath $CronogramaCopia -ErrorAction SilentlyContinue).Path) {
  Copy-Item -LiteralPath $CronogramaPath -Destination $CronogramaCopia -Force
}
$env:CRONOGRAMA_PATH = $CronogramaCopia

Write-Step "Preparando lectura local sin mezclar ejecuciones"
Remove-Item -LiteralPath (Join-Path $WorkDir "pdf_text") -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $WorkDir "sharepoint_pdf_manifest.json") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $WorkDir "pdf_text_summary.json") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $WorkDir "local_synced_supports_result.json") -Force -ErrorAction SilentlyContinue

Write-Step "Extrayendo listado de archivos desde la nueva ruta"
Ensure-EdgeReady -Url $SoportesUrl
Invoke-Step $Node @("work\cdp_sharepoint_list_custom_url.mjs")

Write-Step "Descargando PDFs de soportes desde la nueva ruta"
Ensure-EdgeReady -Url $SoportesUrl
Invoke-Step $Node @("work\download_sharepoint_pdfs_custom_url.mjs")

Write-Step "Extrayendo texto de PDFs"
Invoke-Step $Python @("work\extract_pdf_text.py")

Write-Step "Extrayendo cronograma esperado y fechas digitadas"
Invoke-Step $Python @("work\extract_schedule.py")

Write-Step "Comparando soportes cargados vs cronograma"
Invoke-Step $Python @("work\compare_sharepoint_supports.py")

Write-Step "Comparando autogestiones subidas al aplicativo"
Invoke-Step $Python @("work\compare_app_uploads.py")

Write-Step "Validando pilares, autogestiones, tecnicos y horas"
Invoke-Step $Python @("work\validate_pillars_hours.py")

Write-Step "Generando dashboard HTML"
Invoke-Step $Python @("work\build_dashboard_pillars.py")

if (!(Test-Path -LiteralPath $Dashboard)) {
  throw "No se genero el dashboard esperado: $Dashboard"
}

Write-Host ""
Write-Host "Dashboard generado correctamente con ruta nueva:" -ForegroundColor Green
Write-Host $Dashboard
Write-Host "Periodo consultado: $Anio / $CarpetaMes" -ForegroundColor Green
Write-Host "Ruta de soportes: $SoportesUrl" -ForegroundColor Green
if (-not [string]::IsNullOrWhiteSpace($CronogramaUrl)) {
  Write-Host "Ruta de cronograma: $CronogramaUrl" -ForegroundColor Green
}
Write-Host ""
Write-Host "Para abrirlo:" -ForegroundColor Yellow
Write-Host "start `"$Dashboard`""


