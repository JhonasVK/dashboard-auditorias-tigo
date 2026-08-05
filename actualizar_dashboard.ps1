# Actualiza el dashboard TIGO a partir del Excel y lo publica en GitHub Pages.
# Uso: doble clic en "Actualizar_Dashboard.bat", o ejecutar este script en PowerShell.

$ErrorActionPreference = "Stop"

$ExcelPath = "C:\Users\jvodn\Desktop\Capacitacion\Cerificacion\TIGO\Auditorias_en_Terreno\Dashboard_Auditorias_TIGO.xlsx"
$RepoDir   = "C:\Users\jvodn\Downloads\Web\dashboard-auditorias-tigo"
$TemplatePath = Join-Path $RepoDir "template.html"
$OutputPath   = Join-Path $RepoDir "index.html"

Write-Host "==> Leyendo Excel: $ExcelPath"
if (-not (Test-Path $ExcelPath)) {
    Write-Host "ERROR: no se encontró el archivo Excel en esa ruta." -ForegroundColor Red
    exit 1
}

# ---------- 1. Extraer el .xlsx (es un .zip) ----------
$extractDir = Join-Path $env:TEMP "tigo_dashboard_extract"
if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($ExcelPath, $extractDir)

# ---------- 2. Shared strings ----------
[xml]$ss = New-Object System.Xml.XmlDocument
$ss.Load("$extractDir\xl\sharedStrings.xml")
$strings = @()
foreach ($si in $ss.sst.si) {
    if ($si.t -ne $null) {
        if ($si.t -is [string]) { $strings += $si.t } else { $strings += $si.t.'#text' }
    } elseif ($si.r) {
        $txt = ($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join ''
        $strings += $txt
    } else { $strings += '' }
}

function Col-ToNum($colRef) {
    $col = $colRef -replace '[0-9]', ''
    $num = 0
    foreach ($c in $col.ToCharArray()) { $num = $num * 26 + ([int][char]$c - [int][char]'A' + 1) }
    return $num
}

# ---------- 3. Encontrar la hoja "Datos" dinámicamente ----------
[xml]$wb = New-Object System.Xml.XmlDocument
$wb.Load("$extractDir\xl\workbook.xml")
$datosSheet = $wb.workbook.sheets.sheet | Where-Object { $_.name -eq 'Datos' }
if (-not $datosSheet) { Write-Host "ERROR: no se encontró la hoja 'Datos' en el Excel." -ForegroundColor Red; exit 1 }
$datosRid = ($datosSheet.Attributes | Where-Object { $_.Name -eq 'r:id' }).Value
[xml]$wbRels = New-Object System.Xml.XmlDocument
$wbRels.Load("$extractDir\xl\_rels\workbook.xml.rels")
$rel = $wbRels.Relationships.Relationship | Where-Object { $_.Id -eq $datosRid }
if (-not $rel) { Write-Host "ERROR: no se pudo resolver la relación de la hoja 'Datos'." -ForegroundColor Red; exit 1 }
$sheetFile = Join-Path "$extractDir\xl" $rel.Target

Write-Host "==> Procesando hoja 'Datos' ($sheetFile)"

[xml]$sheet = New-Object System.Xml.XmlDocument
$sheet.Load($sheetFile)
$rows = @()
foreach ($row in $sheet.worksheet.sheetData.row) {
    $rowData = @{}
    if ($row.c) {
        foreach ($c in $row.c) {
            $colNum = Col-ToNum $c.r
            $val = $null
            if ($c.t -eq 's') { if ($c.v -ne $null) { $val = $strings[[int]$c.v] } }
            elseif ($c.t -eq 'str' -or $c.t -eq 'inlineStr') { if ($c.is) { $val = $c.is.t } else { $val = $c.v } }
            else { $val = $c.v }
            $rowData[$colNum] = $val
        }
    }
    $rows += ,$rowData
}

$keys = @('fecha','mes','auditor','tecnico','peticion','actividad','zona','reutiliza','altasFtth','altasFtta','pasamuros','reutilizaRoseta','medicionElectrica','rotulaCto','areaLimpia','clienteConforme','capacitacion','servicioFunciona','entregaProtocolo','notaEstetica','observaciones')

$objs = @()
for ($i = 1; $i -lt $rows.Count; $i++) {
    $r = $rows[$i]
    # saltar filas totalmente vacías
    if ($r.Count -eq 0) { continue }
    $obj = [ordered]@{}
    for ($c = 1; $c -le 21; $c++) {
        $k = $keys[$c - 1]
        $v = if ($r.ContainsKey($c)) { $r[$c] } else { '' }
        if ($k -eq 'notaEstetica') {
            $num = 0.0
            [double]::TryParse($v, [ref]$num) | Out-Null
            $obj[$k] = $num
        } else {
            $obj[$k] = ("$v").Trim()
        }
    }
    $objs += ,([pscustomobject]$obj)
}

Write-Host "==> $($objs.Count) auditorías leídas"

$json = $objs | ConvertTo-Json -Compress
if ($json -match '</script') { Write-Host "ERROR: los datos contienen una secuencia insegura, abortando." -ForegroundColor Red; exit 1 }

# ---------- 4. Inyectar en la plantilla ----------
if (-not (Test-Path $TemplatePath)) { Write-Host "ERROR: no se encontró template.html en $RepoDir" -ForegroundColor Red; exit 1 }
$html = Get-Content $TemplatePath -Raw -Encoding UTF8
$final = $html.Replace('/*__DATA__*/', $json)
[System.IO.File]::WriteAllText($OutputPath, $final, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "==> index.html regenerado ($((Get-Item $OutputPath).Length) bytes)"

# ---------- 5. Commit y push ----------
Push-Location $RepoDir
try {
    git add index.html | Out-Null
    $diff = git diff --cached --name-only
    if (-not $diff) {
        Write-Host "==> No hay cambios respecto a la última publicación. Nada que subir."
    } else {
        $fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
        git commit -q -m "Actualizar datos de auditorías ($fecha)"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: el commit falló. Revisa el mensaje de git de arriba (por ejemplo, identidad de git no configurada)." -ForegroundColor Red
            exit 1
        }
        git push -q
        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERROR: el push a GitHub falló. Revisa tu conexión o sesión de 'gh auth status'." -ForegroundColor Red
            exit 1
        }
        Write-Host "==> Publicado. El sitio se actualizará en 1-2 minutos en:"
        Write-Host "    https://jhonasvk.github.io/dashboard-auditorias-tigo/" -ForegroundColor Cyan
    }
} finally {
    Pop-Location
}
