# ============================================================
#  Descarga las respuestas del formulario de auditorias TIGO desde
#  Google Drive (Firefox, cuenta supervisionenaccion) y reemplaza
#  BBDD\2-Formulario de auditorias de terreno TIGO (Respuestas).xlsx
#
#  El dashboard lee las columnas POR POSICION, asi que antes de
#  reemplazar se verifica que los encabezados sean los mismos.
#  Guarda el Excel anterior en BBDD\respaldo.
# ============================================================
param([int]$EsperaMaxSegundos = 180)
$ErrorActionPreference = 'Stop'

$SHEET_ID      = '1qxZ1UBqXZkzS0SY4M870vKz2YpzULi7Hvkm6lqTXRz8'
$CUENTA_GOOGLE = 'supervisionenaccion@gmail.com'
$URL_DESCARGA  = "https://docs.google.com/spreadsheets/d/$SHEET_ID/export?format=xlsx&authuser=$CUENTA_GOOGLE"
$NOMBRE        = '2-Formulario de auditorias de terreno TIGO (Respuestas)'
$DESTINO       = Join-Path $PSScriptRoot "BBDD\$NOMBRE.xlsx"
$DIR_RESPALDO  = Join-Path $PSScriptRoot 'BBDD\respaldo'
$RESPALDOS_A_GUARDAR = 15

function Falla($msg) { Write-Host "[ERROR] $msg"; exit 1 }

# Encabezados de la primera hoja de un .xlsx (se lee el zip, sin abrir Excel)
function Leer-Encabezados($ruta) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($ruta)
    try {
        function LeerXml($nombre) {
            $e = $zip.GetEntry($nombre); if (-not $e) { return $null }
            $sr = New-Object IO.StreamReader($e.Open()); try { [xml]$sr.ReadToEnd() } finally { $sr.Close() }
        }
        $strings = @()
        $ss = LeerXml 'xl/sharedStrings.xml'
        if ($ss) { foreach ($si in $ss.sst.si) { $strings += ($si.InnerText) } }
        $wb = LeerXml 'xl/workbook.xml'
        $hoja1 = $wb.workbook.sheets.sheet | Select-Object -First 1
        $rid = ($hoja1.Attributes | Where-Object { $_.Name -eq 'r:id' }).Value
        $rels = LeerXml 'xl/_rels/workbook.xml.rels'
        $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target -replace '^/?xl/', ''
        $sh = LeerXml "xl/$target"
        $fila1 = $sh.worksheet.sheetData.row | Select-Object -First 1
        $enc = @()
        foreach ($c in $fila1.c) {
            if ($c.t -eq 's') { $enc += $strings[[int]$c.v] }
            elseif ($c.t -eq 'inlineStr') { $enc += $c.is.InnerText }
            else { $enc += "$($c.v)" }
        }
        return ,$enc
    } finally { $zip.Dispose() }
}

# -- 1. Descargar con Firefox ---------------------------------------
$firefox = @(
    (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\firefox.exe' -ErrorAction SilentlyContinue).'(default)',
    "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
    "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $firefox) { Falla "No se encontro Firefox (debe tener sesion de $CUENTA_GOOGLE)." }

$descargas = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$inicio = Get-Date
Write-Host "      Abriendo la descarga en Firefox ($CUENTA_GOOGLE)..."
Start-Process $firefox -ArgumentList $URL_DESCARGA

$archivo = $null
$limite = $inicio.AddSeconds($EsperaMaxSegundos)
while (-not $archivo -and (Get-Date) -lt $limite) {
    Start-Sleep -Seconds 2
    $nuevo = Get-ChildItem -Path $descargas -Filter "$NOMBRE*.xlsx" -File -ErrorAction SilentlyContinue |
             Where-Object { $_.LastWriteTime -gt $inicio -and $_.Length -gt 0 } |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
    # Firefox crea un archivo vacio y escribe en un ".part": esperar a que termine
    $enCurso = Get-ChildItem -Path $descargas -Filter '*.part' -File -ErrorAction SilentlyContinue |
               Where-Object { $_.LastWriteTime -gt $inicio }
    if ($nuevo -and -not $enCurso) {
        $tam = $nuevo.Length
        Start-Sleep -Seconds 2
        $nuevo.Refresh()
        if ($nuevo.Length -eq $tam) {
            try { $fs = [IO.File]::Open($nuevo.FullName, 'Open', 'Read', 'None'); $fs.Close(); $archivo = $nuevo.FullName } catch {}
        }
    }
}
if (-not $archivo) { Falla "No llego la descarga a '$descargas' en $EsperaMaxSegundos s. Revisa la sesion de $CUENTA_GOOGLE en Firefox." }
Write-Host "      Descargado: $(Split-Path $archivo -Leaf) ($([math]::Round((Get-Item $archivo).Length/1KB)) KB)"

# -- 2. Validar que las columnas no hayan cambiado --------------------
try { $encNuevo = Leer-Encabezados $archivo } catch { Falla "La descarga no es un Excel valido: $($_.Exception.Message)" }
if ("$($encNuevo[0])".Trim() -ne 'Marca temporal') { Falla 'La descarga no tiene "Marca temporal" en A1 (no es la hoja de respuestas).' }
if (Test-Path $DESTINO) {
    $encActual = Leer-Encabezados $DESTINO
    $n = [math]::Min($encActual.Count, $encNuevo.Count)
    $distintas = @()
    for ($i = 0; $i -lt $n; $i++) { if ("$($encActual[$i])".Trim() -ne "$($encNuevo[$i])".Trim()) { $distintas += "col $($i+1): '$($encActual[$i])' -> '$($encNuevo[$i])'" } }
    if ($encNuevo.Count -lt $encActual.Count) { $distintas += "faltan columnas ($($encNuevo.Count) de $($encActual.Count))" }
    if ($distintas.Count) {
        Falla "El formulario cambio de columnas; no se reemplaza la base (el dashboard lee por posicion): $($distintas -join ' | ')"
    }
}

# -- 3. Respaldar y reemplazar -------------------------------------
if (Test-Path $DESTINO) {
    if (-not (Test-Path $DIR_RESPALDO)) { New-Item -ItemType Directory $DIR_RESPALDO | Out-Null }
    Copy-Item $DESTINO (Join-Path $DIR_RESPALDO ("{0}_{1:yyyy-MM-dd_HHmmss}.xlsx" -f $NOMBRE, (Get-Date)))
}
Copy-Item $archivo $DESTINO -Force
Remove-Item $archivo -ErrorAction SilentlyContinue
Get-ChildItem $DIR_RESPALDO -Filter '*.xlsx' -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending | Select-Object -Skip $RESPALDOS_A_GUARDAR | Remove-Item -ErrorAction SilentlyContinue
Write-Host '      Base TIGO reemplazada'
exit 0
