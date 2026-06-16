param(
    [Parameter(Mandatory=$false)]
    [string]$ScriptId
)

$ErrorActionPreference = "Stop"

# 1. Pedir ID si no se pasa como parámetro
if (-not $ScriptId) {
    $ScriptId = Read-Host "Por favor ingresa el ID del proyecto de Google Apps Script"
}

$ProjectDir = "C:\proyectos\WorkScripts\G_Apps_Scripts"
if (-not (Test-Path $ProjectDir)) {
    Write-Host "La ruta $ProjectDir no existe." -ForegroundColor Red
    exit
}

Set-Location $ProjectDir

# 2. Iniciar sesión en clasp
Write-Host "Iniciando sesión en clasp..." -ForegroundColor Cyan
# Nota: Si ya estás logueado, esto igual validará la sesión.
npx clasp login

# 3. Pedir la subcarpeta a desplegar
$Subfolder = Read-Host "Ingresa el nombre de la subcarpeta del proyecto a desplegar (ej: BD_Clientes_Pedidos_06-26)"
$TargetDir = Join-Path $ProjectDir $Subfolder

if (-not (Test-Path $TargetDir)) {
    Write-Host "La subcarpeta $Subfolder no existe. Asegúrate de crearla y tener el código allí primero." -ForegroundColor Red
    exit
}

Set-Location $TargetDir

# 4. Actualizar .clasp.json con el ID nuevo
Write-Host "Configurando el ID del script en .clasp.json..." -ForegroundColor Cyan
$claspConfig = @{
    scriptId = $ScriptId
    rootDir = $TargetDir
} | ConvertTo-Json
Set-Content -Path ".clasp.json" -Value $claspConfig

# 4.5 Configurar ID de Google Sheets
$SheetId = Read-Host "Ingresa el ID de la hoja de cálculo de Google Sheets (la base de datos)"
if ($SheetId) {
    Write-Host "Inyectando ID de Google Sheets en el código..." -ForegroundColor Cyan
    $CodeJsPath = Join-Path $TargetDir "Code.js"
    if (Test-Path $CodeJsPath) {
        $codeContent = Get-Content $CodeJsPath -Raw
        $codeContent = $codeContent -replace 'SpreadsheetApp\.getActiveSpreadsheet\(\)', "SpreadsheetApp.openById(`"$SheetId`")"
        $codeContent = $codeContent -replace 'SpreadsheetApp\.openById\(["''].*?["'']\)', "SpreadsheetApp.openById(`"$SheetId`")"
        Set-Content -Path $CodeJsPath -Value $codeContent
    }
}

# 5. Push y Deploy
Write-Host "Subiendo el código a Google Apps Script (clasp push)..." -ForegroundColor Cyan
npx clasp push --force

Write-Host "Creando el despliegue web (clasp deploy)..." -ForegroundColor Cyan
$deployOutput = npx clasp deploy 2>&1
Write-Host $deployOutput

# 6. Capturar el ID del despliegue con una expresión regular
$deploymentId = $null
# clasp deploy suele imprimir algo como: "Deployed AKfy... @1." o "- AKfy... @1."
if ($deployOutput -match "Deployed\s+([A-Za-z0-9_-]+)\s+@") {
    $deploymentId = $matches[1]
} elseif ($deployOutput -match "-\s+([A-Za-z0-9_-]+)\s+@") {
    $deploymentId = $matches[1]
}

if (-not $deploymentId) {
    Write-Host "No se pudo extraer la URL del despliegue automáticamente de la salida de clasp." -ForegroundColor Red
    exit
}

$AppUrl = "https://script.google.com/macros/s/$deploymentId/exec"
Write-Host ""
Write-Host "✅ URL del Web App obtenida con éxito:" -ForegroundColor Green
Write-Host $AppUrl -ForegroundColor Green
Write-Host ""

# 7. Interacción con el usuario para inyectar en React
$JsxPath = Read-Host "Por favor ingresa la ruta COMPLETA al archivo de React (.jsx) donde deseas inyectar la URL (ej: C:\proyectos\...\App.jsx)"

if (Test-Path $JsxPath) {
    $content = Get-Content $JsxPath -Raw
    
    # Buscar el patrón de la constante (soporta comillas simples o dobles)
    $pattern = 'const\s+APPS_SCRIPT_URL\s*=\s*["''].*?["'']'
    $replacement = "const APPS_SCRIPT_URL = `"$AppUrl`""
    
    if ($content -match $pattern) {
        $content = $content -replace $pattern, $replacement
        Set-Content -Path $JsxPath -Value $content
        Write-Host "¡La URL ha sido inyectada correctamente en $JsxPath!" -ForegroundColor Green
    } else {
        Write-Host "⚠️ No se encontró la variable 'const APPS_SCRIPT_URL = ...' en el archivo." -ForegroundColor Yellow
        Write-Host "Por favor, pégala manualmente: const APPS_SCRIPT_URL = `"$AppUrl`";" -ForegroundColor Yellow
    }
} else {
    Write-Host "La ruta especificada no existe." -ForegroundColor Red
}

# 8. Automatizar la adición en el AdminForm.html
Write-Host ""
Write-Host "Actualizando el menú desplegable del Formulario Admin..." -ForegroundColor Cyan
$AdminFormPath = "C:\proyectos\WorkScripts\OCR_Catalogos-Dash_Pedidos\Integracion\AdminForm.html"

if (Test-Path $AdminFormPath) {
    $adminContent = Get-Content $AdminFormPath -Raw
    # Verificamos si la URL ya existe para no duplicar
    if (-not ($adminContent -match [regex]::Escape($AppUrl))) {
        # Insertamos la nueva etiqueta <option> justo antes de cerrar el </select>
        $newOption = "<option value=`"$AppUrl`">$Subfolder</option>`n                    </select>"
        $adminContent = $adminContent -replace '</select>', $newOption
        Set-Content -Path $AdminFormPath -Value $adminContent
        Write-Host "¡Opción '$Subfolder' agregada automáticamente a AdminForm.html!" -ForegroundColor Green
    } else {
        Write-Host "El proyecto ya existe en el menú de AdminForm.html." -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️ No se encontró el archivo AdminForm.html en la ruta esperada." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=====================================================================" -ForegroundColor Yellow
Write-Host "⚠️ RECORDATORIO IMPORTANTE PARA NUEVOS PROYECTOS:" -ForegroundColor Yellow
Write-Host "Si este es un proyecto nuevo o una nueva base de datos, debes:" -ForegroundColor Yellow
Write-Host "1. Abrir el proyecto en Google Apps Script (https://script.google.com)." -ForegroundColor Yellow
Write-Host "2. Seleccionar la función 'inscribirCliente' o cualquier otra y presionar 'Ejecutar' (Run)." -ForegroundColor Yellow
Write-Host "3. Seguir el asistente para autorizar los permisos de acceso a Google Sheets." -ForegroundColor Yellow
Write-Host "¡Sin este paso, el Dashboard y el Formulario Admin darán error de conexión!" -ForegroundColor Red
Write-Host "=====================================================================" -ForegroundColor Yellow

Write-Host ""
Write-Host "🎉 Proceso completado." -ForegroundColor Cyan
