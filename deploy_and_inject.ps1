param(
    [Parameter(Mandatory=$false)]
    [string]$ScriptId
)

$ErrorActionPreference = "Stop"

$ProjectDir = "C:\proyectos\WorkScripts\G_Apps_Scripts"
if (-not (Test-Path $ProjectDir)) {
    Write-Host "La ruta $ProjectDir no existe." -ForegroundColor Red
    exit
}

Set-Location $ProjectDir

# 1. Iniciar sesión en clasp


# 1.5 Leer .env global por bloques de Proyecto
$GlobalEnvPath = Join-Path $ProjectDir ".env"
$EnvDict = @{} # Para buscar valores generales
$ProjectEnv = @{} # Para agrupar variables por carpeta del proyecto

if (Test-Path $GlobalEnvPath) {
    $envContent = Get-Content $GlobalEnvPath
    $currentProject = $null
    
    foreach ($line in $envContent) {
        if ($line -match "^#\s*Proyecto:\s*(.*)$") {
            $currentProject = $matches[1].Trim()
            $ProjectEnv[$currentProject] = @{}
        } elseif ($line -match "^([^#=]+)=(.*)$") {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $EnvDict[$key] = $val
            if ($currentProject) {
                $ProjectEnv[$currentProject][$key] = $val
            }
        }
    }
}

# 2. Buscar proyectos de Apps Script
Write-Host "`nBuscando proyectos de Apps Script y analizando .env..." -ForegroundColor Cyan
$projects = @()
$dirs = Get-ChildItem -Path $ProjectDir -Directory | Where-Object { $_.Name -ne "node_modules" -and $_.Name -notmatch "^\." }

foreach ($dir in $dirs) {
    $claspFile = Join-Path $dir.FullName ".clasp.json"
    $folderEnvKey = ($dir.Name -replace '[^a-zA-Z0-9]', '_').ToUpper() + "_SCRIPT_ID"
    $scriptId = $null
    
    if (Test-Path $claspFile) {
        $claspData = Get-Content $claspFile -Raw | ConvertFrom-Json
        $scriptId = $claspData.scriptId
    } elseif ($EnvDict.ContainsKey($folderEnvKey)) {
        $scriptId = $EnvDict[$folderEnvKey]
        $claspConfig = @{ scriptId = $scriptId; rootDir = $dir.FullName } | ConvertTo-Json
        Set-Content -Path $claspFile -Value $claspConfig
        Write-Host " [OK] Auto-configurado $($dir.Name) usando $folderEnvKey desde .env" -ForegroundColor Green
    }
    
    if ($scriptId) {
        $gitStatus = git status --porcelain $dir.FullName
        $statusMsg = "Sin cambios pendientes en Git"
        $statusColor = "Green"
        if ($gitStatus) {
            $statusMsg = "MODIFICADO LOCALMENTE (Revisar si falta clasp push)"
            $statusColor = "Yellow"
        }
    } else {
        $statusMsg = "SIN CONFIGURAR (Requiere ingreso manual de ID)"
        $statusColor = "DarkGray"
        $scriptId = "N/A"
    }

    $projects += [PSCustomObject]@{
        FolderName = $dir.Name
        FullPath = $dir.FullName
        ScriptId = $scriptId
        StatusMsg = $statusMsg
        StatusColor = $statusColor
        EnvKey = $folderEnvKey
    }
}

if ($projects.Count -eq 0) {
    Write-Host "No se encontraron proyectos configurados (con .clasp.json)." -ForegroundColor Red
    $Subfolder = Read-Host "Ingresa el nombre de la subcarpeta a desplegar manualmente"
    $TargetDir = Join-Path $ProjectDir $Subfolder
    if (-not $ScriptId) {
        $ScriptId = Read-Host "Ingresa el ID del proyecto de Google Apps Script"
    }
    
    if (-not (Test-Path $TargetDir)) {
        Write-Host "La subcarpeta $Subfolder no existe." -ForegroundColor Red
        exit
    }
    
    # Crear/Actualizar .clasp.json si es manual
    $claspConfig = @{ scriptId = $ScriptId; rootDir = $TargetDir } | ConvertTo-Json
    Set-Content -Path (Join-Path $TargetDir ".clasp.json") -Value $claspConfig
    
} else {
    Write-Host "`n==================================================" -ForegroundColor Magenta
    Write-Host "PROYECTOS DETECTADOS LISTOS PARA CLASP PUSH:" -ForegroundColor Magenta
    Write-Host " [0] + CONFIGURAR NUEVO PROYECTO MANUALMENTE" -ForegroundColor Yellow
    for ($i = 0; $i -lt $projects.Count; $i++) {
        $p = $projects[$i]
        Write-Host " [$($i+1)] $($p.FolderName)" -ForegroundColor Cyan -NoNewline
        Write-Host " (ID: $($p.ScriptId)) " -NoNewline
        Write-Host "-> $($p.StatusMsg)" -ForegroundColor $p.StatusColor
    }
    Write-Host "==================================================`n" -ForegroundColor Magenta
    
    $selection = Read-Host "Selecciona el numero del proyecto a desplegar (0 para nuevo, o presiona Enter para cancelar)"
    if ([string]::IsNullOrWhiteSpace($selection)) { exit }
    
    if ($selection -eq "0") {
        $Subfolder = Read-Host "Ingresa el nombre de la subcarpeta a configurar (ej: BD_Cupones_06-26)"
        $TargetDir = Join-Path $ProjectDir $Subfolder
        $ScriptId = Read-Host "Ingresa el ID del proyecto de Google Apps Script (scriptId)"
        
        if (-not (Test-Path $TargetDir)) {
            Write-Host "La subcarpeta $Subfolder no existe. Creandola..." -ForegroundColor Yellow
            New-Item -ItemType Directory -Path $TargetDir | Out-Null
        }
        
        $claspConfig = @{ scriptId = $ScriptId; rootDir = $TargetDir } | ConvertTo-Json
        Set-Content -Path (Join-Path $TargetDir ".clasp.json") -Value $claspConfig
        Write-Host "`nHas configurado y seleccionado: $Subfolder" -ForegroundColor Green
    } else {
        $selectedIndex = [int]$selection - 1
        if ($selectedIndex -lt 0 -or $selectedIndex -ge $projects.Count) {
            Write-Host "Selección inválida." -ForegroundColor Red
            exit
        }
        
        $selectedProject = $projects[$selectedIndex]
        $Subfolder = $selectedProject.FolderName
        $TargetDir = $selectedProject.FullPath
        $ScriptId = $selectedProject.ScriptId
        
        if ($ScriptId -eq "N/A" -or -not $ScriptId) {
            Write-Host "`nEl proyecto $Subfolder no está configurado." -ForegroundColor Yellow
            $ScriptId = Read-Host "Ingresa el ID del proyecto de Google Apps Script (scriptId)"
            $claspConfig = @{ scriptId = $ScriptId; rootDir = $TargetDir } | ConvertTo-Json
            Set-Content -Path (Join-Path $TargetDir ".clasp.json") -Value $claspConfig
            
            # Guardar en .env para el futuro
            $EnvLine = "`n$($selectedProject.EnvKey)=$ScriptId"
            Add-Content -Path $GlobalEnvPath -Value $EnvLine
            Write-Host " [OK] SCRIPT_ID guardado permanentemente en .env ($($selectedProject.EnvKey))" -ForegroundColor Cyan
        }
        
        Write-Host "`nHas seleccionado: $Subfolder" -ForegroundColor Green
    }
}

Set-Location $TargetDir

# 4.5 Configurar ID de Google Sheets consultando el bloque del proyecto en .env
$SheetId = $null
$EnvKey = ($Subfolder -replace '[^a-zA-Z0-9]', '_').ToUpper() + "_SHEET_ID"

if ($ProjectEnv.ContainsKey($Subfolder)) {
    # Buscar cualquier variable que termine en _SHEET_ID dentro de este proyecto
    foreach ($key in $ProjectEnv[$Subfolder].Keys) {
        if ($key -match "_SHEET_ID$") {
            $SheetId = $ProjectEnv[$Subfolder][$key]
            Write-Host "ID de Google Sheets recuperado automáticamente desde sección '$Subfolder' en .env: $SheetId" -ForegroundColor Green
            $EnvKey = $key # Usar la clave real encontrada
            break
        }
    }
}

if (-not $SheetId) {
    # Si no hay ID en el .env, verificar si ya está en Code.js
    $CodeJsPath = Join-Path $TargetDir "Code.js"
    if (Test-Path $CodeJsPath) {
        $codeContent = Get-Content $CodeJsPath -Raw
        if ($codeContent -match 'SpreadsheetApp\.openById\(["'']([^"'']+)["'']\)') {
            $SheetId = $matches[1]
            Write-Host "ID de Google Sheets detectado automáticamente en Code.js: $SheetId" -ForegroundColor Green
        }
    }
}

if (-not $SheetId) {
    $SheetId = Read-Host "Ingresa el ID de la hoja de cálculo de Google Sheets (la base de datos)"
}

if ($SheetId) {
    # Guardar en .env central para la próxima vez
    if (-not (Test-Path $GlobalEnvPath)) {
        Set-Content -Path $GlobalEnvPath -Value "# Archivo Central de Variables de Entorno para Google Apps Script`n"
    }
    
    # Comprobar si la clave ya existe para no duplicarla
    $envContent = Get-Content $GlobalEnvPath
    if (-not ($envContent -match "^$EnvKey=")) {
        Add-Content -Path $GlobalEnvPath -Value "$EnvKey=$SheetId"
        Write-Host "ID guardado en $GlobalEnvPath bajo la clave $EnvKey para futuros despliegues." -ForegroundColor Cyan
    }

    Write-Host "Inyectando ID de Google Sheets en el código..." -ForegroundColor Cyan
    $CodeJsPath = Join-Path $TargetDir "Code.js"
    if (Test-Path $CodeJsPath) {
        $codeContent = Get-Content $CodeJsPath -Raw
        $codeContent = $codeContent -replace 'SpreadsheetApp\.getActiveSpreadsheet\(\)', "SpreadsheetApp.openById(`"$SheetId`")"
        $codeContent = $codeContent -replace 'SpreadsheetApp\.openById\(["''].*?["'']\)', "SpreadsheetApp.openById(`"$SheetId`")"
        Set-Content -Path $CodeJsPath -Value $codeContent
    }
}

# 4.6 Asegurar que exista appsscript.json (Requerido por Clasp)
$AppScriptJsonPath = Join-Path $TargetDir "appsscript.json"
if (-not (Test-Path $AppScriptJsonPath)) {
    Write-Host "Creando appsscript.json por defecto (requerido para el push)..." -ForegroundColor Yellow
    $ManifestContent = @{
        timeZone = "America/Bogota"
        dependencies = @{}
        webapp = @{
            executeAs = "USER_DEPLOYING"
            access = "ANYONE_ANONYMOUS"
        }
        exceptionLogging = "STACKDRIVER"
    } | ConvertTo-Json
    Set-Content -Path $AppScriptJsonPath -Value $ManifestContent
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

# 7. Interacción con el usuario para inyectar en todas las rutas asociadas al proyecto
Write-Host ""
Write-Host "Procesando rutas detectadas en la sección del proyecto en .env..." -ForegroundColor Cyan

$hasAdminPath = $false
$pathsToInject = @()

if ($ProjectEnv.ContainsKey($Subfolder)) {
    foreach ($key in $ProjectEnv[$Subfolder].Keys) {
        if ($key -match "_PATH$") {
            $pathValue = $ProjectEnv[$Subfolder][$key]
            if (Test-Path $pathValue) {
                if ($key -match "_ADMIN_PATH$" -and $pathValue -match "AdminForm\.html$") {
                    $AdminFormPath = $pathValue
                    $hasAdminPath = $true
                    Write-Host " [Admin] Actualizando menú en: $AdminFormPath" -ForegroundColor Green
                    $adminContent = Get-Content $AdminFormPath -Raw
                    if (-not ($adminContent -match [regex]::Escape($AppUrl))) {
                        # Inyectar la opción SOLO dentro del <select id="script-url">
                        $pattern = '(<select\s+id="script-url"[^>]*>[\s\S]*?)(</select>)'
                        $replacement = "`$1    <option value=`"$AppUrl`">$Subfolder</option>`n                    `$2"
                        $adminContent = $adminContent -replace $pattern, $replacement
                        Set-Content -Path $AdminFormPath -Value $adminContent
                        Write-Host "  -> ¡Opción '$Subfolder' agregada correctamente!" -ForegroundColor Green
                    } else {
                        Write-Host "  -> El proyecto ya existe en el menú de AdminForm.html." -ForegroundColor Yellow
                    }
                } else {
                    Write-Host " [App/Client] Inyectando variable en: $pathValue" -ForegroundColor Green
                    $content = Get-Content $pathValue -Raw
                    
                    # Patrones soportados: const APPS_SCRIPT_URL = "..." o GAS_URL: '...'
                    $patternConst = 'const\s+APPS_SCRIPT_URL\s*=\s*["''].*?["'']'
                    $replacementConst = "const APPS_SCRIPT_URL = `"$AppUrl`""
                    
                    $patternConfig = 'GAS_URL\s*:\s*["''].*?["'']'
                    $replacementConfig = "GAS_URL: `"$AppUrl`""
                    
                    $matched = $false
                    if ($content -match $patternConst) {
                        $content = $content -replace $patternConst, $replacementConst
                        $matched = $true
                    }
                    if ($content -match $patternConfig) {
                        $content = $content -replace $patternConfig, $replacementConfig
                        $matched = $true
                    }
                    
                    if ($matched) {
                        Set-Content -Path $pathValue -Value $content
                        Write-Host "  -> ¡La URL ha sido inyectada correctamente!" -ForegroundColor Green
                    } else {
                        Write-Host "  -> ⚠️ No se encontró la variable 'const APPS_SCRIPT_URL = ...' en el archivo." -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "⚠️ La ruta especificada para $key no existe: $pathValue" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "⚠️ No se encontró una sección para '# Proyecto: $Subfolder' en el .env." -ForegroundColor Yellow
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
Write-Host "=====================================================================" -ForegroundColor Magenta
Write-Host "PASO FINAL: CONTROL DE VERSIONES (GIT PUSH)" -ForegroundColor Magenta
Write-Host "El orden lógico correcto de despliegue completo es:" -ForegroundColor Cyan
Write-Host " 1. Subir cambios a Google Apps Script (clasp push/deploy) [REALIZADO]"
Write-Host " 2. Inyectar la nueva URL en los repositorios Frontend locales [REALIZADO]"
Write-Host " 3. Subir a GitHub (git push) el repositorio de Apps Script para guardar historial."
Write-Host " 4. Subir a GitHub (git push) los repositorios Frontend modificados para que"
Write-Host "    plataformas como Vercel/Netlify se reconstruyan con la nueva URL."
Write-Host "=====================================================================" -ForegroundColor Magenta
Write-Host ""

$doGitPush = Read-Host "¿Deseas hacer el commit y git push automáticamente ahora a todos los repositorios involucrados? (S/N)"
if ($doGitPush -match '^[sS]') {
    $commitMsg = Read-Host "Ingresa el mensaje para el commit (ej: fix: actualizar BD clientes)"
    if ([string]::IsNullOrWhiteSpace($commitMsg)) {
        $commitMsg = "chore: despliegue de Apps Script y actualizacion de URL"
    }

    # 1. Hacer push del repositorio de Apps Script actual
    Write-Host "`n>> [1/2] Realizando git push en el repositorio de Apps Script..." -ForegroundColor Cyan
    Set-Location $ProjectDir
    $gitRootAppsScript = git rev-parse --show-toplevel 2>$null
    if ($gitRootAppsScript) {
        Set-Location $gitRootAppsScript
        git add .
        git commit -m $commitMsg
        git push
    }

    # 2. Hacer push de los repositorios frontend modificados
    Write-Host "`n>> [2/2] Realizando git push en repositorios Frontend vinculados..." -ForegroundColor Cyan
    $frontendDirs = @()
    if ($ProjectEnv.ContainsKey($Subfolder)) {
        foreach ($key in $ProjectEnv[$Subfolder].Keys) {
            if ($key -match "_PATH$") {
                $pathValue = $ProjectEnv[$Subfolder][$key]
                if (Test-Path $pathValue) {
                    $fileDir = Split-Path $pathValue -Parent
                    Set-Location $fileDir
                    $gitRoot = git rev-parse --show-toplevel 2>$null
                    if ($gitRoot -and $frontendDirs -notcontains $gitRoot) {
                        $frontendDirs += $gitRoot
                    }
                }
            }
        }
    }

    foreach ($repo in $frontendDirs) {
        if ($repo -ne $gitRootAppsScript) {
            Write-Host " -> Sincronizando Frontend en: $repo" -ForegroundColor Green
            Set-Location $repo
            git add .
            git commit -m $commitMsg
            git push
        }
    }
    
    Write-Host "`n🎉 Todos los repositorios han sido sincronizados con GitHub." -ForegroundColor Green
} else {
    Write-Host "`nOperación de Git cancelada. Recuerda hacer commit y push manualmente siguiendo el orden lógico." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🎉 Proceso de despliegue completado." -ForegroundColor Cyan
