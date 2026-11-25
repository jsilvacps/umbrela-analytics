$ErrorActionPreference = "Stop"

$Projeto = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Projeto

Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " Umbrela Analytics Enterprise - Gerador Setup " -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

Write-Host "Instalando dependencias Python..." -ForegroundColor Yellow
py -m pip install --upgrade pip
py -m pip install -r requirements.txt

# Extrai versão do script principal
$AppVersion = (Select-String -Path "analisador_macro_lock.py" -Pattern 'APP_VERSION\s*=\s*"([\d.]+)"').Matches[0].Groups[1].Value
Write-Host "Versao detectada: $AppVersion" -ForegroundColor Cyan

Write-Host "Limpando builds antigos..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item "build" -Recurse -Force }
if (Test-Path "dist") { Remove-Item "dist" -Recurse -Force }
if (Test-Path "UmbrelaAnalytics.spec") { Remove-Item "UmbrelaAnalytics.spec" -Force }
# Nao remove pasta installer inteira (pode estar sendo sincronizada pelo Drive)

Write-Host "Gerando executavel..." -ForegroundColor Yellow
py -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name "UmbrelaAnalytics" `
  --icon "logo_umbrela_analytics.ico" `
  --add-data "logo_umbrela_analytics.png;." `
  --hidden-import uuid `
  --hidden-import duckdb `
  --exclude-module matplotlib `
  --exclude-module numpy `
  --exclude-module pandas `
  --exclude-module PIL `
  --exclude-module scipy `
  --exclude-module IPython `
  --exclude-module jupyter `
  --exclude-module notebook `
  --exclude-module pytest `
  --exclude-module cryptography `
  --exclude-module docutils `
  --exclude-module sphinx `
  "analisador_macro_lock.py"

if (!(Test-Path "dist\UmbrelaAnalytics.exe")) {
    throw "Falha ao gerar dist\UmbrelaAnalytics.exe"
}

Write-Host "Procurando Inno Setup..." -ForegroundColor Yellow
$PossiveisISCC = @(
  "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
  "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
  "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)

$ISCC = $PossiveisISCC | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $ISCC) {
    Write-Host "Inno Setup nao encontrado. Tentando instalar via winget..." -ForegroundColor Yellow
    winget install --id JRSoftware.InnoSetup -e --accept-package-agreements --accept-source-agreements
    $ISCC = $PossiveisISCC | Where-Object { Test-Path $_ } | Select-Object -First 1
}

if (-not $ISCC) {
    throw "Inno Setup nao foi encontrado. Instale manualmente o Inno Setup 6 e rode novamente."
}

Write-Host "Gerando instalador Setup v$AppVersion..." -ForegroundColor Green
$TempOut = "$env:TEMP\UmbrelaSetupOut"
if (Test-Path $TempOut) { Remove-Item $TempOut -Recurse -Force }
New-Item -ItemType Directory -Force $TempOut | Out-Null

& $ISCC /DMyAppVersion="$AppVersion" /DMyOutputDir="$TempOut" "UmbrelaAnalytics_Setup.iss"
if ($LASTEXITCODE -ne 0) { throw "Falha na compilacao do Inno Setup." }

# Copia resultado para pasta installer no Drive
if (!(Test-Path "installer")) { New-Item -ItemType Directory -Force "installer" | Out-Null }
$SetupName = "UmbrelaAnalytics_Setup_v$AppVersion.exe"
Copy-Item "$TempOut\$SetupName" "installer\$SetupName" -Force
Remove-Item $TempOut -Recurse -Force

$SetupFile = "installer\$SetupName"
Write-Host ""
Write-Host "FINALIZADO COM SUCESSO!" -ForegroundColor Green
Write-Host "Setup gerado em: $Projeto\$SetupFile" -ForegroundColor Cyan
