param(
    [switch]$SkipInstaller,
    [switch]$NoClean
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
Set-Location $projectRoot

$venvPython = Join-Path $projectRoot '.venv\Scripts\python.exe'
$pythonExe = if (Test-Path $venvPython) { $venvPython } else { 'python' }
$appVersion = (& $pythonExe -c "from src.version import __version__; print(__version__)").Trim()

$pyInstallerArgs = @(
    '--noconfirm',
    '--windowed',
    '--name', 'ImportDataDB',
    '--paths', $projectRoot,
    '--paths', (Join-Path $projectRoot 'src'),
    'app/__main__.py'
)

if (-not $NoClean) {
    $pyInstallerArgs = @('--clean') + $pyInstallerArgs
}

Write-Host '==> Instalando dependencias de build (PyInstaller)...' -ForegroundColor Cyan
& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r requirements.txt pyinstaller

Write-Host "==> Gerando executavel com PyInstaller (versao $appVersion)..." -ForegroundColor Cyan
& $pythonExe -m PyInstaller @pyInstallerArgs

if ($SkipInstaller) {
    Write-Host '==> Build concluido. Instalador foi pulado por -SkipInstaller.' -ForegroundColor Yellow
    exit 0
}

$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
    $defaultIsccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
    if (Test-Path $defaultIsccPath) {
        $isccPath = $defaultIsccPath
    }
    else {
        throw 'Inno Setup (ISCC.exe) nao encontrado. Instale o Inno Setup 6 ou rode com -SkipInstaller.'
    }
}
else {
    $isccPath = $iscc.Source
}

Write-Host '==> Gerando instalador (Inno Setup)...' -ForegroundColor Cyan
& $isccPath "/DMyAppVersion=$appVersion" 'installer/ImportDataDB.iss'

Write-Host "==> Concluido. Verifique a pasta output/ para o instalador da versao $appVersion." -ForegroundColor Green
