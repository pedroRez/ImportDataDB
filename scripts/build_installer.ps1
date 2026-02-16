param(
    [switch]$SkipInstaller,
    [switch]$NoClean
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
Set-Location $projectRoot

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

Write-Host '==> Instalando dependências de build (PyInstaller)...' -ForegroundColor Cyan
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

Write-Host '==> Gerando executável com PyInstaller...' -ForegroundColor Cyan
python -m PyInstaller @pyInstallerArgs

if ($SkipInstaller) {
    Write-Host '==> Build concluído. Instalador foi pulado por -SkipInstaller.' -ForegroundColor Yellow
    exit 0
}

$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
    $defaultIsccPath = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
    if (Test-Path $defaultIsccPath) {
        $isccPath = $defaultIsccPath
    }
    else {
        throw 'Inno Setup (ISCC.exe) não encontrado. Instale o Inno Setup 6 ou rode com -SkipInstaller.'
    }
}
else {
    $isccPath = $iscc.Source
}

Write-Host '==> Gerando instalador (Inno Setup)...' -ForegroundColor Cyan
& $isccPath 'installer/ImportDataDB.iss'

Write-Host '==> Concluído. Verifique a pasta output/ para o instalador.' -ForegroundColor Green
