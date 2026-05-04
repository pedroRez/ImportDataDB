param(
    [switch]$SkipInstaller,
    [switch]$NoClean
)

$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
Set-Location $projectRoot
$venvPython = Join-Path $projectRoot '.venv\Scripts\python.exe'
$pythonExe = if (Test-Path $venvPython) { $venvPython } else { 'python' }

if (-not $NoClean) {
    $pathsToClean = @(
        (Join-Path $projectRoot 'build\ImportDataDB'),
        (Join-Path $projectRoot 'dist\ImportDataDB')
    )
    foreach ($path in $pathsToClean) {
        $fullPath = [System.IO.Path]::GetFullPath($path)
        $rootPath = [System.IO.Path]::GetFullPath($projectRoot)
        if ($fullPath.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-Path $fullPath)) {
            Remove-Item -LiteralPath $fullPath -Recurse -Force
        }
    }
    $outputDir = Join-Path $projectRoot 'output'
    if (Test-Path $outputDir) {
        Get-ChildItem -LiteralPath $outputDir -Filter 'ImportDataDB-Setup*.exe' | Remove-Item -Force
    }
}

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
& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r requirements.txt pyinstaller

Write-Host '==> Gerando executável com PyInstaller...' -ForegroundColor Cyan
& $pythonExe -m PyInstaller @pyInstallerArgs

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
