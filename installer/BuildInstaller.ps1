param(
    [string]$Configuration = "Release",
    [string]$BuildOutputDir
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$issPath = Join-Path $PSScriptRoot "Project_bpi.iss"

if (-not $BuildOutputDir) {
    $candidateRelease = Join-Path $repoRoot "Project_bpi\bin\$Configuration"
    $candidateDebug = Join-Path $repoRoot "Project_bpi\bin\Debug"

    if (Test-Path (Join-Path $candidateRelease "Project_bpi.exe")) {
        $BuildOutputDir = $candidateRelease
    }
    elseif (Test-Path (Join-Path $candidateDebug "Project_bpi.exe")) {
        $BuildOutputDir = $candidateDebug
    }
    else {
        throw "Не найден готовый файл Project_bpi.exe. Сначала соберите приложение."
    }
}

$iscc = (Get-Command ISCC.exe -ErrorAction SilentlyContinue).Source
if (-not $iscc) {
    $defaultIscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if (Test-Path $defaultIscc) {
        $iscc = $defaultIscc
    }
}

if (-not $iscc) {
    throw "Не найден ISCC.exe. Установите Inno Setup 6."
}

$buildOutputDirFull = [System.IO.Path]::GetFullPath($BuildOutputDir)

& $iscc "/DBuildOutputDir=$buildOutputDirFull" $issPath
