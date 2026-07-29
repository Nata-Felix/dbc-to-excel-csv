$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$DuckDbVersion = '1.5.5'
$DuckDbArchiveSha256 = '8375EB1FCF2212E8A0817950354815D4DDE9DD383C2D9FA7B8975B71E278C1BD'
$BuildDirectory = Join-Path $ProjectRoot 'build\windows-x64'
$PayloadDirectory = Join-Path $ProjectRoot 'build\payload'
$DependencyDirectory = Join-Path $ProjectRoot 'build\dependencies\duckdb'
$DistributionDirectory = Join-Path $ProjectRoot 'dist'
$DuckDbArchive = Join-Path $ProjectRoot 'build\dependencies\libduckdb-windows-amd64.zip'

$Cxx = Get-Command 'x86_64-w64-mingw32-clang++.exe' -ErrorAction SilentlyContinue
$Cc = Get-Command 'x86_64-w64-mingw32-clang.exe' -ErrorAction SilentlyContinue
$Windres = Get-Command 'x86_64-w64-mingw32-windres.exe' -ErrorAction SilentlyContinue
if (-not $Cxx -or -not $Cc -or -not $Windres) {
    throw 'LLVM-MinGW nao encontrado no PATH.'
}

New-Item -ItemType Directory -Force -Path $BuildDirectory,$PayloadDirectory,$DependencyDirectory,$DistributionDirectory | Out-Null

$DuckDbLibrary = Join-Path $DependencyDirectory 'duckdb.lib'
$DuckDbDll = Join-Path $DependencyDirectory 'duckdb.dll'
if (-not (Test-Path $DuckDbLibrary) -or -not (Test-Path $DuckDbDll)) {
    $DownloadUrl = "https://github.com/duckdb/duckdb/releases/download/v$DuckDbVersion/libduckdb-windows-amd64.zip"
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $DuckDbArchive
    $ActualHash = (Get-FileHash $DuckDbArchive -Algorithm SHA256).Hash
    if ($ActualHash -ne $DuckDbArchiveSha256) {
        throw "SHA-256 inesperado para DuckDB: $ActualHash"
    }
    Expand-Archive -LiteralPath $DuckDbArchive -DestinationPath $DependencyDirectory -Force
}

& $Cc.Source -O2 -DNDEBUG -I (Join-Path $ProjectRoot 'third_party\blast') `
    -c (Join-Path $ProjectRoot 'third_party\blast\blast.c') `
    -o (Join-Path $BuildDirectory 'blast.o')
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o descompressor BLAST.' }

Push-Location (Join-Path $ProjectRoot 'resources')
try {
    & $Windres.Source -i 'app.rc' -o (Join-Path $BuildDirectory 'gui_resources.o')
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar os recursos da interface.' }
} finally {
    Pop-Location
}

$CoreExecutable = Join-Path $PayloadDirectory 'DBC_Converter_Core.exe'
& $Cxx.Source -std=c++17 -O2 -DNDEBUG -Wall -Wextra -Wpedantic `
    -static -static-libgcc -static-libstdc++ -DDBC_GUI -municode -mwindows `
    -I (Join-Path $ProjectRoot 'third_party\blast') `
    -I (Join-Path $ProjectRoot 'third_party\duckdb') `
    (Join-Path $ProjectRoot 'src\main.cpp') `
    (Join-Path $BuildDirectory 'blast.o') $DuckDbLibrary `
    (Join-Path $BuildDirectory 'gui_resources.o') `
    -lcomctl32 -lcomdlg32 -lshell32 -lole32 -o $CoreExecutable
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o conversor.' }
Copy-Item -LiteralPath $DuckDbDll -Destination (Join-Path $PayloadDirectory 'duckdb.dll') -Force

Push-Location (Join-Path $ProjectRoot 'resources')
try {
    & $Windres.Source -i 'launcher.rc' -o (Join-Path $BuildDirectory 'launcher_resources.o')
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao incorporar o pacote no executavel.' }
} finally {
    Pop-Location
}

$FinalExecutable = Join-Path $DistributionDirectory 'DBC-to-Excel-CSV.exe'
& $Cxx.Source -std=c++17 -O2 -DNDEBUG -Wall -Wextra -Wpedantic `
    -static -static-libgcc -static-libstdc++ -municode -mwindows `
    (Join-Path $ProjectRoot 'src\launcher.cpp') `
    (Join-Path $BuildDirectory 'launcher_resources.o') `
    -lshell32 -o $FinalExecutable
if ($LASTEXITCODE -ne 0) { throw 'Falha ao gerar o executavel portatil.' }

$FinalHash = (Get-FileHash $FinalExecutable -Algorithm SHA256).Hash
Write-Host "Gerado: $FinalExecutable"
Write-Host "SHA-256: $FinalHash"
