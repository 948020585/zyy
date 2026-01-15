Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot

$cscCandidates = @(
  "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe",
  "$env:WINDIR/Microsoft.NET/Framework/v4.0.30319/csc.exe"
)

$csc = $cscCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) {
  throw "csc.exe not found. Install .NET Framework developer tools."
}

$srcDir = Join-Path -Path $scriptRoot -ChildPath "src"
$distDir = Join-Path -Path $scriptRoot -ChildPath "dist"
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$sources = Get-ChildItem -LiteralPath $srcDir -Filter "*.cs" -File | Select-Object -ExpandProperty FullName
if (-not $sources -or $sources.Count -eq 0) {
  throw "No .cs files found under: $srcDir"
}

$commonArgs = @(
  "/nologo",
  "/target:winexe",
  "/optimize+",
  "/reference:System.dll",
  "/reference:System.Core.dll",
  "/reference:System.Data.dll",
  "/reference:System.Windows.Forms.dll",
  "/reference:System.Drawing.dll"
)

$outX64 = Join-Path -Path $distDir -ChildPath "CertPhotoSorter_x64.exe"
$argsX64 = $commonArgs + @("/platform:x64", "/out:$outX64") + $sources
& $csc @argsX64

$outX86 = Join-Path -Path $distDir -ChildPath "CertPhotoSorter_x86.exe"
$argsX86 = $commonArgs + @("/platform:x86", "/out:$outX86") + $sources
& $csc @argsX86

Write-Host "Build done:"
Write-Host " - $outX64"
Write-Host " - $outX86"
