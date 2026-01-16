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
$libDir = Join-Path -Path $scriptRoot -ChildPath "lib"
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
  "/reference:System.Xml.dll",
  "/reference:System.IO.Compression.dll",
  "/reference:System.IO.Compression.FileSystem.dll",
  "/reference:System.Windows.Forms.dll",
  "/reference:System.Drawing.dll"
)

if (Test-Path -LiteralPath $libDir) {
  $commonArgs += @(
    "/reference:$libDir/NPOI.dll",
    "/reference:$libDir/NPOI.OOXML.dll",
    "/reference:$libDir/NPOI.OpenXml4Net.dll",
    "/reference:$libDir/NPOI.OpenXmlFormats.dll",
    "/reference:$libDir/ICSharpCode.SharpZipLib.dll",
    "/reference:$libDir/BouncyCastle.Crypto.dll"
  )
}

$outX64 = Join-Path -Path $distDir -ChildPath "CertPhotoSorter_x64.exe"
$argsX64 = $commonArgs + @("/platform:x64", "/out:$outX64") + $sources
& $csc @argsX64

$outX86 = Join-Path -Path $distDir -ChildPath "CertPhotoSorter_x86.exe"
$argsX86 = $commonArgs + @("/platform:x86", "/out:$outX86") + $sources
& $csc @argsX86

if (Test-Path -LiteralPath $libDir) {
  Copy-Item -LiteralPath "$libDir/NPOI.dll" -Destination $distDir -Force
  Copy-Item -LiteralPath "$libDir/NPOI.OOXML.dll" -Destination $distDir -Force
  Copy-Item -LiteralPath "$libDir/NPOI.OpenXml4Net.dll" -Destination $distDir -Force
  Copy-Item -LiteralPath "$libDir/NPOI.OpenXmlFormats.dll" -Destination $distDir -Force
  Copy-Item -LiteralPath "$libDir/ICSharpCode.SharpZipLib.dll" -Destination $distDir -Force
  Copy-Item -LiteralPath "$libDir/BouncyCastle.Crypto.dll" -Destination $distDir -Force

  if (Test-Path -LiteralPath "$libDir/THIRD_PARTY_NOTICES.txt") {
    Copy-Item -LiteralPath "$libDir/THIRD_PARTY_NOTICES.txt" -Destination $distDir -Force
  }
  if (Test-Path -LiteralPath "$libDir/LICENSE.NPOI") {
    Copy-Item -LiteralPath "$libDir/LICENSE.NPOI" -Destination $distDir -Force
  }
}

Write-Host "Build done:"
Write-Host " - $outX64"
Write-Host " - $outX86"
