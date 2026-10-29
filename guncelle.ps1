# Program her acilista GitHub'daki en yeni surumu indirir (server.js ve public klasoru).
# Urunler ve ayarlar (data klasoru) ile baslat.bat'a dokunulmaz. Internet yoksa eski surumle devam edilir.
$ErrorActionPreference = 'Stop'
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $url = 'https://codeload.github.com/serhanicoz389-a11y/g-lendirme/zip/refs/heads/claude/kirtasiye-qr-app-s4pkw0'
    $tmp = Join-Path $env:TEMP 'dilekofis-guncelleme'
    if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
    New-Item $tmp -ItemType Directory | Out-Null
    $zip = Join-Path $tmp 'program.zip'
    Invoke-WebRequest $url -OutFile $zip -UseBasicParsing -TimeoutSec 30
    Expand-Archive $zip $tmp -Force
    $src = (Get-ChildItem $tmp -Directory | Select-Object -First 1).FullName
    if (-not (Test-Path (Join-Path $src 'server.js'))) { throw 'eksik' }
    foreach ($f in 'server.js', 'package.json', 'README.md', 'guvenlik-duvari.ps1', 'guncelle.ps1') {
        Copy-Item (Join-Path $src $f) (Join-Path $PSScriptRoot $f) -Force
    }
    Copy-Item (Join-Path $src 'public\*') (Join-Path $PSScriptRoot 'public') -Recurse -Force
    Remove-Item $tmp -Recurse -Force
    Write-Host '  Program guncel.'
} catch {
    Write-Host '  Guncelleme alinamadi, mevcut surumle devam ediliyor.'
}
