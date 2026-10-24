# Telefonlarin bu bilgisayardaki fiyat sayfasina ulasabilmesi icin
# Windows Guvenlik Duvari'nda 3000 numarali porta izin verir.
# baslat.bat tarafindan ilk calistirmada otomatik cagrilir.

$ruleName = 'DilekOfisKatalog'

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        Start-Process powershell -Verb RunAs -Wait -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } catch {
        Write-Host '  Guvenlik duvari izni verilmedi. Telefonlar sayfayi acamayabilir.'
    }
    exit
}

Remove-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue
New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort 3000 -Profile Any | Out-Null

# Windows'un Node.js icin daha once olusturdugu "Engelle" kurallari izin kuralini ezer; onlari kaldir.
$nodePaths = @("$PSScriptRoot\node\node.exe", "$env:ProgramFiles\nodejs\node.exe", "$env:LOCALAPPDATA\Programs\nodejs\node.exe")
$cmd = Get-Command node -ErrorAction SilentlyContinue
if ($cmd) { $nodePaths += $cmd.Source }
Get-NetFirewallApplicationFilter |
    Where-Object { $nodePaths -contains $_.Program } |
    Get-NetFirewallRule |
    Where-Object { $_.Action -eq 'Block' } |
    Remove-NetFirewallRule

Write-Host '  Guvenlik duvari izni verildi.'
Start-Sleep -Seconds 2
