New-NetFirewallRule `
    -DisplayName "PesoCtrl ADB Outbound" `
    -Direction Outbound `
    -Program "C:\platform-tools\adb.exe" `
    -Action Allow `
    -ErrorAction SilentlyContinue

New-NetFirewallRule `
    -DisplayName "PesoCtrl ADB Inbound" `
    -Direction Inbound `
    -Program "C:\platform-tools\adb.exe" `
    -Action Allow `
    -ErrorAction SilentlyContinue

New-NetFirewallRule `
    -DisplayName "PesoCtrl ADB TCP 5555" `
    -Direction Outbound `
    -Protocol TCP `
    -RemotePort 5555 `
    -Action Allow `
    -ErrorAction SilentlyContinue

Get-NetFirewallRule -DisplayName "PesoCtrl ADB*" |
    Select-Object DisplayName, Direction, Action, Enabled |
    Format-Table -AutoSize

Read-Host "Regras criadas. Pressione Enter para fechar"
