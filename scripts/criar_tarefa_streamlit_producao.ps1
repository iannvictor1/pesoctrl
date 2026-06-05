param(
    [string]$TaskName = "PesoCtrlStreamlit",
    [string]$ProjectDir = "C:\Users\Supervidor Externo\OneDrive\Desktop\pesoctrl",
    [int]$Port = 8501
)

$ErrorActionPreference = "Stop"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Abra o PowerShell como Administrador para criar a tarefa agendada."
}

$scriptPath = Join-Path $ProjectDir "scripts\iniciar_streamlit_producao.ps1"

if (!(Test-Path -LiteralPath $scriptPath)) {
    throw "Script nao encontrado: $scriptPath"
}

$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ProjectDir `"$ProjectDir`" -HostAddress 0.0.0.0 -Port $Port"

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument $arguments `
    -WorkingDirectory $ProjectDir

$trigger = New-ScheduledTaskTrigger -AtStartup

$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DisallowStartIfOnBatteries:$false `
    -ExecutionTimeLimit (New-TimeSpan -Days 0) `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName

Write-Host ""
Write-Host "Tarefa '$TaskName' criada e iniciada." -ForegroundColor Green
Write-Host "Acesse em: http://127.0.0.1:$Port"
