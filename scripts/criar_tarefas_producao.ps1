param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$TaskPrefix = 'PesoCtrl',
    [ValidateRange(0,23)][int]$BackupHour = 23,
    [ValidateRange(0,59)][int]$BackupMinute = 30
)
$ErrorActionPreference = 'Stop'
$logDir = Join-Path $ProjectDir '.runtime'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
Start-Transcript -Path (Join-Path $logDir 'instalacao_docker.log') -Force | Out-Null
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Abra o PowerShell como Administrador para registrar tarefas e liberar a rede local.'
}
$ProjectDir = (Resolve-Path -LiteralPath $ProjectDir).Path
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15) -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable
function Register-PesoTask($Name, $Script, $Trigger) {
    $scriptPath = Join-Path $ProjectDir "scripts\$Script"
    if (!(Test-Path -LiteralPath $scriptPath)) { throw "Script nao encontrado: $scriptPath" }
    $arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -ProjectDir `"$ProjectDir`""
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arguments -WorkingDirectory $ProjectDir
    Register-ScheduledTask -TaskName $Name -Action $action -Trigger $Trigger -Principal $principal -Settings $settings -Force | Out-Null
}
Push-Location -LiteralPath $ProjectDir
try {
    $json = & docker compose config --format json
    if ($LASTEXITCODE -ne 0) { throw 'Configuracao Docker invalida.' }
    $compose = ($json -join "`n") | ConvertFrom-Json
} finally { Pop-Location }
Register-PesoTask "$TaskPrefix-Docker" 'iniciar_docker_producao.ps1' (New-ScheduledTaskTrigger -AtLogOn -User $currentUser)
Register-PesoTask "$TaskPrefix-BackupPostgres" 'backup_postgres.ps1' (New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours($BackupHour).AddMinutes($BackupMinute)))
# Desativa as tarefas conhecidas do modo anterior, sem exclui-las.
foreach ($oldName in @("$TaskPrefix-Postgres", "$TaskPrefix-API", "$TaskPrefix-Frontend", "$TaskPrefix-Streamlit", "${TaskPrefix}Streamlit")) {
    $old = Get-ScheduledTask -TaskName $oldName -ErrorAction SilentlyContinue
    if ($old) {
        Stop-ScheduledTask -InputObject $old
        Disable-ScheduledTask -InputObject $old | Out-Null
    }
}
foreach ($service in @('frontend', 'streamlit')) {
    $port = [int]$compose.services.$service.ports[0].published
    $rule = "$TaskPrefix-Docker-$service"
    if (Get-NetFirewallRule -Name $rule -ErrorAction SilentlyContinue) { Remove-NetFirewallRule -Name $rule }
    New-NetFirewallRule -Name $rule -DisplayName $rule -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -Profile Any -RemoteAddress LocalSubnet | Out-Null
}
Start-ScheduledTask -TaskName "$TaskPrefix-Docker"
Write-Host 'Tarefa Docker registrada e disparada. A inicializacao ocorre no login deste usuario.' -ForegroundColor Green
Write-Host 'Backup diario mantido; requer usuario conectado e Docker em execucao.'
Stop-Transcript | Out-Null
