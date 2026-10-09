param([string]$ProjectDir = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot 'iniciar_docker_producao.ps1') -ProjectDir $ProjectDir -Build
Write-Host 'Para iniciar ao entrar no Windows, execute como Administrador: .\scripts\criar_tarefas_producao.ps1'
