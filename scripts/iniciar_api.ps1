param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$HostAddress = "0.0.0.0",
    [int]$Port = 8001
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $ProjectDir

if (!(Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install -r requirements.txt

$tentativas = 30
for ($i = 1; $i -le $tentativas; $i++) {
    $portaAberta = Test-NetConnection -ComputerName "127.0.0.1" -Port 15432 -InformationLevel Quiet
    if ($portaAberta) {
        break
    }

    if ($i -eq $tentativas) {
        throw "PostgreSQL nao respondeu na porta 15432."
    }

    Start-Sleep -Seconds 2
}

.\.venv\Scripts\python.exe -m uvicorn backend.main:app --host $HostAddress --port $Port
