param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$HostAddress = "0.0.0.0",
    [int]$Port = 8501
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $ProjectDir

if (!(Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}

$requirementsFile = $null
foreach ($candidate in @("requirements.txt", "requeriments.txt")) {
    $path = Join-Path $ProjectDir $candidate
    if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Length -gt 0)) {
        $requirementsFile = $path
        break
    }
}

if ($requirementsFile) {
    .\.venv\Scripts\python.exe -m pip install -r $requirementsFile
}

.\.venv\Scripts\python.exe -m streamlit run tela_controle.py --server.address $HostAddress --server.port $Port
