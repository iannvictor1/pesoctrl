param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [int]$Port = 5174
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath (Join-Path $ProjectDir "frontend")

if (!(Test-Path -LiteralPath ".\node_modules")) {
    npm install
}

npm run build
npm run preview -- --host 0.0.0.0 --port $Port
