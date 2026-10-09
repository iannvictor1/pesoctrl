param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$BackupDir = (Join-Path $ProjectDir "backups\postgres"),
    [int]$DiasReter = 30
)

$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $ProjectDir
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFile = Join-Path $BackupDir "pesoctrl_$timestamp.dump"

docker exec pesoctrl_postgres pg_dump -U pesoctrl -d pesoctrl -Fc -f "/tmp/pesoctrl.dump"
if ($LASTEXITCODE -ne 0) { throw 'pg_dump falhou; backups anteriores foram preservados.' }
docker cp "pesoctrl_postgres:/tmp/pesoctrl.dump" $backupFile
if ($LASTEXITCODE -ne 0) { throw 'Falha ao copiar backup do container.' }
if (!(Test-Path -LiteralPath $backupFile) -or (Get-Item -LiteralPath $backupFile).Length -eq 0) {
    throw 'Backup vazio ou ausente; backups anteriores foram preservados.'
}
docker exec pesoctrl_postgres rm -f "/tmp/pesoctrl.dump"

$limite = (Get-Date).AddDays(-1 * $DiasReter)
Get-ChildItem -LiteralPath $BackupDir -Filter "*.dump" |
    Where-Object { $_.LastWriteTime -lt $limite } |
    Remove-Item -Force

Write-Host "Backup criado: $backupFile" -ForegroundColor Green
