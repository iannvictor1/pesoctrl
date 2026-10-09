param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [switch]$Build,
    [int]$DockerTimeoutSeconds = 180
)
$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $ProjectDir
if (!(Get-Command docker -ErrorAction SilentlyContinue)) {
    throw "Instale o Docker Desktop com containers Linux e Docker Compose."
}
function Test-DockerReady {
    $ErrorActionPreference = "Continue"
    $kind = & docker info --format '{{.OSType}}' 2>&1
    if ($LASTEXITCODE -ne 0) { Write-Verbose ($kind | Out-String); return $false }
    return ($LASTEXITCODE -eq 0 -and $kind -eq 'linux')
}
if (!(Test-DockerReady)) {
    $desktop = Join-Path $env:ProgramFiles "Docker\Docker\Docker Desktop.exe"
    if (Test-Path -LiteralPath $desktop) {
        Start-Process -FilePath $desktop -WindowStyle Hidden
    }
    $deadline = (Get-Date).AddSeconds($DockerTimeoutSeconds)
    while (!(Test-DockerReady)) {
        if ((Get-Date) -ge $deadline) { throw "Docker Linux nao ficou pronto em $DockerTimeoutSeconds segundos." }
        Start-Sleep -Seconds 3
    }
}
# Evita que os bind mounts de arquivos sejam criados como diretorios.
foreach ($stateFile in @('recebimento_global.json', 'somador_entrega.json')) {
    if (!(Test-Path -LiteralPath $stateFile)) {
        [IO.File]::WriteAllText((Join-Path $ProjectDir $stateFile), '{}')
    }
    if (!(Test-Path -LiteralPath $stateFile -PathType Leaf)) { throw "$stateFile precisa ser um arquivo." }
}
$composeArgs = @('compose', 'up', '-d', '--wait', '--wait-timeout', '180')
if ($Build) { $composeArgs += '--build' } else { $composeArgs += '--no-build' }
& docker @composeArgs
if ($LASTEXITCODE -ne 0) { throw "Falha ao iniciar PesoCtrl. Consulte docker compose logs --tail 100." }
& docker compose ps
Write-Host "PesoCtrl iniciado. Acesse http://IP-DESTE-PC:5174 (ou FRONTEND_PORT do .env)." -ForegroundColor Green
