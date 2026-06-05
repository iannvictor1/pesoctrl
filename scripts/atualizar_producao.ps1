param(
    [string]$ProjectDir = (Split-Path -Parent $PSScriptRoot),
    [string]$Service = "",
    [string]$Task = ""
)

$ErrorActionPreference = "Stop"

function Write-Step($Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Invoke-GitPull() {
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        $output = & git pull 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($output) {
        $output | ForEach-Object { Write-Host $_ }
    }

    if ($exitCode -ne 0) {
        throw "git pull falhou com codigo $exitCode."
    }
}

function Protect-LocalConfig() {
    if (!(Test-Path -LiteralPath "config.py")) {
        return
    }

    Write-Step "Protegendo config.py local"
    & git update-index --skip-worktree -- config.py
    if ($LASTEXITCODE -ne 0) {
        throw "Nao foi possivel proteger config.py antes do git pull."
    }
}

function Assert-SupportedPython($PythonExe) {
    $versionText = & $PythonExe -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
    $parts = $versionText.Split(".")
    $major = [int]$parts[0]
    $minor = [int]$parts[1]

    if ($major -ne 3 -or $minor -lt 9 -or $minor -gt 12) {
        throw "Python $versionText nao e suportado. Use Python 3.9 ate 3.12."
    }
}

function Get-RequirementsFile() {
    $candidates = @("requirements.txt", "requeriments.txt")

    foreach ($candidate in $candidates) {
        $path = Join-Path $ProjectDir $candidate
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Length -gt 0)) {
            return $path
        }
    }

    return $null
}

function Restart-OptionalService($Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }

    Write-Step "Reiniciando servico: $Name"
    Restart-Service -Name $Name -Force
}

function Restart-OptionalScheduledTask($Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }

    Write-Step "Reiniciando tarefa agendada: $Name"
    $taskName = if ($Name.StartsWith("\")) { $Name } else { "\$Name" }

    $queryOutput = schtasks.exe /Query /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($queryOutput | Out-String).Trim()
        if ($message -match "Acesso negado|Access is denied") {
            throw "Acesso negado ao consultar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
        }

        throw "Tarefa agendada '$taskName' nao encontrada. Detalhe: $message"
    }

    $endOutput = schtasks.exe /End /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($endOutput | Out-String).Trim()
        if ($message -notmatch "nao esta em execucao|not currently running") {
            if ($message -match "Acesso negado|Access is denied") {
                throw "Acesso negado ao parar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
            }

            throw "Nao foi possivel parar a tarefa agendada '$taskName'. Detalhe: $message"
        }
    }

    Start-Sleep -Seconds 2

    $runOutput = schtasks.exe /Run /TN $taskName 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($runOutput | Out-String).Trim()
        if ($message -match "Acesso negado|Access is denied") {
            throw "Acesso negado ao iniciar a tarefa '$taskName'. Abra o PowerShell como Administrador e rode o update novamente."
        }

        throw "Nao foi possivel iniciar a tarefa agendada '$taskName'. Detalhe: $message"
    }
}

Write-Step "Entrando no projeto"
Set-Location -LiteralPath $ProjectDir

Protect-LocalConfig

Write-Step "Atualizando codigo pelo Git"
Invoke-GitPull

Write-Step "Atualizando dependencias Python"
if (!(Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    python -m venv .venv
}

Assert-SupportedPython ".\.venv\Scripts\python.exe"
.\.venv\Scripts\python.exe -m pip install --upgrade pip

$requirementsFile = Get-RequirementsFile
if ($requirementsFile) {
    .\.venv\Scripts\python.exe -m pip install -r $requirementsFile
}
else {
    Write-Host "Nenhum arquivo de dependencias com conteudo encontrado."
}

Restart-OptionalService -Name $Service
Restart-OptionalScheduledTask -Name $Task

Write-Host ""
Write-Host "Atualizacao concluida." -ForegroundColor Green
