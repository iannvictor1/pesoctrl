param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,

    [string]$ProjectDir = "C:\Users\Supervidor Externo\OneDrive\Desktop\pesoctrl",
    [string]$Service = "",
    [string]$Task = "",
    [string]$UserName = ""
)

$ErrorActionPreference = "Stop"

$credential = $null
if (![string]::IsNullOrWhiteSpace($UserName)) {
    $credential = Get-Credential -UserName $UserName -Message "Credenciais da maquina de producao"
}

$winRmPort = 5985
$connection = Test-NetConnection -ComputerName $ComputerName -Port $winRmPort -WarningAction SilentlyContinue
if (!$connection.TcpTestSucceeded) {
    throw "Nao foi possivel acessar $ComputerName na porta WinRM $winRmPort. Verifique se a maquina esta ligada, se o IP esta correto, se ela esta acessivel na rede e se o WinRM/firewall estao habilitados no computador de producao."
}

$script = {
    param($RemoteProjectDir, $RemoteService, $RemoteTask)

    function Invoke-RemoteGitPull() {
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
            throw "git pull remoto falhou com codigo $exitCode."
        }
    }

    function Protect-RemoteLocalConfig() {
        if (!(Test-Path -LiteralPath "config.py")) {
            return
        }

        Write-Host ""
        Write-Host "==> Protegendo config.py local" -ForegroundColor Cyan
        & git update-index --skip-worktree -- config.py
        if ($LASTEXITCODE -ne 0) {
            throw "Nao foi possivel proteger config.py antes do git pull remoto."
        }
    }

    Set-Location -LiteralPath $RemoteProjectDir

    Protect-RemoteLocalConfig

    Write-Host ""
    Write-Host "==> Sincronizando scripts pelo Git" -ForegroundColor Cyan
    Invoke-RemoteGitPull

    & .\scripts\atualizar_producao.ps1 `
        -ProjectDir $RemoteProjectDir `
        -Service $RemoteService `
        -Task $RemoteTask
}

$params = @{
    ComputerName = $ComputerName
    ScriptBlock = $script
    ArgumentList = @($ProjectDir, $Service, $Task)
}

if ($credential) {
    $params.Credential = $credential
}

Invoke-Command @params
