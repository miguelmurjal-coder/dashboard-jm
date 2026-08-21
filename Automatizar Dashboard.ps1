$ErrorActionPreference = 'Stop'

$repo = $PSScriptRoot
$stateDir = Join-Path $repo '.automation'
$logFile = Join-Path $stateDir 'atualizacoes.log'
$lockFile = Join-Path $stateDir 'atualizacao.lock'
$gitSafeDirectory = $repo.Replace('\', '/')
$gitCommand = Get-Command git -ErrorAction SilentlyContinue
$gitExecutable = if ($gitCommand) {
    $gitCommand.Source
} else {
    Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\native\git\cmd\git.exe'
}

New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

if (-not (Test-Path -LiteralPath $gitExecutable)) {
    throw "Git não foi encontrado em $gitExecutable."
}

function Write-UpdateLog {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -LiteralPath $logFile -Value "[$timestamp] $Message" -Encoding UTF8
}

$lockStream = $null

try {
    try {
        $lockStream = [System.IO.File]::Open(
            $lockFile,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
    }
    catch [System.IO.IOException] {
        Write-UpdateLog 'Ignorada: já existe outra atualização em curso.'
        exit 0
    }

    Write-UpdateLog 'Início da atualização.'

    $env:PKE_NO_PAUSE = '1'
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $builderOutput = & (Join-Path $repo 'Atualizar Dashboard.cmd') 2>&1
    $builderExitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousErrorActionPreference
    foreach ($line in $builderOutput) {
        Write-UpdateLog "Builder: $line"
    }
    if ($builderExitCode -ne 0) {
        throw "O gerador terminou com o código $builderExitCode."
    }

    $gitBase = @('-c', "safe.directory=$gitSafeDirectory")
    & $gitExecutable @gitBase diff --quiet -- index.html
    $diffExitCode = $LASTEXITCODE

    if ($diffExitCode -eq 0) {
        Write-UpdateLog 'Sem alterações; não foi criado commit.'
        exit 0
    }
    if ($diffExitCode -ne 1) {
        throw "Não foi possível verificar as alterações do Git (código $diffExitCode)."
    }

    & $gitExecutable @gitBase add -- index.html
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao preparar index.html para commit.' }

    $commitMessage = 'Atualizar dashboard ' + (Get-Date -Format 'yyyy-MM-dd HH:mm')
    & $gitExecutable @gitBase commit -m $commitMessage
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao criar o commit automático.' }

    & $gitExecutable @gitBase push origin main
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao enviar a atualização para o GitHub.' }

    Write-UpdateLog 'Atualização enviada; publicação no Neocities iniciada.'
}
catch {
    Write-UpdateLog "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($lockStream) {
        $lockStream.Dispose()
    }
}
