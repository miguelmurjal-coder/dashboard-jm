$ErrorActionPreference = 'Stop'

$repo = $PSScriptRoot
$stateDir = Join-Path $repo '.automation'
$logFile = Join-Path $stateDir 'atualizacoes.log'
$lockFile = Join-Path $stateDir 'atualizacao.lock'
$sourceStateFile = Join-Path $stateDir 'fontes.json'
$sourcePath = [System.IO.Path]::GetFullPath(
    (Join-Path $repo '..\..\..\INDICADOR MASTER 2026.xlsx')
)
$vacationUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQXk2yWBJ5SFJUPhJG7oBWhqs5tJylVsDWBl6GGndu2oWrwti6e6csHZpmxaJG9ywzStdmR0_4Q2URX/pub?gid=1810617663&single=true&output=csv'
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
    if ((Test-Path -LiteralPath $logFile) -and (Get-Item -LiteralPath $logFile).Length -ge 1MB) {
        $archiveLog = Join-Path $stateDir 'atualizacoes.anterior.log'
        Move-Item -LiteralPath $logFile -Destination $archiveLog -Force
    }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message$([Environment]::NewLine)"
    [System.IO.File]::AppendAllText($logFile, $line, [System.Text.UTF8Encoding]::new($false))
}

function Get-RemoteContentHash {
    param([string]$Uri)
    $response = Invoke-WebRequest -Uri $Uri -UseBasicParsing -TimeoutSec 30
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$response.Content)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha256.ComputeHash($bytes))).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
    }
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

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "O livro Excel não foi encontrado: $sourcePath"
    }

    $previousState = $null
    if (Test-Path -LiteralPath $sourceStateFile) {
        try {
            $previousState = Get-Content -LiteralPath $sourceStateFile -Raw -Encoding UTF8 | ConvertFrom-Json
        }
        catch {
            Write-UpdateLog 'Aviso: o estado anterior das fontes não pôde ser lido; será reconstruído.'
        }
    }

    $excelHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
    $vacationHash = if ($previousState) { [string]$previousState.vacationHash } else { '' }
    try {
        $vacationHash = Get-RemoteContentHash -Uri $vacationUrl
    }
    catch {
        Write-UpdateLog 'Aviso: não foi possível verificar o plano de férias online; será usado o último estado conhecido.'
    }

    if ($previousState -and
        $excelHash -eq [string]$previousState.excelHash -and
        $vacationHash -eq [string]$previousState.vacationHash) {
        Write-UpdateLog 'Fontes sem alterações; reconstrução ignorada.'
        exit 0
    }

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

    $newState = [ordered]@{
        excelHash = $excelHash
        vacationHash = $vacationHash
        checkedAt = (Get-Date).ToString('o')
    }

    $gitBase = @('-c', "safe.directory=$gitSafeDirectory")
    & $gitExecutable @gitBase diff --quiet -- index.html
    $diffExitCode = $LASTEXITCODE

    if ($diffExitCode -eq 0) {
        $newState | ConvertTo-Json | Set-Content -LiteralPath $sourceStateFile -Encoding UTF8
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

    $newState | ConvertTo-Json | Set-Content -LiteralPath $sourceStateFile -Encoding UTF8
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
