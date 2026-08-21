param(
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'

$repo = $PSScriptRoot
$sourcePath = [System.IO.Path]::GetFullPath(
    (Join-Path $repo '..\..\..\INDICADOR MASTER 2026.xlsx')
)
$updateScript = Join-Path $repo 'Automatizar Dashboard.ps1'
$stateDir = Join-Path $repo '.automation'
$logFile = Join-Path $stateDir 'vigilancia.log'

New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

function Write-WatcherLog {
    param([string]$Message)
    if ((Test-Path -LiteralPath $logFile) -and (Get-Item -LiteralPath $logFile).Length -ge 1MB) {
        $archiveLog = Join-Path $stateDir 'vigilancia.anterior.log'
        Move-Item -LiteralPath $logFile -Destination $archiveLog -Force
    }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message$([Environment]::NewLine)"
    [System.IO.File]::AppendAllText($logFile, $line, [System.Text.UTF8Encoding]::new($false))
}

if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
    throw "O livro Excel não foi encontrado: $sourcePath"
}
if (-not (Test-Path -LiteralPath $updateScript -PathType Leaf)) {
    throw "A rotina de atualização não foi encontrada: $updateScript"
}

if ($ValidateOnly) {
    Write-Output "Excel: $sourcePath"
    Write-Output "Atualizador: $updateScript"
    Write-Output 'Validação concluída.'
    exit 0
}

$lastWriteTime = (Get-Item -LiteralPath $sourcePath).LastWriteTimeUtc
Write-WatcherLog 'Vigilância iniciada.'

while ($true) {
    Start-Sleep -Seconds 5

    try {
        $currentWriteTime = (Get-Item -LiteralPath $sourcePath).LastWriteTimeUtc
        if ($currentWriteTime -eq $lastWriteTime) {
            continue
        }

        $lastWriteTime = $currentWriteTime
        Write-WatcherLog 'Alteração detetada; a aguardar que a gravação termine.'
        Start-Sleep -Seconds 8

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $updateScript
        if ($LASTEXITCODE -eq 0) {
            Write-WatcherLog 'Atualização imediata concluída.'
        }
        else {
            Write-WatcherLog "A atualização terminou com o código $LASTEXITCODE; a rotina de 15 minutos voltará a tentar."
        }

        $lastWriteTime = (Get-Item -LiteralPath $sourcePath).LastWriteTimeUtc
    }
    catch {
        Write-WatcherLog "ERRO: $($_.Exception.Message)"
        Start-Sleep -Seconds 15
    }
}
