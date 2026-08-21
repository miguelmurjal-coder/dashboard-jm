@echo off
setlocal
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..\..\..\INDICADOR MASTER 2026.xlsx") do set "SOURCE=%%~fI"

if not exist "%SOURCE%" (
  echo ERRO: O ficheiro INDICADOR MASTER 2026.xlsx nao foi encontrado.
  echo Caminho procurado: %SOURCE%
  goto end
)

set "CODEX_PYTHON=C:\Users\pke\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%CODEX_PYTHON%" (
  "%CODEX_PYTHON%" "%SCRIPT_DIR%dashboard_builder.py" "%SOURCE%"
  goto result
)

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%SCRIPT_DIR%dashboard_builder.py" "%SOURCE%"
  goto result
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "%SCRIPT_DIR%dashboard_builder.py" "%SOURCE%"
  goto result
)

echo ERRO: Python nao foi encontrado neste computador.
echo Abra esta tarefa no Codex para atualizar o pacote.
goto end

:result
if errorlevel 1 (
  echo.
  echo A atualizacao falhou. O dashboard anterior foi mantido.
) else (
  echo.
  echo Pode agora abrir index.html.
)

:end
echo.
if defined PKE_NO_PAUSE exit /b
pause
endlocal
