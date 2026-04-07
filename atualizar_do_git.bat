@echo off
setlocal enabledelayedexpansion

REM ===== CONFIG =====
set REPO_DIR=C:\brev\inativos\clientes_sem_mov_app
set BRANCH=main
set LOG_FILE=%REPO_DIR%\auto_update.log
REM ==================

call :log "=================================================="
call :log "Iniciando atualização automática..."

cd /d "%REPO_DIR%" || (
  call :log "ERRO: pasta do repositório não encontrada."
  exit /b 1
)

git rev-parse --is-inside-work-tree >nul 2>&1 || (
  call :log "ERRO: não é um repositório git válido."
  exit /b 1
)

REM Detecta alterações locais
git status --porcelain > "%TEMP%\git_changes.tmp"
for /f %%A in ('type "%TEMP%\git_changes.tmp" ^| find /c /v ""') do set CHANGED=%%A
del "%TEMP%\git_changes.tmp" >nul 2>&1

if NOT "!CHANGED!"=="0" (
  call :log "AVISO: alterações locais detectadas. Pull ignorado para evitar conflito."
  call :log "DICA: rode 'git status' para ver os arquivos."
  exit /b 0
)

call :log "Executando git fetch origin..."
git fetch origin >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  call :log "ERRO no git fetch."
  exit /b 1
)

for /f %%i in ('git rev-parse HEAD') do set LOCAL_HASH=%%i
for /f %%i in ('git rev-parse origin/%BRANCH%') do set REMOTE_HASH=%%i

if "!LOCAL_HASH!"=="!REMOTE_HASH!" (
  call :log "Sem atualizações. Já está no último commit (!LOCAL_HASH!)."
  exit /b 0
)

call :log "Atualização encontrada. Executando git pull --ff-only..."
git pull --ff-only origin %BRANCH% >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  call :log "ERRO no git pull."
  exit /b 1
)

call :log "Atualizado com sucesso para !REMOTE_HASH!."
exit /b 0

:log
set MSG=%~1
echo [%date% %time%] %MSG%
>> "%LOG_FILE%" echo [%date% %time%] %MSG%
goto :eof