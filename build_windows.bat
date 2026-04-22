@echo off
setlocal

cd /d "%~dp0"

set "APP_NAME=DataHub"
set "ENTRY_FILE=app.py"
set "APP_CORE_DIR=app_core"
set "BUILD_EXE=dist\%APP_NAME%.exe"
set "PORTABLE_DIR=dist\%APP_NAME%"

if not exist "%ENTRY_FILE%" (
    if exist "modular_app\app.py" (
        cd /d "%~dp0modular_app"
    )
)

echo ==========================================
echo  GERANDO EXE PORTATIL DO APP
echo ==========================================

echo.
echo [1/5] Validando estrutura do projeto...
if not exist "%ENTRY_FILE%" (
    echo Arquivo de entrada nao encontrado: %cd%\%ENTRY_FILE%
    goto erro
)

if not exist "%APP_CORE_DIR%\__init__.py" (
    echo Pasta modular app_core nao encontrada em: %cd%\%APP_CORE_DIR%
    goto erro
)

echo.
echo [2/5] Instalando dependencias de build...
python -m pip install --upgrade pip
if errorlevel 1 goto erro

python -m pip install --upgrade pyinstaller psycopg2-binary
if errorlevel 1 goto erro

echo.
echo [3/5] Limpando artefatos anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist "%APP_NAME%.spec" del /f /q "%APP_NAME%.spec"
if exist "%APP_CORE_DIR%\__pycache__" rmdir /s /q "%APP_CORE_DIR%\__pycache__"

echo.
echo [4/5] Gerando executavel com PyInstaller...
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name %APP_NAME% ^
  --paths "%cd%" ^
  --hidden-import ui ^
  --add-data "danfe_emitente_logo.png;." ^
  --collect-submodules app_core ^
  "%ENTRY_FILE%"
if errorlevel 1 goto erro

echo.
echo [5/5] Montando pasta portatil...
if not exist "%PORTABLE_DIR%" mkdir "%PORTABLE_DIR%"

copy /Y "%BUILD_EXE%" "%PORTABLE_DIR%\%APP_NAME%.exe" >nul
if errorlevel 1 goto erro

if exist "config.json" copy /Y "config.json" "%PORTABLE_DIR%\config.json" >nul
if exist "databrev.key" copy /Y "databrev.key" "%PORTABLE_DIR%\databrev.key" >nul
if exist "log" xcopy /E /I /Y "log" "%PORTABLE_DIR%\log" >nul

echo.
echo ==========================================
echo  COMPILACAO CONCLUIDA
echo ==========================================
echo Executavel: %cd%\%PORTABLE_DIR%\%APP_NAME%.exe
echo Pasta portatil: %cd%\%PORTABLE_DIR%
echo.
pause
exit /b 0

:erro
echo.
echo ==========================================
echo  ERRO NA COMPILACAO
echo ==========================================
echo Verifique as mensagens acima para identificar a falha.
echo Diretorio atual: %cd%
echo.
pause
exit /b 1
