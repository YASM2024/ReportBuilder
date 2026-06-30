@echo off
echo.
echo ========================================
echo   ReportGenerator V1.0
echo ========================================
echo.

setlocal EnableExtensions

rem exe と同じフォルダを作業ディレクトリにする（配布フォルダで実行）
cd /d "%~dp0"

set "APP_CONFIG=%~dp0config\app.yaml"
set "READ_CONFIG=%~dp0config\read_app_config.ps1"
set "PYTHON="
set "REPORT_CONFIG=config\report.yaml"

if exist "%APP_CONFIG%" (
    for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%READ_CONFIG%" -ConfigPath "%APP_CONFIG%" -Key python`) do set "PYTHON=%%P"
    for /f "usebackq delims=" %%R in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%READ_CONFIG%" -ConfigPath "%APP_CONFIG%" -Key report`) do set "REPORT_CONFIG=%%R"
)

if not defined PYTHON set "PYTHON=.venv\Scripts\python.exe"

rem 相対パスはアプリルート基準で解決
if not exist "%PYTHON%" (
    if exist "%~dp0%PYTHON%" set "PYTHON=%~dp0%PYTHON%"
)

if not exist "%PYTHON%" (
    echo エラー: Python が見つかりません。
    echo   %PYTHON%
    echo config\app.yaml の python を確認してください。
    echo   例: python: venv/Scripts/python.exe
    set "EXIT_CODE=1"
    goto :abnormal
)

"%PYTHON%" main.py -c %REPORT_CONFIG%
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" goto :normal
goto :abnormal

:normal
echo.
echo ========================================
echo   正常終了
echo ========================================
goto :finish

:abnormal
echo.
echo ========================================
echo   異常終了  (終了コード: %EXIT_CODE%)
echo ========================================
goto :finish

:finish
echo.
pause
exit /b %EXIT_CODE%
