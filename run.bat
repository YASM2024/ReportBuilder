@echo off
echo.
echo ========================================
echo   ReportGenerator V1.0
echo ========================================
echo.

setlocal EnableExtensions

rem exe と同じフォルダを作業ディレクトリにする（配布フォルダで実行）
cd /d "%~dp0"

set "PYTHON=.venv\Scripts\python.exe"
if not exist "%PYTHON%" (
    echo エラー: Python が見つかりません。
    echo   %PYTHON%
    echo 先にセットアップを実行してください。
    echo   python -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    set "EXIT_CODE=1"
    goto :abnormal
)

"%PYTHON%" main.py -c config\report.yaml
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
