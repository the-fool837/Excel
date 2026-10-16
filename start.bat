@echo off
setlocal enabledelayedexpansion
set "HERE=%~dp0"
set "SCRIPT=%HERE%fill_excel.py"
set "PYW="

rem 1) 优先使用已自带 openpyxl 的 Python
set "CAND=%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\pythonw.exe"
if exist "%CAND%" set "PYW=%CAND%"

rem 2) 其次使用 PATH 里的 pythonw
if not defined PYW (
  for /f "delims=" %%i in ('where pythonw 2^>nul') do (
    if not defined PYW set "PYW=%%i"
  )
)

rem 3) 再用 py 启动器
if not defined PYW (
  for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
    if not defined PYW set "PYW=%%i"
  )
)

if not defined PYW (
  echo.
  echo   没有找到可用的 Python 3，请先安装 Python 并勾选 Add Python to PATH。
  echo.
  pause
  exit /b 1
)

start "" "%PYW%" "%SCRIPT%"
exit /b 0