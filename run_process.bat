@echo off
setlocal enabledelayedexpansion

rem 取得檔名參數 (-f filename) 或互動輸入
if "%~1"=="-f" (
  set "FILENAME=%~2"
) else (
  if not "%~1"=="" (
    set "FILENAME=%~1"
  ) else (
    set /p FILENAME=請輸入要處理的檔案名（Enter 使用預設: 回收問卷_大學部應屆畢業生_ReWu.xlsx）: 
    if "%FILENAME%"=="" set "FILENAME=回收問卷_大學部應屆畢業生_ReWu.xlsx"
  )
)

rem 選擇要用的 python（可透過環境變數 PYTHON 指定）
if defined PYTHON (
  set "PYTHON_CMD=%PYTHON%"
) else (
  set "PYTHON_CMD=python"
)

rem 建立虛擬環境（如無 .venv）
if not exist ".venv" (
  echo 建立虛擬環境 .venv
  "%PYTHON_CMD%" -m venv .venv
)

rem 啟用虛擬環境
call .venv\Scripts\activate.bat

echo 升級 pip 並安裝相依
python -m pip install --upgrade pip
if exist requirements.txt (
  pip install -r requirements.txt
) else (
  pip install pandas numpy xlsxwriter openpyxl
)

rem 去除前後單/雙引號（互動輸入可能包含引號）
set "first=!FILENAME:~0,1!"
set "last=!FILENAME:~-1!"
if "!first!"=="'" set "FILENAME=!FILENAME:~1!"
if "!last!"=="'" set "FILENAME=!FILENAME:~0,-1!"
set "first=!FILENAME:~0,1!"
set "last=!FILENAME:~-1!"
if "!first!"=="\"" set "FILENAME=!FILENAME:~1!"
if "!last!"=="\"" set "FILENAME=!FILENAME:~0,-1!"

echo 執行 process.py 處理：%FILENAME%
python process.py -f "%FILENAME%"

for %%F in ("%FILENAME%") do set "NAME=%%~nF"
echo 輸出檔案：%NAME%_Result.xlsx



endlocal
echo 完成。