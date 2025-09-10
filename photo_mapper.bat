@echo off
setlocal ENABLEDELAYEDEXPANSION

REM --- Usage check ---
if "%~1"=="" (
  echo Drag a folder onto this .bat to create a KML from geotagged images.
  echo.
  echo Example: Drop "C:\Photos\Trip\" here.
  pause
  exit /b 1
)

set "TARGET=%~1"
set "SCRIPT=%~dp0photo_mapper.py"

if not exist "%SCRIPT%" (
  echo Can't find Python script at: %SCRIPT%
  echo Make sure photo_mapper.py is in the same folder as this .bat.
  pause
  exit /b 1
)

REM --- Find Python (python > py -3 > py) ---
set "PYTHON="
where python >nul 2>&1 && set "PYTHON=python"
if not defined PYTHON (
  where py >nul 2>&1 && set "PYTHON=py -3"
)

if not defined PYTHON (
  echo Could not find Python. Install Python 3 or ensure it's on PATH.
  pause
  exit /b 1
)

REM --- Run ---
%PYTHON% "%SCRIPT%" "%TARGET%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo Done.
) else (
  echo Finished with errors. Exit code: %RC%
)
pause
exit /b %RC%
