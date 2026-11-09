@echo off
REM Starts the CSV cleaner in your browser. First start takes 1-2 minutes.
cd /d "%~dp0"
where uv >nul 2>nul
if errorlevel 1 (
  echo Installing uv (Python package manager, one-time)...
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  set "PATH=%USERPROFILE%\.local\bin;%PATH%"
)
uv run --python 3.12 --with-requirements requirements.txt streamlit run app.py
pause
