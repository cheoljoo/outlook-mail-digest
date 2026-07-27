@echo off
setlocal

where uv >nul 2>nul
if errorlevel 1 (
    echo [INFO] uv not found. Installing uv...
    powershell -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
    set "PATH=%USERPROFILE%\.local\bin;%PATH%"
)

where uv >nul 2>nul
if errorlevel 1 (
    echo [ERROR] uv installation failed. Install manually: https://docs.astral.sh/uv/
    exit /b 1
)

cd /d %~dp0
uv run sync_outlook_digest.py
