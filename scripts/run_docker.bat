@echo off
echo 🚀 Starting PowerPoint Agent Container...

:: Get parent directory for mounting
pushd %~dp0..
set HOST_DIR=%CD%
popd

echo 📂 Mounting: %HOST_DIR% -^> /app

:: Run interactive container
docker run --rm -it -v "%HOST_DIR%:/app" ppt-agent-tools:latest

