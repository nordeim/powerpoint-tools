@echo off
echo 🐳 Building PowerPoint Agent Tools Docker Image...

:: Get parent directory
pushd %~dp0..
docker build -t ppt-agent-tools:latest -f docker\Dockerfile .
popd

echo ✅ Build Complete. Image: ppt-agent-tools:latest
pause

