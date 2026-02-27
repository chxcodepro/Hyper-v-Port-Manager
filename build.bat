@echo off
setlocal

echo 正在安装依赖...
pip install pyinstaller
if errorlevel 1 goto :error

echo 正在构建便携版 exe...
pyinstaller build.spec --clean
if errorlevel 1 goto :error

if exist "release\PortManager-Portable" rmdir /s /q "release\PortManager-Portable"
mkdir "release\PortManager-Portable"
copy /y "dist\PortManager.exe" "release\PortManager-Portable\PortManager.exe" >nul
copy /y "config.json" "release\PortManager-Portable\config.json" >nul
copy /y "README.md" "release\PortManager-Portable\README.md" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'release\PortManager-Portable\*' -DestinationPath 'release\PortManager-Portable.zip' -Force"
if errorlevel 1 goto :error

echo 完成！便携版已生成: release\PortManager-Portable.zip
pause
exit /b 0

:error
echo 构建失败，请检查上面的日志。
pause
exit /b 1
