@echo off
REM 一键启动 bank2excel-h5 私有转换服务
REM 启动后: 手机与电脑同一 WiFi → 手机浏览器打开 http://<电脑IP>:8766
chcp 65001 >nul
echo ============================================
echo  银行对账单 PDF - Excel  私有转换服务
echo ============================================
echo.
set VENV=C:\Users\Administrator\Documents\银行对账单转化pdf\.venv
cd /d %~dp0
"%VENV%\Scripts\python.exe" server.py --host 0.0.0.0 --port 8766
pause
