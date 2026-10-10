@echo off
cd /d "%~dp0"

netstat -ano | findstr /R /C:"127\.0\.0\.1:3001 .*LISTENING" >nul
if %errorlevel%==0 (
    start "" "http://127.0.0.1:3001/"
    exit /b 0
)

npm run dev -- --host 127.0.0.1 --port 3001 --strictPort --open /