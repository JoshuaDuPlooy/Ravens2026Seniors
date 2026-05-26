@echo off
echo Starting Senior Open 2026 website...
echo.
echo Open your browser and go to: http://localhost:8000/groups.html
echo.
echo Press Ctrl+C to stop the server when done.
echo.
cd /d "%~dp0"
python -m http.server 8000
pause
