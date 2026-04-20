@echo off
chcp 65001 > nul
echo.
echo ===============================================
echo   БГУИР Нормоконтроль — СТП 01-2024
echo ===============================================
echo.
echo Запуск локального сервера...
echo Откроется браузер с приложением.
echo.
echo Для остановки нажмите Ctrl+C
echo.

:: Запускаем сервер на порту 8080
start /min "" python -c "import http.server, socketserver, os; os.chdir(r'%~dp0app'); s = socketserver.TCPServer(('', 8080), http.server.SimpleHTTPRequestHandler); print('Сервер: http://localhost:8080'); s.serve_forever()"

timeout /t 2 /nobreak > nul
start "" "http://localhost:8080"

echo Сервер запущен: http://localhost:8080
echo.
pause
