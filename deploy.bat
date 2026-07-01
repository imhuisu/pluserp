@echo off
echo [1/3] Git Pull...
cd /d "D:\Plusdoor Web\production"
git pull

echo [1.5/3] Installing packages...
"C:\Users\LHS\AppData\Local\Programs\Python\Python39-32\python.exe" -m pip install -r requirements.txt

echo [2/3] Stopping Flask...
taskkill /F /IM pythonw.exe /T >nul 2>&1
timeout /t 2 /nobreak >nul

echo [3/3] Starting Flask...
start "" /B "C:\Users\LHS\AppData\Local\Programs\Python\Python39-32\pythonw.exe" -m waitress --host=127.0.0.1 --port=5000 app:app

echo Done! Flask is running.
timeout /t 2 /nobreak >nul
