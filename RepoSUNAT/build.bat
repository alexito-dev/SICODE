@echo off
cd /d "%~dp0frontend"
call npm install
call npm run build || exit /b 1
cd /d "%~dp0"
"%~dp0..\venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --name RepoSUNAT --icon "%~dp0assets\icon.ico" --add-data "%~dp0static;static" --distpath dist --workpath build --specpath build app.py
