@echo off
echo 📦 Instalando dependencias de build...
pip install -r requirements-build.txt >nul
echo 🔨 Gerando executavel...
pyinstaller --onefile --noconsole --name "MegaSenaAnalyzer" mega_analyzer.py
echo ✅ Pronto! O arquivo esta em: dist\MegaSenaAnalyzer.exe
pause
