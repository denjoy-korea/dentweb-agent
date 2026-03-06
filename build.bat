@echo off
echo Building DentWeb Automation Agent...
pyinstaller --onefile --name dentweb-agent ^
  --add-data "images;images" ^
  --add-data "config.example.json;." ^
  main.py
echo Done. Output: dist\dentweb-agent.exe
pause
