@echo off
SetLocal

del /s /q %APPDATA%\pyinstaller\*
del /s /q .\build\*
::pip install pypiwin32
pyinstaller -F --distpath . parse.py

EndLocal
pause
