@echo off
SetLocal

pyinstaller -F --distpath . parse.py

EndLocal
pause
