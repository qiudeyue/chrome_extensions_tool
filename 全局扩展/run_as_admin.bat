@echo off
powershell -Command "Start-Process python -Verb RunAs -ArgumentList 'chrome_extension_manager.py'" 