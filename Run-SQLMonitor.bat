@echo off
:: Run SQLMonitor directly without compiling - works on any Windows machine
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%~dp0SQLMonitor.ps1"
