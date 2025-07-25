@echo off
cd /d %~dp0
chcp 65001 >nul
powershell.exe -executionpolicy bypass -file "Win11Debloat_CN.ps1"
pause