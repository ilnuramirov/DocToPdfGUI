@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0Build-EXE.ps1'"
pause