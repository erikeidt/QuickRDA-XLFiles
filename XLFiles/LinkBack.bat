!@echo off
if NOT "%QRDebug%"=="" echo on
rem LinkBack.bat
rem wscript %QuickRDA%\LinkBack.vbs //x %1%
wscript "%~dp0"\LinkBack.vbs %1%
if NOT "%QRDebug%"=="" pause
rem LinkBack.bat
