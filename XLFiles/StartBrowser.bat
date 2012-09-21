@echo off
if NOT "%QRDebug%"=="" echo on
rem StartBrowser.bat
:L1
start "viewer" %1.svg
shift
if NOT "%1"=="" goto L1
rem rem start "viewer" %1MM.txt
if NOT "%QRDebug%"=="" pause
rem StartBrowser.bat

