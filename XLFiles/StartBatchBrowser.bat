@echo off
if NOT "%QRDebug%"=="" echo on
rem StartBatchBrowser.bat
Call "%~dp0"StartBatchJob.bat %1
Call "%~dp0"StartBrowser.bat %1
if NOT "%QRDebug%"=="" pause
rem StartBatchBrowser.bat
