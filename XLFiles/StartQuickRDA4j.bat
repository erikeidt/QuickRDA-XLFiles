@echo off
if NOT "%QRDebug%"=="" echo on
rem StartQuickRDA4j.bat
cd "%~dp0"
%~d0
rem java -jar QuickRDA.jar %1 %2 %3
rem java -classpath .\JACOB\jacob.jar:. -jar QuickRDA.jar %1 %2 %3
java -Djava.library.path=.\JACOB -classpath .\QuickRDA.jar;.\JEB.jar;.\JACOB\jacob.jar. com.hp.QuickRDA.L5.ExcelTool.QuickRDAMain %1 %2 %3
if NOT "%QRDebug%"=="" pause
rem StartQuickRDA4j.bat
if %ERRORLEVEL% NEQ 0 pause
