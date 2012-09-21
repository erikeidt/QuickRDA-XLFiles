@echo off
if NOT "%QRDebug%"=="" echo on
rem StartBatchJob.bat
rem having office in the path causes font (spacing & size estimation) problems for GraphViz, dunno why
set PATH=%PATH:Microsoft Office\Office=NOTHING%
del %1.svg 2> nul
if not exist "%systemdrive%\Program Files\Graphviz 2.29\bin\dot.exe" GOTO :T2
"%systemdrive%\Program Files\Graphviz 2.29\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T2
if not exist "%systemdrive%\Program Files (x86)\Graphviz 2.29\bin\dot.exe" GOTO :T3
"%systemdrive%\Program Files (x86)\Graphviz 2.29\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T3
if not exist "%systemdrive%\Program Files\Graphviz 2.28\bin\dot.exe" GOTO :T4
"%systemdrive%\Program Files\Graphviz 2.28\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T4
if not exist "%systemdrive%\Program Files (x86)\Graphviz 2.28\bin\dot.exe" GOTO :T5
"%systemdrive%\Program Files (x86)\Graphviz 2.28\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T5
if not exist "%systemdrive%\Program Files\Graphviz2.27\bin\dot.exe" GOTO :T6
"%systemdrive%\Program Files\Graphviz2.27\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T6
if not exist "%systemdrive%\Program Files (x86)\Graphviz2.27\bin\dot.exe" GOTO :T7
"%systemdrive%\Program Files (x86)\Graphviz2.27\bin\dot.exe" -Gcharset=latin1 -Tsvg -o%1.svg -Kdot %1.txt
GOTO :DONE
:T7
echo
echo *** QuickRDA cannot find GraphViz ***
echo   Please edit the following file with the path to GraphViz
echo     %QuickRDA%\StartBatchJob.bat
echo       (see also: http://www.graphviz.org)
echo 
echo   NOTE:GraphViz 2.26.3 is no longer supported by QuickRDA
echo	There is a new stable build of GraphViz, version 2.28.0 that should be used instead.
echo
pause
:DONE
if NOT "%QRDebug%"=="" pause
rem StartBatchJob.bat

