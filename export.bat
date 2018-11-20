@echo off

SET PATHSIAP=E:\SIAP_MejC
SET PATHEXP=E:\ExportSICOSS_MC

echo Abra el SICOSS y luego siga las instrucciones. Presione enter para continuar...
pause > nul
start %PATHSIAP%\siap.exe
echo.
echo Si ya abrio SICOSS, presione enter...
pause > nul
IF NOT EXIST sijp12\ MKDIR sijp12
IF NOT EXIST %PATHEXP%\ MKDIR %PATHEXP%
copy %PATHSIAP%\sijp12\SIJP12.mdb sijp12\
copy %PATHSIAP%\sijp12\SI220000.mdb sijp12\

echo.
echo Ya puede cerrar SIAP.
echo.
echo.
SET /P MONTH=Que mes desea exportar? (1-12):


export.exe %MONTH% T

SET TXTNAME=SICOSS_MC_%MONTH%.txt
move eSICOSS.txt %PATHEXP%\%TXTNAME%

explorer %PATHEXP%

