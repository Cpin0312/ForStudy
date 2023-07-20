@ECHO OFF

REM |--------------------------------------------------------------------------------------------------------------------
REM | Purpose:  Generic Excel Addin Install
REM |--------------------------------------------------------------------------------------------------------------------


REM
REM     /E   = Copies directories and sub-directories, including empty ones. Same as /S /E. May be used to modify /T. 
REM     /D:m-d-y = Copies files changed on or after the specified date. 
REM        If no date is given, copies only those files whose source time is newer than the destination time. 
REM     /K   = Copies attributes. Normal Xcopy will reset read-only attributes. 
REM     /Q   = Does not display file names while copying. 
REM     /R   = Overwrites read-only files. 
REM     /Y   = Suppresses prompting to confirm you want to overwrite an existing destination file. 
REM
Set mydate=%date:~0,4%%date:~5,2%%date:~8,2%
REM Copy the install directory and sub-directories
REM echo f | XCOPY ".\createSql_Spc.xlam" "%AppData%\Microsoft\AddIns\createSql_Spc.xlam" /E /K /Q /R /Y /D
REM echo f | XCOPY "%AppData%\Microsoft\Excel\XLSTART\createSql_Spc.xlam" ".\createSql_Spc_"`date '+%Y%m%d%H%M%S'`".xlam" /E /K /Q /R /Y /D
    echo f | XCOPY "%AppData%\Microsoft\Excel\XLSTART\voyagerSPC.xla" ".\voyagerSPC.xla" /E /K /Q /R /Y /D
REM echo f | XCOPY ".\createSql_Spc.xlam" "%AppData%\Roaming\Microsoft\Excel\XLSTART\createSql_Spc.xlam" /E /K /Q /R /Y /D