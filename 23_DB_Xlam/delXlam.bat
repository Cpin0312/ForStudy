@ECHO OFF

REM
REM /P ファイルの削除前に確認メッセージを表示する 
REM /F 読み取り専用ファイルを削除する 
REM /S 指定したファイルを全てのサブディレクトリから削除し、削除したファイル名を表示する  
REM /Q 削除前に確認メッセージを表示しない 
REM

REM echo f | DEL "%AppData%\Microsoft\Excel\XLSTART\createSql_Spc.xlam" /F
    echo f | DEL "%AppData%\Microsoft\Excel\XLSTART\voyagerSPC.xla" /F 
