@echo off
cd /d %~dp0

Set CUR_PATH=%1
set PGHOST=%2
set PGPORT=%3
set PGDATABASE=%4
set PGUSER=%5
set PGPASSWORD=%6

REM 文字コード設定
SET PGCLIENTENCODING=utf-8
chcp 65001

REM SQLファイルを実行
for /r %7 %%A in (*.*) do (
    REM 処理失敗したら、続行処理しない
    if %ERRORLEVEL% NEQ 0 goto :continue
    REM 実行内容
    psql --set ON_ERROR_STOP=on -f %%A
    
)

exit %ERRORLEVEL%
