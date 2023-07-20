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

REM 実行内容(いろいろ確認)
psql -c "Select 'Connect Success Time : ' || current_timestamp" -t

exit %ERRORLEVEL%
