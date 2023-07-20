@echo off
cd /d %~dp0

Set CUR_PATH=%1
set PGHOST=%2
set PGPORT=%3
set PGDATABASE=%4
set PGUSER=%5
set PGPASSWORD=%6
set SQL_INPUT_PATH=%7
set SQL_OUTPUT_PATH=%8

set sqlDropTblFile=%SQL_INPUT_PATH%getDropTableQuery.sql
set sqlDropSeqFile=%SQL_INPUT_PATH%getDropSeqQuery.sql
set dropTblFile=%SQL_OUTPUT_PATH%
set dropSeqFile=%SQL_OUTPUT_PATH%

set dropTblFile=%dropTblFile%dropTbl.txt
set dropSeqFile=%dropSeqFile%dropSeq.txt



REM 文字コード設定
SET PGCLIENTENCODING=utf-8
chcp 65001

REM 開始
echo BEGIN; > %dropTblFile%

REM 実行内容
psql -f %sqlDropTblFile% -t >> %dropTblFile%

REM 終了
echo END; >> %dropTblFile%

REM TBL削除
psql --set ON_ERROR_STOP=on -f %dropTblFile%



REM 開始
echo BEGIN; > %dropSeqFile%

REM 実行内容
psql -f %sqlDropSeqFile% -t >> %dropSeqFile%

REM 終了
echo END; >> %dropSeqFile%

set PGUSER=%9
REM 引数は9までしか使えないため、シフトで10個目のパラメタwp使用する
shift
set PGPASSWORD=%9
REM TBL削除
psql --set ON_ERROR_STOP=on -f %dropSeqFile%

exit %ERRORLEVEL%
