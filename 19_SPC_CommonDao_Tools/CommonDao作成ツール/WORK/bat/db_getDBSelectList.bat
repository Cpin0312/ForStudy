@echo off
cd /d %~dp0

Set CUR_PATH=%1
set PGHOST=%2
set PGPORT=%3
set PGDATABASE=%4
set PGUSER=%5
set PGPASSWORD=%6
set EXE_SQL=%7
set T_NAME_FILE_PATH=%8

REM �����R�[�h�ݒ�
SET PGCLIENTENCODING=utf-8
chcp 65001

REM ���s���e
psql -f %EXE_SQL% -t > %T_NAME_FILE_PATH%

exit %ERRORLEVEL%
