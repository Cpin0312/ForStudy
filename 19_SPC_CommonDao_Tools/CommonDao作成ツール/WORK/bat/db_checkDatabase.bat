@echo off
cd /d %~dp0

Set CUR_PATH=%1
set PGHOST=%2
set PGPORT=%3
set PGDATABASE=%4
set PGUSER=%5
set PGPASSWORD=%6

REM �����R�[�h�ݒ�
SET PGCLIENTENCODING=utf-8
chcp 65001

REM ���s���e(���낢��m�F)
psql -c "Select 'Connect Success Time : ' || current_timestamp" -t

exit %ERRORLEVEL%
