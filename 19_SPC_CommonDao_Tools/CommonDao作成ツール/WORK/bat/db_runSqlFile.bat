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

REM SQL�t�@�C�������s
for /r %7 %%A in (*.*) do (
    REM �������s������A���s�������Ȃ�
    if %ERRORLEVEL% NEQ 0 goto :continue
    REM ���s���e
    psql --set ON_ERROR_STOP=on -f %%A
    
)

exit %ERRORLEVEL%
