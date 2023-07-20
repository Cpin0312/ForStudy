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



REM �����R�[�h�ݒ�
SET PGCLIENTENCODING=utf-8
chcp 65001

REM �J�n
echo BEGIN; > %dropTblFile%

REM ���s���e
psql -f %sqlDropTblFile% -t >> %dropTblFile%

REM �I��
echo END; >> %dropTblFile%

REM TBL�폜
psql --set ON_ERROR_STOP=on -f %dropTblFile%



REM �J�n
echo BEGIN; > %dropSeqFile%

REM ���s���e
psql -f %sqlDropSeqFile% -t >> %dropSeqFile%

REM �I��
echo END; >> %dropSeqFile%

set PGUSER=%9
REM ������9�܂ł����g���Ȃ����߁A�V�t�g��10�ڂ̃p�����^wp�g�p����
shift
set PGPASSWORD=%9
REM TBL�폜
psql --set ON_ERROR_STOP=on -f %dropSeqFile%

exit %ERRORLEVEL%
