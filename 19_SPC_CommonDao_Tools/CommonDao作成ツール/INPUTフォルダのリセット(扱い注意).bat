@echo off
cd /d %~dp0

echo �t�@�C���p�X�̍폜�ƍ쐬���B�B�B

Set CUR_PATH=%1

Set INPUT_PATH=%CUR_PATH%INPUT

rmdir /s /q %INPUT_PATH%

mkdir %INPUT_PATH%\DB_TABLE_DLL\
mkdir %INPUT_PATH%\DB_INDEX_DLL\
mkdir %INPUT_PATH%\DB_SEQ_DLL\

exit %ERRORLEVEL%
