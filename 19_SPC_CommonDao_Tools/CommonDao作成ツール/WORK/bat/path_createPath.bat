@echo off
cd /d %~dp0

echo ファイルパスの削除と作成中。。。

Set CUR_PATH=%1

Set OUTPUT_PATH=%CUR_PATH%OUTPUT\

rmdir /s /q %OUTPUT_PATH%
rmdir /s /q %CUR_PATH%WORK\output\
rmdir /s /q %CUR_PATH%WORK\bat\java\

mkdir %CUR_PATH%WORK\bat\java\
mkdir %CUR_PATH%WORK\output\
mkdir %CUR_PATH%WORK\output\sql\
mkdir %CUR_PATH%WORK\output\txt\
mkdir %CUR_PATH%WORK\output\xml\

mkdir %CUR_PATH%INPUT\DB_INDEX_DLL\
mkdir %CUR_PATH%INPUT\DB_TABLE_DLL\

mkdir %OUTPUT_PATH%
mkdir %OUTPUT_PATH%xml\
mkdir %OUTPUT_PATH%xml\conf\
mkdir %OUTPUT_PATH%xml\conf\jar\
mkdir %OUTPUT_PATH%xml\conf\jar\spring\

mkdir %OUTPUT_PATH%xml\src\
mkdir %OUTPUT_PATH%xml\src\core\
mkdir %OUTPUT_PATH%xml\src\core\java\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\jfk\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\jfk\batch\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\jfk\batch\common\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\jfk\batch\common\db\
mkdir %OUTPUT_PATH%xml\src\core\java\jp\hitachisoft\jfk\batch\common\db\sqlmap\

exit %ERRORLEVEL%
