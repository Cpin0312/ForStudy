@echo off
cd /d %~dp0

Set CNT=0

REM ファイルを実行
for /r %1 %%A in (*.*) do (

	Set /a CNT+=1
)

exit %CNT%