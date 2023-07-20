REM # 本バッチファイルディレクトリを基準とする。
cd /d %~d0%~p0

REM # APIモジュール生成
java -jar %1 %2 true

REM # バッチモジュール生成
java -jar %1 %3 true 

exit %ERRORLEVEL%
