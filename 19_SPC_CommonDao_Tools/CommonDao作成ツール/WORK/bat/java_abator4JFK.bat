REM # �{�o�b�`�t�@�C���f�B���N�g������Ƃ���B
cd /d %~d0%~p0

REM # API���W���[������
java -jar %1 %2 true

REM # �o�b�`���W���[������
java -jar %1 %3 true 

exit %ERRORLEVEL%
