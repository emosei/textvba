@echo off

rem
rem �݌v���i�[�t�H���_�ɃR�s�[����
rem

set TARGET=%1
set DEST_DIR="\\192.168.10.250\Common\AP�`�[�����L�b��\90_���̑�\99_cijnext\�l�t�H���_\�]��"

rem ���̃o�b�`�����݂���t�H���_���J�����g��
pushd %0\..
cls

xcopy /V /Y /D *.bas %DEST_DIR%
xcopy /V /Y /D *.cls %DEST_DIR%


rem pause
exit

exit /b
