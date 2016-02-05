@echo off

rem
rem 設計書格納フォルダにコピーする
rem

set TARGET=%1
set DEST_DIR="\\192.168.10.250\Common\APチーム共有暫定\90_その他\99_cijnext\個人フォルダ\江元"

rem このバッチが存在するフォルダをカレントに
pushd %0\..
cls

xcopy /V /Y /D *.bas %DEST_DIR%
xcopy /V /Y /D *.cls %DEST_DIR%


rem pause
exit

exit /b
