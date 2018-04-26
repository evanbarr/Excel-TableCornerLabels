@echo off
setlocal

set zDEST=%appdata%\Microsoft\AddIns
set zFILE=CornerLabels.xlam

if not exist %zFILE% goto NoFileToCopy

if exist "%zDEST%\%zFILE%" del "%zDEST%\%zFILE%"
copy %zFILE% "%zDEST%\%zFILE%"

pause
goto :EOF



:NoFileToCopy
echo Could not find file "%zFILE%".  Installation aborted.
