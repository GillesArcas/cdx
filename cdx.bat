@echo off
rem install: doskey cd=<full path>/cdx.bat $*

if "%~1"=="*" goto guiselect

cd /d "%~1"
py %~dpn0.py insert "%cd%"
goto :eof

:guiselect
for /F "delims=*" %%1 in ('py %~dpn0.py select d:/') do cd /d %%1
