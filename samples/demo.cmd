@echo off
rem Demo batch script for Script Debugger
setlocal enabledelayedexpansion
set NAME=World
set COUNT=0
echo Hello %NAME%

:loop
set /a COUNT+=1
echo Iteration !COUNT!
if !COUNT! lss 3 goto loop

if exist "%~dp0" (
    set FOUND=yes
    echo Folder exists: %~dp0
) else (
    set FOUND=no
)

call :greet Alice
call :greet Bob
set "RESULT=done with %COUNT% iterations"
echo %RESULT%
exit /b 7

:greet
set LAST_GREETED=%1
echo Greetings, %1!
goto :eof
