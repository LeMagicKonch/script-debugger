@echo off
rem Launches the Script Debugger GUI without a console window.
rem Optionally pass a script path: "Script Debugger.cmd" C:\path\to\script.cmd
start "" pythonw "%~dp0script_debugger.py" %*
