@ECHO ON
title QM Start

cd /D %~dp0\

%~dp0\venv\Scripts\python.exe main.py

cmd.exe