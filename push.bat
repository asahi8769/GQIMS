@ECHO ON
title push Start

cd /D %~dp0\

%~dp0\venv\Scripts\python.exe _push.py

cmd.exe