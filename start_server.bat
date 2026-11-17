@echo off
title PICT Marks Portal Local Server
echo Starting Local Web Server at http://localhost:8000 ...
start "" "http://localhost:8000"
python -m http.server 8000
pause
