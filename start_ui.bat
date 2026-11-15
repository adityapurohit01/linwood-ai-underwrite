@echo off
title Linwood Capital — AI Underwriting Engine
echo ======================================================================
echo           LINWOOD CAPITAL -- AI UNDERWRITING ENGINE
echo ======================================================================
echo.
echo [1/2] Launching AI Underwriting Engine Web Portal...
echo [2/2] Opening your browser at http://localhost:8000 ...
echo.
timeout /t 2 /nobreak >nul
start "" http://localhost:8000
python app.py
pause
