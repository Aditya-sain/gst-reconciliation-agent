@echo off
title GST Reconciliation Agent
echo ========================================================
echo     Starting GST Reconciliation AI Agent (Web App)
echo ========================================================
echo.
echo Opening server at http://localhost:5000 ...
start http://localhost:5000
python app.py
pause
