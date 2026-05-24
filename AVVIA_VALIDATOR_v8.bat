@echo off
cd /d "%~dp0"
title Trading Strategy Validator v8
echo.
echo =====================================================
echo   TRADING STRATEGY VALIDATOR v8
echo   Framework Morrone + CSCV + Neff (PCA) + Archivio
echo   Limite upload: 2 GB  (ZIP supportato)
echo =====================================================
echo.
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERRORE: Python non trovato nel PATH.
    echo  Installa Python da https://www.python.org
    pause & exit /b 1
)
python -m pip install streamlit pandas numpy scipy matplotlib openpyxl python-docx --quiet
echo.
echo  Avvio su: http://localhost:8501
echo  Per chiudere: premi CTRL+C
echo.
python -m streamlit run trading_validator_v8.py ^
    --server.headless false ^
    --browser.gatherUsageStats false ^
    --server.maxUploadSize 2048
pause
