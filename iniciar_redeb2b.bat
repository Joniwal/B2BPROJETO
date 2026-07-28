@echo off
REM ==========================================================================
REM Atalho para iniciar o painel REDEB2B com um duplo-clique.
REM Ativa o ambiente virtual (venv), inicia o servidor Flask e abre o
REM navegador automaticamente em http://localhost:5000
REM ==========================================================================

cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo O ambiente virtual "venv" nao foi encontrado nesta pasta.
    echo Rode primeiro: python -m venv venv
    echo E depois: venv\Scripts\activate ^&^& python -m pip install -r requirements.txt
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

REM Abre o navegador depois de 2 segundos (tempo do Flask subir)
start "" cmd /c "timeout /t 2 >nul && start http://localhost:5000"

python app.py

pause
