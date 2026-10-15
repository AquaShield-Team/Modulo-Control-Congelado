@echo off
title AquaShield - Control Congelado
echo ==================================================================
echo   AQUASHIELD - CONTROL CONGELADO: INICIO CON DATOS EN VIVO
echo ==================================================================
echo.
echo [1/2] Sincronizando con Azure Synapse Data Lake (SQL en vivo)...
python "%~dp0extraer_synapse.py"
echo.
echo [2/2] Cruzando datos con planillas, carteras de clientes y SAP...
python "%~dp0actualizar_datos.py"
echo.
echo Iniciando Servicio de Sincronizacion Web en segundo plano...
start /B python "%~dp0servidor_sync.py" >nul 2>&1
timeout /t 1 /nobreak >nul
echo Abriendo Dashboard en tu navegador...
start "" "http://localhost:8765"
exit
