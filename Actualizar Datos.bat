@echo off
title AquaShield - Actualizar Datos Control Congelado
echo ==================================================================
echo   AQUASHIELD - CONTROL CONGELADO: ACTUALIZACION AUTOMATICA
echo ==================================================================
echo.
echo [1/2] Consultando Azure Synapse Analytics (Data Lake)...
python "%~dp0extraer_synapse.py"
echo.
echo [2/2] Triangulando datos con Asignacion de Clientes y VF05N...
python "%~dp0actualizar_datos.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Ocurrio un inconveniente al procesar los archivos.
    pause
    exit /b %ERRORLEVEL%
)
echo.
echo ==================================================================
echo   DATOS ACTUALIZADOS CON EXITO. ABRIENDO DASHBOARD...
echo ==================================================================
timeout /t 2 /nobreak >nul
start "" "%~dp0index.html"
exit
