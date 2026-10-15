@echo off
title AquaShield - Publicar Actualizacion en GitHub
echo ==================================================================
echo   AQUASHIELD: PUBLICAR ACTUALIZACION EN GITHUB PAGES
echo ==================================================================
echo.
cd /d "%~dp0"
echo [1/3] Preparando cambios...
git add data_control_congelado.js reglas_clientes.json index.html README.md
echo [2/3] Registrando version...
git commit -m "Actualizacion de base de datos COMEX (%date% %time%)"
echo [3/3] Subiendo a GitHub...
git push origin main
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==================================================================
    echo   PUBLICACION EXITOSA!
    echo   Tu equipo ya tiene la data actualizada en:
    echo   https://aquashield-team.github.io/Modulo-Control-Congelado/
    echo ==================================================================
) else (
    echo.
    echo [ERROR] Hubo un problema al subir a GitHub. Verifica tu conexion.
)
echo.
pause
