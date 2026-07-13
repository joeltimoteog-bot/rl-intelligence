@echo off
chcp 65001 >nul
title Subir cambios - RL Intelligence
cd /d C:\rl-intelligence

echo.
echo  ============================================
echo   SUBIR CAMBIOS A LAS DOS PAGINAS
echo   1) joeltimoteog-bot.github.io/rl-intelligence
echo   2) rl-intelligence.github.io
echo  ============================================
echo.

rem -- limpiar locks de git si quedaron colgados --
if exist .git\index.lock del /f .git\index.lock >nul 2>&1
if exist .git\HEAD.lock del /f .git\HEAD.lock >nul 2>&1

rem -- registrar el segundo remoto si no existe --
git remote get-url pagina >nul 2>&1
if errorlevel 1 git remote add pagina https://github.com/rl-intelligence/rl-intelligence.github.io.git

rem -- commit de lo que haya cambiado --
git add -A
git commit -m "update: cambios del sistema %date% %time%" >nul 2>&1
if errorlevel 1 (
  echo  [i] No hay cambios nuevos que confirmar. Subiendo lo pendiente...
) else (
  echo  [OK] Cambios confirmados.
)

echo.
echo  Subiendo a pagina 1 (joeltimoteog-bot)...
git push origin main
if errorlevel 1 (
  echo  [X] Fallo el push a origin. Revisa tu conexion o credenciales.
) else (
  echo  [OK] Pagina 1 actualizada.
)

echo.
echo  Subiendo a pagina 2 (rl-intelligence.github.io)...
git push pagina main:main --force
if errorlevel 1 (
  echo  [i] Reintentando contra la rama master...
  git push pagina main:master --force
  if errorlevel 1 (
    echo  [X] Fallo el push a la pagina 2. Inicia sesion con la cuenta duena de rl-intelligence.
  ) else (
    echo  [OK] Pagina 2 actualizada (rama master).
  )
) else (
  echo  [OK] Pagina 2 actualizada.
)

echo.
echo  Listo. Espera 1-2 minutos y refresca con Ctrl+Shift+R.
echo.
pause
