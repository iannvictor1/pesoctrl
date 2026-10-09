@echo off
setlocal

if "%~1"=="" (
  echo Uso: scripts\limpar_logcat_coletor.bat SERIAL_OU_IP:5555
  exit /b 1
)

set ALVO=%~1

adb -s "%ALVO%" logcat -c
if errorlevel 1 (
  echo Falha ao limpar logcat do coletor %ALVO%.
  exit /b 1
)

echo Logcat limpo no coletor %ALVO%.
