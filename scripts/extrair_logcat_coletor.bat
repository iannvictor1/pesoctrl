@echo off
setlocal

if "%~1"=="" (
  echo Uso: scripts\extrair_logcat_coletor.bat SERIAL_OU_IP:5555 [arquivo_saida]
  exit /b 1
)

set ALVO=%~1
set SAIDA=%~2

if "%SAIDA%"=="" (
  set SAIDA=logcat_%ALVO::=_%.txt
)

adb -s "%ALVO%" logcat -d -v time -s DCS-Wedge > "%SAIDA%"
if errorlevel 1 (
  echo Falha ao extrair logcat do coletor %ALVO%.
  exit /b 1
)

echo Logcat salvo em "%SAIDA%".
