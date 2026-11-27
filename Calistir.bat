@echo off
rem Gelen klasorundeki ZIP'leri hemen isler. Sadece kontrol icin: Calistir.bat -Deneme
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0FaturaAktarim.ps1" %*
echo.
pause
