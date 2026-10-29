@echo off
title Dilek Ofis Fiyat Katalogu
cd /d "%~dp0"

rem ZIP'in icinden calistirilirsa server.js bulunamaz
if exist "%~dp0server.js" goto dosyalar_tamam
echo.
echo  HATA: Program dosyalari bulunamadi.
echo  ZIP dosyasina sag tiklayip "Tumunu ayikla" deyin,
echo  sonra ayiklanan klasorun icindeki baslat.bat dosyasini calistirin.
echo.
pause
exit /b
:dosyalar_tamam

set "NODE_VER=22.12.0"
set "NODE_ARCH=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "NODE_ARCH=arm64"
set "NODE_EXE="

rem 1) Daha once bu klasore indirilmis Node.js
if exist "%~dp0node\node.exe" set "NODE_EXE=%~dp0node\node.exe"
if defined NODE_EXE goto calistir

rem 2) Bilgisayarda kurulu Node.js
for /f "delims=" %%i in ('where node 2^>nul') do if not defined NODE_EXE set "NODE_EXE=%%i"
if defined NODE_EXE goto calistir
if exist "%ProgramFiles%\nodejs\node.exe" set "NODE_EXE=%ProgramFiles%\nodejs\node.exe"
if defined NODE_EXE goto calistir
if exist "%ProgramFiles(x86)%\nodejs\node.exe" set "NODE_EXE=%ProgramFiles(x86)%\nodejs\node.exe"
if defined NODE_EXE goto calistir
if exist "%LOCALAPPDATA%\Programs\nodejs\node.exe" set "NODE_EXE=%LOCALAPPDATA%\Programs\nodejs\node.exe"
if defined NODE_EXE goto calistir

rem 3) Hicbiri yoksa Node.js'i bu klasore otomatik indir (kurulum gerekmez)
echo.
echo  Node.js bulunamadi. Simdi otomatik indiriliyor (yaklasik 35 MB)...
echo  Lutfen bekleyin, bu islem sadece ilk seferde yapilir.
echo.
set "NODE_DIR=node-v%NODE_VER%-win-%NODE_ARCH%"
set "NODE_ZIP=%TEMP%\%NODE_DIR%.zip"
curl -L --fail -o "%NODE_ZIP%" "https://nodejs.org/dist/v%NODE_VER%/%NODE_DIR%.zip"
if errorlevel 1 goto indirme_hatasi
if exist "%~dp0%NODE_DIR%" rmdir /s /q "%~dp0%NODE_DIR%"
tar -xf "%NODE_ZIP%" -C "%~dp0."
if errorlevel 1 goto indirme_hatasi
if exist "%~dp0node" rmdir /s /q "%~dp0node"
ren "%~dp0%NODE_DIR%" node
del "%NODE_ZIP%" >nul 2>nul
if not exist "%~dp0node\node.exe" goto indirme_hatasi
set "NODE_EXE=%~dp0node\node.exe"
echo  Node.js hazir.

:calistir
rem Telefonlarin baglanabilmesi icin guvenlik duvari izni (sadece ilk seferde sorar)
netsh advfirewall firewall show rule name=DilekOfisKatalog >nul 2>nul
if not errorlevel 1 goto baslat
echo.
echo  Telefonlarin fiyat sayfasini acabilmesi icin Windows guvenlik duvari izni gerekiyor.
echo  Birazdan acilacak pencerede "Evet" deyin.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0guvenlik-duvari.ps1"
:baslat
echo.
echo  Guncellemeler kontrol ediliyor...
if exist "%~dp0guncelle.ps1" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0guncelle.ps1"
echo.
echo  Program baslatiliyor, yonetim paneli tarayicida kendiliginden acilacak...
"%NODE_EXE%" server.js
echo.
echo  Program kapandi. Yukaridaki mesajin fotografini cekip gonderin.
echo  (Hata ayrintisi: hata-kaydi.txt dosyasinda)
pause
exit /b

:indirme_hatasi
echo.
echo  Node.js otomatik indirilemedi. Internet baglantinizi kontrol edin.
echo  Elle kurmak icin: https://nodejs.org adresinden "LTS" surumunu indirip kurun,
echo  sonra bilgisayari yeniden baslatip baslat.bat dosyasini tekrar calistirin.
echo.
start "" https://nodejs.org/tr/download
pause
