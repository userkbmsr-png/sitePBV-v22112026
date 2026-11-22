@echo off
REM Compileaza roMail.vb cu vbc.exe inclus in Windows (nu necesita Visual Studio)
cd /d "%~dp0"
set VBC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\vbc.exe
if not exist "%VBC%" (
    echo EROARE: nu gasesc %VBC%
    pause
    exit /b 1
)
if not exist "FTP_Upload.vb" (
    echo EROARE: FTP_Upload.vb nu se afla in acelasi folder cu build_FTP_Upload.bat
    pause
    exit /b 1
)
if exist FTP_Upload.exe del FTP_Upload.exe
"%VBC%" /nologo /target:winexe /optimize+ /out:FTP_Upload.exe /win32icon:ftp.ico /reference:System.dll,System.Windows.Forms.dll,System.Drawing.dll,System.Security.dll FTP_Upload.vb
if exist FTP_Upload.exe (
    echo.
    echo Gata: %~dp0FTP_Upload.exe
) else (
    echo.
    echo EROARE: executabilul nu a fost creat. Vezi mesajele de mai sus.
)
pause
