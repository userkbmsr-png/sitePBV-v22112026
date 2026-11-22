@echo off
REM Compileaza roMail.vb cu vbc.exe inclus in Windows (nu necesita Visual Studio)
cd /d "%~dp0"
set VBC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\vbc.exe
if not exist "%VBC%" (
    echo EROARE: nu gasesc %VBC%
    pause
    exit /b 1
)
if not exist "roMail.vb" (
    echo EROARE: roMail.vb nu se afla in acelasi folder cu build_romail.bat
    pause
    exit /b 1
)
if exist roMail.exe del roMail.exe
"%VBC%" /nologo /target:winexe /optimize+ /out:roMail.exe /win32icon:mail.ico /reference:System.dll,System.Drawing.dll,System.Windows.Forms.dll roMail.vb
if exist roMail.exe (
    echo.
    echo Gata: %~dp0roMail.exe
) else (
    echo.
    echo EROARE: executabilul nu a fost creat. Vezi mesajele de mai sus.
)
pause
