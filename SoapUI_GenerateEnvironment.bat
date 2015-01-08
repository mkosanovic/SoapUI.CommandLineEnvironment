@echo off

REM ===============================================================================
REM I made this script to enable acces to soapUI tools without adding them to path
REM ===============================================================================

REM if soap ui is installed for current user modify the shortcut
dir "%APPDATA%\Microsoft\Windows\Start Menu\Programs\SmartBear" >nul 2>&1 && (
	for /f "tokens=*" %%i IN ('DIR /s /b "%APPDATA%\Microsoft\Windows\Start Menu\Programs\SmartBear\" ^| find "SoapUI-"') DO SET SOAPLINKDIR=%%i
	call :CreateReadSoapDir %SOAPLINKDIR%
)

REM if soap ui is installed for all users modify this shortcut
dir "%ALLUSERSPROFILE%\Microsoft\Windows\StartMenu\Programs\SmartBear" >nul 2>&1 && (
		for /f "tokens=*" %%i IN ('DIR /s /b "%ALLUSERSPROFILE%\Microsoft\Windows\StartMenu\Programs\SmartBear" ^| find "SoapUI-"')  DO SET SOAPLINKDIR=%%i	
		call :CreateReadSoapDir %SOAPLINKDIR%
		REM for /f "tokens=*" %%i IN ('DIR /s /b "%ALLUSERSPROFILE%\Microsoft\Windows\StartMenu\Programs\SmartBear" ^| find "SoapUI-"') DO (call :CreateReadSoapDir %%i)	
)

GOTO :eof

REM creates Read soap directory
:CreateReadSoapDir
echo Set oWs = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo Set oLink  = oWs.CreateShortcut("%SOAPLINKDIR%") >> CreateShortcut.vbs
echo Set objFSO = CreateObject("Scripting.FileSystemObject") >> CreateShortcut.vbs
echo Set objFolder = objFSO.GetFile(oLink.TargetPath) >> CreateShortcut.vbs
echo WScript.Echo objFolder.ParentFolder >> CreateShortcut.vbs

REM cscript /nologo CreateShortcut.vbs

REM execute vbscript
REM read target path from output 
for /f "tokens=*" %%i IN ('cscript /nologo CreateShortcut.vbs') DO SET SOAPDIR=%%i

REM create variables.bat
echo %SOAPDIR% > test

echo @echo off > "%SOAPDIR%\soapuivars.bat"
echo rem set soapui bats to path >> "%SOAPDIR%\soapuivars.bat"
echo SET PATH=%%~dp0;%%PATH%% >> "%SOAPDIR%\soapuivars.bat"
echo SETLOCAL enabledelayedexpansion >> "%SOAPDIR%\soapuivars.bat"
echo pushd "%%~dp0" >>"%SOAPDIR%\soapuivars.bat"
echo echo Environment has been modified to use soapui >> "%SOAPDIR%\soapuivars.bat"
echo popd >> "%SOAPDIR%\soapuivars.bat"
echo endlocal >> "%SOAPDIR%\soapuivars.bat"
echo cd "%%~dp0" >> "%SOAPDIR%\soapuivars.bat"

Set targetPath= "%%windir%%\system32\cmd.exe /k %SOAPDIR%\soapuivars.bat"

REM create shortcut script
echo Set oWs = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo Set objFSO = CreateObject("Scripting.FileSystemObject") >> CreateShortcut.vbs
echo Set objFolder = objFSO.GetFile("%SOAPLINKDIR%") >> CreateShortcut.vbs
echo WScript.Echo objFolder.ParentFolder >> CreateShortcut.vbs
echo Set oLink = oWs.CreateShortcut(objFolder.ParentFolder ^& "\test.lnk")>> CreateShortcut.vbs
REM echo oLink.TargetPath ="%%windir%%\system32\cmd.exe /k %SOAPDIR%\soapuivars.bat" >> CreateShortcut.vbs
echo oLink.TargetPath ="%%windir%%\system32\cmd.exe" >> CreateShortcut.vbs
echo oLink.Arguments = "/k ""%SOAPDIR%\soapuivars.bat""" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs

REM create shortcut in start menu
cscript /nologo CreateShortcut.vbs

REM delete CreateShortcut
del CreateShortcut.vbs

GOTO :eof