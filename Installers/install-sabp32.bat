set PRODUCT=SABP
set BINARY=sabp32
if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS

echo 64-bit Windows installed
if exist "%programfiles(x86)%\Microsoft Office\Office12\outlook.exe" goto 64OS32OL12
if exist "%programfiles(x86)%\Microsoft Office\Office14\outlook.exe" goto 64OS32OL14
if exist "%programfiles%\Microsoft Office\Office14\outlook.exe" goto 64OS64OL14
goto NOOFFICE

:64OS32OL12
echo Office 12 installed
set MAPISVCTARGET="%programfiles(x86)%\Common Files\System\MSMAPI\1033"
set InstallDir=%programfiles(x86)%\Microsoft Office\Office12
set SrcDir=..\Debug
goto END

:64OS32OL14
echo Office 14 32 bit installed
set MAPISVCTARGET="%programfiles(x86)%\Common Files\System\MSMAPI\1033"
set InstallDir=%programfiles(x86)%\Microsoft Office\Office14
set SrcDir=..\Debug
goto END

:64OS64OL14
echo Office 14 64 bit installed
set MAPISVCTARGET="%programfiles%\Common Files\System\MSMAPI\1033"
set InstallDir=%programfiles%\Microsoft Office\Office14
set SrcDir=..\x64\Debug
goto END

:32BITOS
echo 32-bit Windows installed
if exist "%programfiles%\Microsoft Office\Office12\outlook.exe" goto 32OL12
if exist "%programfiles%\Microsoft Office\Office14\outlook.exe" goto 32OL14
goto NOOFFICE

:32OL12
echo Office 12 installed
set MAPISVCTARGET="%programfiles%\Common Files\System\MSMAPI\1033"
set InstallDir=%programfiles%\Microsoft Office\Office12
set SrcDir=..\Debug
goto END

:32OL14
echo Office 14 installed
set MAPISVCTARGET="%programfiles%\Common Files\System\MSMAPI\1033"
set InstallDir=%programfiles%\Microsoft Office\Office14
set SrcDir=..\Debug
goto END

:INSTALL
copy mapisvc.inf %MAPISVCTARGET% /y
copy %SrcDir%\%BINARY%.dll %InstallDir% /y
copy %SrcDir%\%BINARY%.pdb %InstallDir% /y
goto :END

:NOOFFICE
echo Office not installed.

:END
