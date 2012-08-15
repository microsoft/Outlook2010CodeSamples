if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
set InstallDir=%ProgramFiles(x86)%\WrappedPST
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
set InstallDir=%ProgramFiles%\WrappedPST
goto END

:32BITOS
echo 32-bit Windows installed
set InstallDir=%ProgramFiles%\WrappedPST
goto END

:END
rundll32 "%InstallDir%\wrppst32.dll" RemoveFromMAPISVC
del  "%InstallDir%\wrppst32.dll"
del  "%InstallDir%\wrppst32.pdb"
