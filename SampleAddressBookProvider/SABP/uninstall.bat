if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
del "%programfiles(x86)%\Microsoft Office\Office12\sabp32.dll"
del "%programfiles(x86)%\Microsoft Office\Office12\sabp32.pdb"
del "%programfiles(x86)%\Common Files\System\MSMAPI\1033\mapisvc.inf"
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
del "%programfiles%\Microsoft Office\Office12\sabp32.dll"
del "%programfiles%\Microsoft Office\Office12\sabp32.pdb"
del "%programfiles%\Common Files\System\MSMAPI\1033\mapisvc.inf"
goto END

:32BITOS
echo 32-bit Windows installed
del "%programfiles%\Microsoft Office\Office12\sabp32.dll"
del "%programfiles%\Microsoft Office\Office12\sabp32.pdb"
del "%programfiles%\Common Files\System\MSMAPI\1033\mapisvc.inf"

:END
