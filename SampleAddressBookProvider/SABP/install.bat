if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
copy mapisvc.inf "%programfiles(x86)%\Common Files\System\MSMAPI\1033" /y
copy Debug\sabp32.dll "%programfiles(x86)%\Microsoft Office\Office12\" /y
copy Debug\sabp32.pdb "%programfiles(x86)%\Microsoft Office\Office12\" /y
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
copy mapisvc.inf "%programfiles%\Common Files\System\MSMAPI\1033" /y
copy x64\Debug\sabp32.dll %windir%\system32
copy x64\Debug\sabp32.pdb %windir%\system32
goto END

:32BITOS
echo 32-bit Windows installed
copy mapisvc.inf "%programfiles%\Common Files\System\MSMAPI\1033" /y
copy Debug\sabp32.dll "%programfiles%\Microsoft Office\Office12\" /y
copy Debug\sabp32.pdb "%programfiles%\Microsoft Office\Office12\" /y

:END
