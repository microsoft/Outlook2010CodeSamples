if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK
if exist "C:\Program Files\Microsoft Office\Office15\outlook.exe" goto 64BITOUTLOOK
if exist "C:\Program Files\Microsoft Office\Office16\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
rundll32 %windir%\syswow64\mrxp32.dll RemoveFromMAPISVC
del %windir%\sysWow64\mrxp32.dll
del %windir%\sysWow64\mrxp32.pdb
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
rundll32 %windir%\system32\mrxp32.dll RemoveFromMAPISVC
del %windir%\system32\mrxp32.dll
del %windir%\system32\mrxp32.pdb
goto END

:32BITOS
echo 32-bit Windows installed
rundll32 %windir%\system32\mrxp32.dll RemoveFromMAPISVC
del %windir%\system32\mrxp32.dll
del %windir%\system32\mrxp32.pdb

:END

