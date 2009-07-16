if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows installed
rundll32 %windir%\syswow64\wrppst32.dll RemoveFromMAPISVC
del  %windir%\syswow64\wrppst32.dll
del  %windir%\syswow64\wrppst32.pdb
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
rundll32 %windir%\system32\wrppst32.dll RemoveFromMAPISVC
del  %windir%\system32\wrppst32.dll
del  %windir%\system32\wrppst32.pdb
goto END

:32BITOS
echo 32-bit Windows installed
rundll32 %windir%\system32\wrppst32.dll RemoveFromMAPISVC
del  %windir%\system32\wrppst32.dll
del  %windir%\system32\wrppst32.pdb

:END


