if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
copy Debug\wrppst32.dll %windir%\syswow64
copy Debug\wrppst32.pdb %windir%\syswow64
rundll32 %windir%\syswow64\wrppst32.dll MergeWithMAPISVC
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
copy x64\Debug\wrppst32.dll %windir%\system32
copy x64\Debug\wrppst32.pdb %windir%\system32
rundll32 %windir%\system32\wrppst32.dll MergeWithMAPISVC
goto END

:32BITOS
echo 32-bit Windows installed
copy Debug\wrppst32.dll %windir%\system32
copy Debug\wrppst32.pdb %windir%\system32
rundll32 %windir%\system32\wrppst32.dll MergeWithMAPISVC

:END

