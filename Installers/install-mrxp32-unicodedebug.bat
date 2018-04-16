if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
copy .\Unicode_Debug\mrxp32.dll %windir%\sysWow64 /y
copy .\Unicode_Debug\mrxp32.pdb %windir%\sysWow64 /y
rundll32 %windir%\syswow64\mrxp32.dll MergeWithMAPISVC
goto END

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
copy x64\Unicode_Debug\mrxp32.dll %windir%\system32 /y
copy x64\Unicode_Debug\mrxp32.pdb %windir%\system32 /y
rundll32 %windir%\system32\mrxp32.dll MergeWithMAPISVC
goto END

:32BITOS
echo 32-bit Windows installed
copy .\Unicode_Debug\mrxp32.dll %windir%\system32 /y
copy .\Unicode_Debug\mrxp32.pdb %windir%\system32 /y
rundll32 %windir%\system32\mrxp32.dll MergeWithMAPISVC

:END

