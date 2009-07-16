if "%programfiles(x86)%XXX"=="XXX" goto 32BIT
echo 64-bit Windows installed
copy .\unicode_debug\mrxp32.dll %windir%\sysWow64
copy .\unicode_debug\mrxp32.pdb %windir%\sysWow64
%windir%\syswow64\rundll32.exe %windir%\syswow64\mrxp32.dll MergeWithMAPISVC
goto END

:32BIT
echo 32-bit Windows installed
copy .\unicode_debug\mrxp32.dll %windir%\system32
copy .\unicode_debug\mrxp32.pdb %windir%\system32
%windir%\system32\rundll32.exe %windir%\system32\mrxp32.dll MergeWithMAPISVC

:END

