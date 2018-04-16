set PRODUCT=MRXP
set BINARY=mrxp32
if "%programfiles(x86)%XXX"=="XXX" goto 32BITOS
if exist "C:\Program Files\Microsoft Office\Office14\outlook.exe" goto 64BITOUTLOOK
if exist "C:\Program Files\Microsoft Office\Office15\outlook.exe" goto 64BITOUTLOOK
if exist "C:\Program Files\Microsoft Office\Office16\outlook.exe" goto 64BITOUTLOOK

echo 64-bit Windows and 32 bit Office installed
set InstallDir=%ProgramFiles(x86)%\%PRODUCT%
set SrcDir=..\Debug
goto INSTALL

:64BITOUTLOOK
echo 64-bit Windows and 64-bit Office installed
set InstallDir= %ProgramFiles%\%PRODUCT%
set SrcDir=..\x64\Debug
goto INSTALL

:32BITOS
echo 32-bit Windows installed
set InstallDir=%ProgramFiles%\%PRODUCT%
set SrcDir=..\Debug
goto INSTALL

:INSTALL
md "%InstallDir%"
copy "%SrcDir%\%BINARY%.dll" "%InstallDir%"
copy "%SrcDir%\%BINARY%.pdb" "%InstallDir%"
rundll32 "%InstallDir%\%BINARY%.dll" MergeWithMAPISVC
