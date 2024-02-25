set mypath=%~dp0
echo %mypath%
echo sysPath %SystemRoot%
copy "%mypath%\Aurora.Network.dll" %SystemRoot%\SysWOW64
REGSVR32 /s %SystemRoot%\SysWOW64\Aurora.Network.dll

PAUSE