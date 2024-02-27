@echo off
cls
echo *********************************************
echo *********************************************
echo *** Registrar Librerias de Porvoo AO **
echo *********************************************
echo *********************************************
echo ** Creado por ^[humB]^ ** 

echo *********************************************
echo.
echo NOTA: 
echo Si esta utilizando Windows Vista o 7, tiene que hacer segundo click 
echo sobre "Registrar Liberias" echo y seleccionar 
echo "Ejecutar como Administrador" para que haga efecto.
echo.
pause
echo Registrando RICHTX32.OCX (Microsoft Rich Textbox Control)...
regsvr32 RICHTX32.OCX
echo Registrando MSWINSCK.OCX (Microsoft Winsock Control)...
regsvr32 MSWINSCK.OCX
echo Registrando MSINET.OCX (Microsoft Internet Transfer Control)...
regsvr32 MSINET.OCX
echo Registrando COMCTL32.OCX (Microsoft Windows Common Controls 5.0)...
regsvr32 COMCTL32.OCX
echo Registrando MSCOMCTL.OCX (Microsoft Windows Common Controls 6.0)...
regsvr32 MSCOMCTL.OCX
echo Registrando COMDLG32.OCX (Microsoft Common Dialog Control)...
regsvr32 COMDLG32.OCX
echo Registrando CSWSK32.OCX (Catalyst SocketWrench Control)...
regsvr32 CSWSK32.OCX
echo Registrando FLASH10D.OCX (Microsoft flash Control)...
regsvr32 FLASH10D.OCX
echo Registrando VBALPROGBAR6.OCX (Progress Bar for Visual Basic Type Library)...
regsvr32 vbalProgBar6.OCX
echo Registrando DX7VB.DLL (DirectX 7 for Visual Basic Type Library)...
regsvr32 DX7VB.DLL
echo Registrando MSSTDFMT.DLL (Microsoft Data Formatting Object Library)...
regsvr32 MSSTDFMT.DLL
echo Registrando SCRRUN.DLL (Microsoft Scripting Runtime)...
regsvr32 SCRRUN.DLL
echo.
echo Registro terminado...
echo.
pause
exit