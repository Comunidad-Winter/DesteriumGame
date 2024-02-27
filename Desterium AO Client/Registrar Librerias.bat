@echo off
cls
echo 	************************************************
echo 	**************** Nightmare AO ******************
echo 	************************************************
echo.
echo    IMPORTANTE: 
echo 	Si esta utilizando Windows Vista o 7, tiene que hacer click derecho
echo 	sobre "Registrar Liberias" echo y seleccionar 
echo 	"Ejecutar como Administrador" para que haga efecto.
echo.
pause

echo Registrando AAMD532.DLL
regsvr32 AAMD532.DLL -s

echo Registrando FM20.DLL
regsvr32 FM20.DLL -s

echo Registrando CSWSK32.OCX
regsvr32 CSWSK32.OCX -s

echo Registrando MSINET.OCX
regsvr32 MSINET.OCX -s

echo Registrando MSWINSCK.OCX
regsvr32 MSWINSCK.OCX -s

echo Registrando RICHTX32.OCX
regsvr32 RICHTX32.OCX -s

echo Registrando UNZIP32.DLL
regsvr32 UNZIP32.DLL -s

echo Registrando VBABDX.DLL
regsvr32 VBABDX.DLL -s

echo Registrando VBALPROGBAR6.OCX
regsvr32 VBALPROGBAR6.OCX -s

echo Registrando VBDABL.DLL
regsvr32 VBDABL.DLL -s

echo Registrando MSVBVM50.DLL
regsvr32 MSVBVM50.DLL -s

echo.
echo Librerias registradas
echo.
pause
exit