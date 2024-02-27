Attribute VB_Name = "ProtocolCmdParse"
'Desterium AO
'
'Copyright (C) 2006 Juan Martín Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Desterium AO is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'

Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
10        If LenB(UserName) = 0 Then Exit Sub
          
          Dim i As Long
          Dim nameLength As Long
          
20        If (InStrB(UserName, "+") <> 0) Then
30            UserName = Replace$(UserName, "+", " ")
40        End If
          
50        UserName = UCase$(UserName)
60        nameLength = Len(UserName)
          
70        i = 1
80        Do While i <= LastChar
90            If UCase$(charlist(i).Nombre) = UserName Or _
                  UCase$(Left$(charlist(i).Nombre, nameLength + 2)) = UserName & " <" Then
100               Exit Do
110           Else
120               i = i + 1
130           End If
140       Loop
          
150       If i <= LastChar Then
160           Call WriteWhisper(i, Mensaje)
170       End If
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modification: 16/11/2009
      'Interpreta, valida y ejecuta el comando ingresado
      '26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
      '16/11/2009: ZaMa - Ahora el /ct admite radio
      '***************************************************
          Dim TmpArgos() As String
          
          Dim Comando As String
          Dim ArgumentosAll() As String
          Dim ArgumentosRaw As String
          Dim Argumentos2() As String
          Dim Argumentos3() As String
          Dim Argumentos4() As String
          Dim CantidadArgumentos As Long
          Dim notNullArguments As Boolean
          
          Dim tmpArr() As String
          Dim tmpInt As Integer
          
          ' TmpArgs: Un array de a lo sumo dos elementos,
          ' el primero es el comando (hasta el primer espacio)
          ' y el segundo elemento es el resto. Si no hay argumentos
          ' devuelve un array de un solo elemento
10        TmpArgos = Split(RawCommand, " ", 2)
          
20        Comando = Trim$(UCase$(TmpArgos(0)))
          
30        If UBound(TmpArgos) > 0 Then
              ' El string en crudo que este despues del primer espacio
40            ArgumentosRaw = TmpArgos(1)
              
              'veo que los argumentos no sean nulos
50            notNullArguments = LenB(Trim$(ArgumentosRaw))
              
              ' Un array separado por blancos, con tantos elementos como
              ' se pueda
60            ArgumentosAll = Split(TmpArgos(1), " ")
              
              ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
70            CantidadArgumentos = UBound(ArgumentosAll) + 1
              
              ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
              ' 2, 3 y 4 elementos respectivamente. Eso significa
              ' que pueden tener menos, por lo que es imperativo
              ' preguntar por CantidadArgumentos.
              
80            Argumentos2 = Split(TmpArgos(1), " ", 2)
90            Argumentos3 = Split(TmpArgos(1), " ", 3)
100           Argumentos4 = Split(TmpArgos(1), " ", 4)
110       Else
120           CantidadArgumentos = 0
130       End If
          
          ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
140       If Comando = "" Then Comando = " "
          
150       If Left$(Comando, 1) = "/" Then
              ' Comando normal
              
160           Select Case Comando
              
                  Case "/VERP"
170                  If notNullArguments Then
180                       Call WriteLookProcess(ArgumentosRaw)
190                   Else
                          'Avisar que falta el parametro
200                                                                With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
210                        Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
220                        End With
230                   End If

240           Case "/CAER"
250                   Call writeDropItems
              
260               Case "/SEG"
270                   Call WriteSafeToggle
              
280               Case "/ONLINE"
290                   Call WriteOnline
                      
300                   Case "/SUBIRFAMA"
310                   If UserEstado = 1 Then 'Muerto
320                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
330                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
340                       End With
350                       Exit Sub
360                   End If
370                   Call Writeusarbono
                      
380                   Case "/PREMIUM"
390                   If UserEstado = 1 Then 'Muerto
400                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
410                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
420                       End With
430                       Exit Sub
440                   End If
450                   Call WritePremium
                      
460                   Case "/ORO"
470                   If UserEstado = 1 Then 'Muerto
480                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
490                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
500                       End With
510                       Exit Sub
520                   End If
530                   Call WriteOro
                      
540                                   Case "/PLATA"
550                   If UserEstado = 1 Then 'Muerto
560                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
570                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
580                       End With
590                       Exit Sub
600                   End If
610                   Call WritePlata
                      
620                                   Case "/BRONCE"
630                   If UserEstado = 1 Then 'Muerto
640                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
650                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
660                       End With
670                       Exit Sub
680                   End If
690                   Call WriteBronce
                      
700               Case "/SALIR"
                 '  With FontTypes(FontTypeNames.FONTTYPE_INFO)
                  'Call ShowConsoleMsg("Gracias por jugar Desterium AO.", .red, .green, .blue, .bold, .italic)
                 ' End With
710                   If UserParalizado Then 'Inmo
720                       With FontTypes(FontTypeNames.FONTTYPE_WARNING)
730                           Call _
                                  ShowConsoleMsg("No puedes salir estando paralizado.", _
                                  .red, .green, .blue, .bold, .italic)
740                       End With
750                       Exit Sub
760                   End If
770                   If frmMain.macrotrabajo.Enabled Then Call _
                          frmMain.DesactivarMacroTrabajo
780                   Call WriteQuit
                      
790               Case "/SALIRCLAN"
800                   Call WriteGuildLeave
                      
810               Case "/BALANCE"
820                   If UserEstado = 1 Then 'Muerto
830                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
840                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
850                       End With
860                       Exit Sub
870                   End If
880                   Call WriteRequestAccountState
                      
890               Case "/QUIETO"
900                   If UserEstado = 1 Then 'Muerto
910                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
920                           Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
930                       End With
940                       Exit Sub
950                   End If
960                   Call WritePetStand
                      
970               Case "/ACOMPAÑAR"
980                   If UserEstado = 1 Then 'Muerto
990                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
1000                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1010                      End With
1020                      Exit Sub
1030                  End If
1040                  Call WritePetFollow
                      
1050              Case "/LIBERAR"
1060                  If UserEstado = 1 Then 'Muerto
1070                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1080                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1090                      End With
1100                      Exit Sub
1110                  End If
1120                  Call WriteReleasePet
                      
1130              Case "/ENTRENAR"
1140                  If UserEstado = 1 Then 'Muerto
1150                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1160                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1170                      End With
1180                      Exit Sub
1190                  End If
1200                  Call WriteTrainList
                      
1210              Case "/DESCANSAR"
1220                  If UserEstado = 1 Then 'Muerto
1230                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1240                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1250                      End With
1260                      Exit Sub
1270                  End If
1280                  Call WriteRest
                     
                      
1290              Case "/CAPTIONS"
1300                  If notNullArguments Then
1310                      WriteRequieredCaptions ArgumentosRaw
1320                  Else
                      
1330                  End If
                      
1340              Case "/FIANZA"
1350                  If UserEstado = 1 Then 'Muerto
1360                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1370                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1380                      End With
1390                      Exit Sub
1400                  End If
                     
1410                  If notNullArguments Then
1420                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
1430                          Call WriteFianzah(ArgumentosRaw)
1440                      Else
                              'No es numerico
1450                          Call _
                                  ShowConsoleMsg("Cantidad incorecta. Utilice /Fianza CANTIDAD.")
1460                      End If
1470                  Else
                          'Avisar que falta el parametro
1480                      Call _
                              ShowConsoleMsg("Faltan paramtetros. Utilice /Fianza CANTIDAD.")
1490                  End If
                      
1500              Case "/MEDITAR"
1510                  If UserEstado = 1 Then 'Muerto
1520                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1530                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1540                      End With
1550                      Exit Sub
1560                  End If
1570                  Call WriteMeditate

1580              Case "/VERPENAS"
1590                  If notNullArguments Then
1600                      Call WriteVerpenas(ArgumentosRaw)
1610                  Else
                          'Avisar que falta el parametro
1620                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")
1630                  End If

1640              Case "/CONSULTA"
1650                  Call WriteConsulta
                  
1660              Case "/RESUCITAR"
1670                  Call WriteResucitate
                      
1680              Case "/CURAR"
1690                  Call WriteHeal
                                    
1700              Case "/EST"
1710                  Call WriteRequestStats
                  
1720              Case "/AYUDA"
1730                  Call WriteHelp
                      
                      
1740              Case "/COMERCIAR"
                         
1750                  If UserEstado = 1 Then 'Muerto
1760                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1770                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1780                      End With
1790                      Exit Sub
                      
1800                  ElseIf Comerciando Then 'Comerciando
1810                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1820                          Call ShowConsoleMsg("Ya estás comerciando", .red, _
                                  .green, .blue, .bold, .italic)
1830                      End With
1840                      Exit Sub
1850                  End If
1860                  Call WriteCommerceStart
                      
1870              Case "/BOVEDA"
1880                  If UserEstado = 1 Then 'Muerto
1890                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
1900                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
1910                      End With
1920                      Exit Sub
1930                  End If
1940                  Call WriteBankStart
                      
1950              Case "/ENLISTAR"
1960                  Call WriteEnlist
                          
1970              Case "/INFORMACION"
1980                  Call WriteInformation
                      
1990              Case "/RECOMPENSA"
2000                  Call WriteReward
                      
2010              Case "/UPTIME"
2020                  Call WriteUpTime
                  
2210              Case "/COMPARTIRNPC"
2220                  If UserEstado = 1 Then 'Muerto
2230                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
2240                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
2250                      End With
2260                      Exit Sub
2270                  End If
                      
2280                  Call WriteShareNpc
                      
2290              Case "/NOCOMPARTIRNPC"
2300                  If UserEstado = 1 Then 'Muerto
2310                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
2320                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
2330                      End With
2340                      Exit Sub
2350                  End If
                      
2360                  Call WriteStopSharingNpc
                      
2370              Case "/ENCUESTA"
2380                  If CantidadArgumentos = 0 Then
                          ' Version sin argumentos: Inquiry
2390                      Call WriteInquiry
2400                  Else
                          ' Version con argumentos: InquiryVote
2410                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Byte) Then
2420                          Call WriteInquiryVote(ArgumentosRaw)
2430                      Else
                              'No es numerico
2440                          Call _
                                  ShowConsoleMsg("Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1.")
2450                      End If
2460                  End If
              
2470              Case "/CMSG"
                      'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
2480                  If CantidadArgumentos > 0 Then
2490                      Call WriteGuildMessage(ArgumentosRaw)
2500                  Else
                          'Avisar que falta el parametro
2510                      Call ShowConsoleMsg("Escriba un mensaje.")
2520                  End If
              
2530              Case "/PMSG"
                      'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
2540                  If CantidadArgumentos > 0 Then
2550                      Call WriteGroupMessage(ArgumentosRaw)
2560                  Else
                          'Avisar que falta el parametro
2570                      Call ShowConsoleMsg("Escriba un mensaje.")
2580                  End If
                  
2590              Case "/CENTINELA"
2600                  If notNullArguments Then
2610                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) _
                              Then
2620                          Call WriteCentinelReport(CInt(ArgumentosRaw))
2630                      Else
                              'No es numerico
2640                          Call _
                                  ShowConsoleMsg("El código de verificación debe ser numerico. Utilice /centinela X, siendo X el código de verificación.")
2650                      End If
2660                  Else
                          'Avisar que falta el parametro
2670                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /centinela X, siendo X el código de verificación.")
2680                  End If
              
2690              Case "/ONLINECLAN"
2700                  Call WriteGuildOnline
                      
                  
                      
2730              Case "/BMSG"
2740                  If notNullArguments Then
2750                      Call WriteCouncilMessage(ArgumentosRaw)
2760                  Else
                          'Avisar que falta el parametro
2770                      Call ShowConsoleMsg("Escriba un mensaje.")
2780                  End If
                      
2790              Case "/ROL"
2800                  If notNullArguments Then
2810                      Call WriteRoleMasterRequest(ArgumentosRaw)
2820                  Else
                          'Avisar que falta el parametro
2830                      Call ShowConsoleMsg("Escriba una pregunta.")
2840                  End If
                      
2850              Case "/GM"
2860                  Call WriteGMRequest

2870              Case "/DESC"
2880                  If UserEstado = 1 Then 'Muerto
2890                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
2900                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
2910                      End With
2920                      Exit Sub
2930                  End If
                      
2940                  Call WriteChangeDescription(ArgumentosRaw)
                  
2950              Case "/VOTO"
2960                  If notNullArguments Then
2970                      Call WriteGuildVote(ArgumentosRaw)
2980                  Else
                          'Avisar que falta el parametro
2990                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")
3000                  End If
                     
3010             Case "/PENAS"
3020                 WritePunishments UserName
                      
3030              Case "/CONTRASEÑA"
3040                  Call frmNewPassword.Show(vbModal, frmMain)
                  
                  
3050              Case "/APOSTAR"
3060                  If UserEstado = 1 Then 'Muerto
3070                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
3080                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
3090                      End With
3100                      Exit Sub
3110                  End If
3120                  If notNullArguments Then
3130                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) _
                              Then
3140                          Call WriteGamble(ArgumentosRaw)
3150                      Else
                              'No es numerico
3160                          Call _
                                  ShowConsoleMsg("Cantidad incorrecta. Utilice /apostar CANTIDAD.")
3170                      End If
3180                  Else
                          'Avisar que falta el parametro
3190                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /apostar CANTIDAD.")
3200                  End If
                      
                      
3210                              Case "/ABANDONAR"
3220                  If UserEstado = 1 Then 'Muerto
3230                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
3240                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
3250                      End With
3260                      Exit Sub
3270                  End If
                      
3280                  Call WriteLeaveFaction
           
3290                  Case "/RETIRARTODO"
3300                  Call ParseUserCommand("/RETIRAR 50000000")
          
                      
3310              Case "/RETIRAR"
3320                  If UserEstado = 1 Then 'Muerto
3330                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
3340                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
3350                      End With
3360                      Exit Sub
3370                  End If
                      
3380                  If notNullArguments Then
                          ' Version con argumentos: BankExtractGold
3390                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
3400                          Call WriteBankExtractGold(ArgumentosRaw)
3410                      Else
                              'No es numerico
3420                          Call _
                                  ShowConsoleMsg("Cantidad incorrecta. Utilice /retirar CANTIDAD.")
3430                      End If
3440                  End If

3450  Case "/DEPOSITARTODO"
3460                  Call WriteBankDepositGold(UserGLD)

3470              Case "/DEPOSITAR"
3480                  If UserEstado = 1 Then 'Muerto
3490                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
3500                          Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, _
                                  .blue, .bold, .italic)
3510                      End With
3520                      Exit Sub
3530                  End If

3540                  If notNullArguments Then
3550                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
3560                          Call WriteBankDepositGold(ArgumentosRaw)
3570                      Else
                              'No es numerico
3580                          Call _
                                  ShowConsoleMsg("Cantidad incorecta. Utilice /depositar CANTIDAD.")
3590                      End If
3600                  Else
                          'Avisar que falta el parametro
3610                      Call _
                              ShowConsoleMsg("Faltan paramtetros. Utilice /depositar CANTIDAD.")
3620                  End If

3630                  If notNullArguments Then
3640                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
3650                          Call WriteBankDepositGold(ArgumentosRaw)
3660                      Else
                              'No es numerico
3670                          Call _
                                  ShowConsoleMsg("Cantidad incorecta. Utilice /depositar CANTIDAD.")
3680                      End If
3690                  Else
                          'Avisar que falta el parametro
3700                      Call _
                              ShowConsoleMsg("Faltan paramtetros. Utilice /depositar CANTIDAD.")
3710                  End If
                      
3720              Case "/DENUNCIAR"
3730                  If notNullArguments Then
3740                      Call WriteDenounce(ArgumentosRaw)
3750                  Else
                          'Avisar que falta el parametro
3760                      Call ShowConsoleMsg("Formule su denuncia.")
3770                  End If
                      
3780                              Case "/SOLICITUD"
3790                  If notNullArguments Then
3800                      Call WriteSolicitudes(ArgumentosRaw)
3810                  Else
                          'Avisar que falta el parametro
3820                      Call ShowConsoleMsg("Formule su solicitud.")
3830                  End If
                      
                      
3840                    Case "/LEVEL"
3850    Call WriteLevel
       
3860    Case "/RESET"
3870    Call WriteReset
                      
3880              Case "/FUNDARCLAN"
3890                  If UserLvl >= 45 Then
3900                      Call WriteGuildFundate
3910                  Else
3920                      Call _
                              ShowConsoleMsg("Para fundar un clan tenés que ser nivel 45, tener 90 skills en liderazgo y haber recolectado los 3 amuletos de líder.")
3930                  End If
                  
3940              Case "/FUNDARCLANGM"
3950                  Call WriteGuildFundation(eClanType.ct_GM)

                  '
                  ' BEGIN GM COMMANDS
                  '
                  
                  Case "/BUSCADOR"
                Call WriteSearcherShow
                  
4140              Case "/CR"
4150                  If notNullArguments Then
4160                      Call WriteCuentaRegresiva(ArgumentosRaw)
4170                  Else
                          'Avisar que falta el parametro
4180                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /CUENTAREGRESIVA TIEMPO (En segundos).")
4190                  End If
                  
4200              Case "/GMSG"
4210                  If notNullArguments Then
4220                      Call WriteGMMessage(ArgumentosRaw)
4230                  Else
                          'Avisar que falta el parametro
4240                      Call ShowConsoleMsg("Escriba un mensaje.")
4250                  End If
                      
4260              Case "/SHOWNAME"
4270                  Call WriteShowName
                      
4280              Case "/ONLINEREAL"
4290                  Call WriteOnlineRoyalArmy
                      
4300              Case "/ONLINECAOS"
4310                  Call WriteOnlineChaosLegion
                      
4320              Case "/IRCERCA"
4330                  If notNullArguments Then
4340                      Call WriteGoNearby(ArgumentosRaw)
4350                  Else
                          'Avisar que falta el parametro
4360                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /ircerca NICKNAME.")
4370                  End If
                      
4380                   Case "/SEBUSCA"
4390                  If notNullArguments Then
4400                      Call WriteGoSeBuscaa(ArgumentosRaw)
4410                  Else
                          'Avisar que falta el parametro
4420                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /SeBusca NICKNAME.")
4430                  End If
                      
4440              Case "/REM"
4450                  If notNullArguments Then
4460                      Call WriteComment(ArgumentosRaw)
4470                  Else
                          'Avisar que falta el parametro
4480                      Call ShowConsoleMsg("Escriba un comentario.")
4490                  End If
                  
4500              Case "/HORA"
4510                  Call Protocol.WriteServerTime
                  
4520              Case "/DONDE"
4530                  If notNullArguments Then
4540                      Call WriteWhere(ArgumentosRaw)
4550                  Else
                          'Avisar que falta el parametro
4560                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /donde NICKNAME.")
4570                  End If
                      
4580              Case "/NENE"
4590                  If notNullArguments Then
4600                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) _
                              Then
4610                          Call WriteCreaturesInMap(ArgumentosRaw)
4620                      Else
                              'No es numerico
4630                          Call _
                                  ShowConsoleMsg("Mapa incorrecto. Utilice /nene MAPA.")
4640                      End If
4650                  Else
                          'Por default, toma el mapa en el que esta
4660                      Call WriteCreaturesInMap(UserMap)
4670                  End If
                      
4680              Case "/TELEPLOC"
4690                  Call WriteWarpMeToTarget
                      
4700              Case "/ACTIVARGLOBAL"
4710                  Call WriteGlobalStatus
                 
4720              Case "/GLOBAL"
4730                  If notNullArguments Then
4740                      Call WriteGlobalMessage(ArgumentosRaw)
4750                  Else
                          'Avisar que falta el parametro
4760                      Call ShowConsoleMsg("Escriba un mensaje.")
4770                  End If
                      
4780              Case "/TELEP"
4790                  If notNullArguments And CantidadArgumentos >= 4 Then
4800                      If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) _
                              And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) _
                              Then
4810                          Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), _
                                  ArgumentosAll(2), ArgumentosAll(3))
4820                      Else
                              'No es numerico
4830                          Call _
                                  ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
4840                      End If
4850                  ElseIf CantidadArgumentos = 3 Then
4860                      If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) _
                              Then
                              'Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
4870                          Call WriteWarpChar("YO", ArgumentosAll(0), _
                                  ArgumentosAll(1), ArgumentosAll(2))
4880                      ElseIf ValidNumber(ArgumentosAll(1), _
                              eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), _
                              eNumber_Types.ent_Byte) Then
                              'Por defecto, si no se indica el mapa, se teletransporta al mismo donde esta el usuario
4890                          Call WriteWarpChar(ArgumentosAll(0), UserMap, _
                                  ArgumentosAll(1), ArgumentosAll(2))
4900                      Else
                              'No uso ningun formato por defecto
4910                          Call _
                                  ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
4920                      End If
4930                  ElseIf CantidadArgumentos = 2 Then
4940                      If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) _
                              Then
                              ' Por defecto, se considera que se quiere unicamente cambiar las coordenadas del usuario, en el mismo mapa
4950                          Call WriteWarpChar("YO", UserMap, ArgumentosAll(0), _
                                  ArgumentosAll(1))
4960                      Else
                              'No uso ningun formato por defecto
4970                          Call _
                                  ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")
4980                      End If
4990                  Else
                          'Avisar que falta el parametro
5000                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /telep NICKNAME MAPA X Y.")
5010                  End If
                      
5020              Case "/SILENCIAR"
5030                  If notNullArguments Then
5040                      Call WriteSilence(ArgumentosRaw)
5050                  Else
                          'Avisar que falta el parametro
5060                      Call _
                              ShowConsoleMsg("Faltan parámetros. Utilice /silenciar NICKNAME.")
5070                  End If
                      
5080              Case "/SHOW"
5090                  If notNullArguments Then
5100                      Select Case UCase$(ArgumentosAll(0))
                              Case "SOS"
5110                              Call WriteSOSShowList
                                  
5120                          Case "INT"
5130                              Call WriteShowServerForm
                                  
                                  
5140                      End Select
5150                  End If
                      
5160              Case "/IRA"
5170                  If notNullArguments Then
5180                      Call WriteGoToChar(ArgumentosRaw)
5190                  Else
                          'Avisar que falta el parametro
5200                                           With _
                                                   FontTypes(FontTypeNames.FONTTYPE_INFO)
5210                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5220                       End With
5230                  End If
              
5240              Case "/INVISIBLE"
5250                  Call WriteInvisible
                      
5260              Case "/PANELGM"
5270                  Call WriteGMPanel
                      
5280              Case "/TRABAJANDO"
5290                  Call WriteWorking
                      
5300              Case "/OCULTANDO"
5310                  Call WriteHiding
                      
5320              Case "/CARCEL"
5330                  If notNullArguments Then
5340                      tmpArr = Split(ArgumentosRaw, "@")
5350                      If UBound(tmpArr) = 2 Then
5360                          If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
5370                              Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
5380                          Else
                                  'No es numerico
5390                                                    With _
                                                            FontTypes(FontTypeNames.FONTTYPE_INFO)
5400                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5410                       End With
5420                          End If
5430                      Else
                              'Faltan los parametros con el formato propio
5440                                               With _
                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
5450                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5460                       End With
5470                      End If
5480                  Else
                          'Avisar que falta el parametro
5490                                           With _
                                                   FontTypes(FontTypeNames.FONTTYPE_INFO)
5500                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5510                       End With
5520                  End If
                      
5530              Case "/RMATA"
5540                  Call WriteKillNPC
                      
5550              Case "/ADVERTENCIA"
5560                  If notNullArguments Then
5570                      tmpArr = Split(ArgumentosRaw, "@", 2)
5580                      If UBound(tmpArr) = 1 Then
5590                          Call WriteWarnUser(tmpArr(0), tmpArr(1))
5600                      Else
                              'Faltan los parametros con el formato propio
5610                                                With _
                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
5620                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5630                       End With
5640                      End If
5650                  Else
                          'Avisar que falta el parametro
5660                                            With _
                                                    FontTypes(FontTypeNames.FONTTYPE_INFO)
5670                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5680                       End With
5690                  End If
                  
5700              Case "/INFO"
5710                  If notNullArguments Then
5720                      Call WriteRequestCharInfo(ArgumentosRaw)
5730                  Else
                          'Avisar que falta el parametro
5740                                            With _
                                                    FontTypes(FontTypeNames.FONTTYPE_INFO)
5750                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5760                       End With
5770                  End If
                      
5780              Case "/STAT"
5790                  If notNullArguments Then
5800                      Call WriteRequestCharStats(ArgumentosRaw)
5810                  Else
                          'Avisar que falta el parametro
5820                       With FontTypes(FontTypeNames.FONTTYPE_INFO)
5830                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5840                       End With
5850                  End If
                      
5860              Case "/BAL"
5870                  If notNullArguments Then
5880                      Call WriteRequestCharGold(ArgumentosRaw)
5890                  Else
                          'Avisar que falta el parametro
5900                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
5910                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
5920                       End With
5930                  End If
                      
5940              Case "/INV"
5950                  If notNullArguments Then
5960                      Call WriteRequestCharInventory(ArgumentosRaw)
5970                  Else
                          'Avisar que falta el parametro
5980                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
5990                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6000                       End With
6010                  End If
                      
6020              Case "/BOV"
6030                  If notNullArguments Then
6040                      Call WriteRequestCharBank(ArgumentosRaw)
6050                  Else
                          'Avisar que falta el parametro
6060                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
6070                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6080                       End With
6090                  End If
                      
6100              Case "/SKILLS"
6110                  If notNullArguments Then
6120                      Call WriteRequestCharSkills(ArgumentosRaw)
6130                  Else
                          'Avisar que falta el parametro
6140                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
6150                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6160                       End With
6170                  End If
                      
6180              Case "/REVIVIR"
6190                  If notNullArguments Then
6200                      Call WriteReviveChar(ArgumentosRaw)
6210                  Else
                          'Avisar que falta el parametro
6220                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
6230                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6240                       End With
6250                  End If
                      
6260              Case "/ONLINEGM"
6270                  Call WriteOnlineGM
                      
6280              Case "/ONLINEMAP"
6290                  If notNullArguments Then
6300                      If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              Then
6310                          Call WriteOnlineMap(ArgumentosAll(0))
6320                      Else
6330                          Call ShowConsoleMsg("Mapa incorrecto.")
6340                      End If
6350                  Else
6360                      Call WriteOnlineMap(UserMap)
6370                  End If
                      
6380              Case "/PERDON"
6390                  If notNullArguments Then
6400                      Call WriteForgive(ArgumentosRaw)
6410                  Else
                          'Avisar que falta el parametro
6420                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
6430                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6440                       End With
6450                  End If
                      
6460              Case "/ECHAR"
6470                  If notNullArguments Then
6480                      Call WriteKick(ArgumentosRaw)
6490                  Else
                          'Avisar que falta el parametro
6500                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
6510                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6520                       End With
6530                  End If
                      
6540              Case "/EJECUTAR"
6550                  If notNullArguments Then
6560                      Call WriteExecute(ArgumentosRaw)
6570                  Else
                          'Avisar que falta el parametro
6580                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
6590                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
6600                       End With
6610                  End If
                      
6620              Case "/BAN"
6630                  If notNullArguments Then
6640                      tmpArr = Split(ArgumentosRaw, "@", 2)
6650                      If UBound(tmpArr) = 1 Then
6660                          Call WriteBanChar(tmpArr(0), tmpArr(1))
6670                      Else
                              'Faltan los parametros con el formato propio
6680                          With FontTypes(FontTypeNames.FONTTYPE_INFO)
6690                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
6700              End With
6710                      End If
6720                  Else
                          'Avisar que falta el parametro
6730                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
6740                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
6750              End With
6760                  End If
                      
6770              Case "/UNBAN"
6780                  If notNullArguments Then
6790                      Call WriteUnbanChar(ArgumentosRaw)
6800                  Else
                          'Avisar que falta el parametro
6810                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
6820                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
6830              End With
6840                  End If
                      
6850              Case "/SEGUIR"
6860                  Call WriteNPCFollow
                      
6870              Case "/SUM"
6880                  If notNullArguments Then
6890                      Call WriteSummonChar(ArgumentosRaw)
6900                  Else
                          'Avisar que falta el parametro
6910                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
6920                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
6930              End With
6940                  End If
                      
6950              Case "/CC"
6960                  Call WriteSpawnListRequest
                      
6970              Case "/RESETINV"
6980                  Call WriteResetNPCInventory
                      
6990              Case "/LIMPIAR"
7000                  Call WriteCleanWorld
                      
7010              Case "/GMROL"
7020                  If notNullArguments Then
7030                      Call WriteServerMessage(ArgumentosRaw)
7040                  Else
                          'Avisar que falta el parametro
7050                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
7060                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
7070              End With
7080                  End If
                      
                              
7090              Case "/RMSG"
7100                  If notNullArguments Then
7110                      Call WriteRolMensaje(ArgumentosRaw)
7120                  Else
7130                     With FontTypes(FontTypeNames.FONTTYPE_INFO)
7140                  Call ShowConsoleMsg("Comando desconocido.", .red, .green, .blue, _
                          .bold, .italic)
7150              End With
7160                  End If
                              
7170                         Case "/SEGUIMIENTO"
       
7180  If notNullArguments Then
7190      Call WriteSeguimiento(ArgumentosRaw)
7200  End If
                              
7210              Case "/MAPMSG"
7220                  If notNullArguments Then
7230                      Call WriteMapMessage(ArgumentosRaw)
7240                  Else
                          'Avisar que falta el parametro
7250                      Call ShowConsoleMsg("Escriba un mensaje.")
7260                  End If
                      
7270              Case "/ACEPTAR"
7280               If notNullArguments Then
7290                      Call Protocol.WriteAcceptFight(ArgumentosRaw)
7300                  Else
                          'Avisar que falta el parametro
7310                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
7320                       Call _
                               ShowConsoleMsg("Tipea /ACEPTAR seguido del nombre del personaje.", _
                               .red, .green, .blue, .bold, .italic)
7330                       End With
7340                  End If
7350              Case "/NICK2IP"
7360                  If notNullArguments Then
7370                      Call WriteNickToIP(ArgumentosRaw)
7380                  Else
                          'Avisar que falta el parametro
7390                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
7400                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
7410                       End With
7420                  End If
                      
7430              Case "/IP2NICK"
7440                  If notNullArguments Then
7450                      If validipv4str(ArgumentosRaw) Then
7460                          Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
7470                      Else
                              'No es una IP
7480                                                                    With _
                                                                            FontTypes(FontTypeNames.FONTTYPE_INFO)
7490                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
7500                       End With
7510                      End If
7520                  Else
                          'Avisar que falta el parametro
7530                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
7540                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
7550                       End With
7560                  End If
                      
7570              Case "/ONCLAN"
7580                  If notNullArguments Then
7590                      Call WriteGuildOnlineMembers(ArgumentosRaw)
7600                  Else
                          'Avisar sintaxis incorrecta
7610                      Call ShowConsoleMsg("Utilice /onclan nombre del clan.")
7620                  End If
                      
7630              Case "/CT"
7640                  If notNullArguments And CantidadArgumentos >= 3 Then
7650                      If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) _
                              Then
                              
7660                          If CantidadArgumentos = 3 Then
7670                              Call WriteTeleportCreate(ArgumentosAll(0), _
                                      ArgumentosAll(1), ArgumentosAll(2))
7680                          Else
7690                              If ValidNumber(ArgumentosAll(3), _
                                      eNumber_Types.ent_Byte) Then
7700                                  Call WriteTeleportCreate(ArgumentosAll(0), _
                                          ArgumentosAll(1), ArgumentosAll(2), _
                                          ArgumentosAll(3))
7710                              Else
                                      'No es numerico
7720                                  Call _
                                          ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO(Opcional).")
7730                              End If
7740                          End If
7750                      Else
                              'No es numerico
7760                                                                   With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
7770                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
7780                       End With
7790                      End If
7800                  Else
                          'Avisar que falta el parametro
7810                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
7820                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
7830                       End With
7840                  End If
                      
7850              Case "/DT"
7860                  Call WriteTeleportDestroy
                      
7870              Case "/LLUVIA"
7880                  Call WriteRainToggle
                      
7890              Case "/SETDESC"
7900                  Call WriteSetCharDescription(ArgumentosRaw)
                  
7910              Case "/FORCEMIDIMAP"
7920                  If notNullArguments Then
                          'elegir el mapa es opcional
7930                      If CantidadArgumentos = 1 Then
7940                          If ValidNumber(ArgumentosAll(0), _
                                  eNumber_Types.ent_Byte) Then
                                  'eviamos un mapa nulo para que tome el del usuario.
7950                              Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
7960                          Else
                                  'No es numerico
7970                              Call _
                                      ShowConsoleMsg("Midi incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
7980                          End If
7990                      Else
8000                          If ValidNumber(ArgumentosAll(0), _
                                  eNumber_Types.ent_Byte) And _
                                  ValidNumber(ArgumentosAll(1), _
                                  eNumber_Types.ent_Integer) Then
8010                              Call WriteForceMIDIToMap(ArgumentosAll(0), _
                                      ArgumentosAll(1))
8020                          Else
                                  'No es numerico
8030                              Call _
                                      ShowConsoleMsg("Valor incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
8040                          End If
8050                      End If
8060                  Else
                          'Avisar que falta el parametro
8070                      Call _
                              ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")
8080                  End If
                      
8090              Case "/FORCEWAVMAP"
8100                  If notNullArguments Then
                          'elegir la posicion es opcional
8110                      If CantidadArgumentos = 1 Then
8120                          If ValidNumber(ArgumentosAll(0), _
                                  eNumber_Types.ent_Byte) Then
                                  'eviamos una posicion nula para que tome la del usuario.
8130                              Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
8140                          Else
                                  'No es numerico
8150                              Call _
                                      ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
8160                          End If
8170                      ElseIf CantidadArgumentos = 4 Then
8180                          If ValidNumber(ArgumentosAll(0), _
                                  eNumber_Types.ent_Byte) And _
                                  ValidNumber(ArgumentosAll(1), _
                                  eNumber_Types.ent_Integer) And _
                                  ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) _
                                  And ValidNumber(ArgumentosAll(3), _
                                  eNumber_Types.ent_Byte) Then
8190                              Call WriteForceWAVEToMap(ArgumentosAll(0), _
                                      ArgumentosAll(1), ArgumentosAll(2), _
                                      ArgumentosAll(3))
8200                          Else
                                  'No es numerico
8210                              Call _
                                      ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
8220                          End If
8230                      Else
                              'Avisar que falta el parametro
8240                          Call _
                                  ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
8250                      End If
8260                  Else
                          'Avisar que falta el parametro
8270                      Call _
                              ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
8280                  End If
                      
8290              Case "/REALMSG"
8300                  If notNullArguments Then
8310                      Call WriteRoyalArmyMessage(ArgumentosRaw)
8320                  Else
                          'Avisar que falta el parametro
8330                      Call ShowConsoleMsg("Escriba un mensaje.")
8340                  End If
                       
8350              Case "/CAOSMSG"
8360                  If notNullArguments Then
8370                      Call WriteChaosLegionMessage(ArgumentosRaw)
8380                  Else
                          'Avisar que falta el parametro
8390                      Call ShowConsoleMsg("Escriba un mensaje.")
8400                  End If
                      
8410              Case "/CIUMSG"
8420                  If notNullArguments Then
8430                      Call WriteCitizenMessage(ArgumentosRaw)
8440                  Else
                          'Avisar que falta el parametro
8450                      Call ShowConsoleMsg("Escriba un mensaje.")
8460                  End If
                  
8470              Case "/CRIMSG"
8480                  If notNullArguments Then
8490                      Call WriteCriminalMessage(ArgumentosRaw)
8500                  Else
                          'Avisar que falta el parametro
8510                      Call ShowConsoleMsg("Escriba un mensaje.")
8520                  End If
                  
8530              Case "/TALKAS"
8540                  If notNullArguments Then
8550                      Call WriteTalkAsNPC(ArgumentosRaw)
8560                  Else
                          'Avisar que falta el parametro
8570                      Call ShowConsoleMsg("Escriba un mensaje.")
8580                  End If
              
8590              Case "/MASSDEST"
8600                  Call WriteDestroyAllItemsInArea
          
8610              Case "/ACEPTCONSE"
8620                  If notNullArguments Then
8630                      Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
8640                  Else
                          'Avisar que falta el parametro
8650                                                                With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
8660                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
8670                       End With
8680                  End If
                      
8690              Case "/ACEPTCONSECAOS"
8700                  If notNullArguments Then
8710                      Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
8720                  Else
                          'Avisar que falta el parametro
8730                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
8740                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
8750                       End With
8760                  End If
                      
8770              Case "/PISO"
8780                  Call WriteItemsInTheFloor
                      
8790              Case "/ESTUPIDO"
8800                  If notNullArguments Then
8810                      Call WriteMakeDumb(ArgumentosRaw)
8820                  Else
                          'Avisar que falta el parametro
8830                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
8840                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
8850                       End With
8860                  End If
                      
8870              Case "/NOESTUPIDO"
8880                  If notNullArguments Then
8890                      Call WriteMakeDumbNoMore(ArgumentosRaw)
8900                  Else
                          'Avisar que falta el parametro
8910                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
8920                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
8930                       End With
8940                  End If
                      
8950              Case "/DUMPSECURITY"
8960                  Call WriteDumpIPTables
                      
8970              Case "/KICKCONSE"
8980                  If notNullArguments Then
8990                      Call WriteCouncilKick(ArgumentosRaw)
9000                  Else
                          'Avisar que falta el parametro
9010                                                              With _
                                                                      FontTypes(FontTypeNames.FONTTYPE_INFO)
9020                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9030                       End With
9040                  End If
                      
9050                              Case "/VERHD" '//Disco.
9060                  If notNullArguments Then
9070                      Call WriteCheckHD(ArgumentosRaw)
9080                  Else
9090                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9100                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9110                       End With
9120                  End If
                     
9130              Case "/BANHD"
9140                  If notNullArguments Then
9150                      Call WriteBanHD(ArgumentosRaw)
9160                  Else
9170                                                              With _
                                                                      FontTypes(FontTypeNames.FONTTYPE_INFO)
9180                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9190                       End With
9200                  End If
                     
9210              Case "/UNBANHD"
9220                  If notNullArguments Then
9230                      Call WriteUnBanHD(ArgumentosRaw)
9240                  Else
9250                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9260                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9270                       End With
9280                  End If
                     
                  '///Disco.
                      
9290              Case "/TRIGGER"
9300                  If notNullArguments Then
9310                      If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) _
                              Then
9320                          Call WriteSetTrigger(ArgumentosRaw)
9330                      Else
                              'No es numerico
9340                          Call _
                                  ShowConsoleMsg("Numero incorrecto. Utilice /trigger NUMERO.")
9350                      End If
9360                  Else
                          'Version sin parametro
9370                      Call WriteAskTrigger
9380                  End If
                      
9390              Case "/BANIPLIST"
9400                  Call WriteBannedIPList
                      
9410              Case "/BANIPRELOAD"
9420                  Call WriteBannedIPReload
                      
9430              Case "/MIEMBROSCLAN"
9440                  If notNullArguments Then
9450                      Call WriteGuildMemberList(ArgumentosRaw)
9460                  Else
                          'Avisar que falta el parametro
9470                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9480                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9490                       End With
9500                  End If
                      
9510              Case "/BANCLAN"
9520                  If notNullArguments Then
9530                      Call WriteGuildBan(ArgumentosRaw)
9540                  Else
                          'Avisar que falta el parametro
9550                                                                               With _
                                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9560                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9570                       End With
9580                  End If
                      
9590              Case "/BANIP"
9600                  If CantidadArgumentos >= 2 Then
9610                      If validipv4str(ArgumentosAll(0)) Then
9620                          Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), _
                                  vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) _
                                  - Len(ArgumentosAll(0)) - 1))
9630                      Else
                              'No es una IP, es un nick
9640                          Call WriteBanIP(False, str2ipv4l("0.0.0.0"), _
                                  ArgumentosAll(0), Right$(ArgumentosRaw, _
                                  Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
9650                      End If
9660                  Else
                          'Avisar que falta el parametro
9670                                                              With _
                                                                      FontTypes(FontTypeNames.FONTTYPE_INFO)
9680                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9690                       End With
9700                  End If
                      
9710              Case "/UNBANIP"
9720                  If notNullArguments Then
9730                      If validipv4str(ArgumentosRaw) Then
9740                          Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
9750                      Else
                              'No es una IP
9760                                                                   With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
9770                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9780                       End With
9790                      End If
9800                  Else
                          'Avisar que falta el parametro
9810                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9820                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9830                       End With
9840                  End If
                      
9850              Case "/CI"
9860                 If notNullArguments And CantidadArgumentos >= 2 Then
9870                      If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              And ValidNumber(ArgumentosAll(1), _
                              eNumber_Types.ent_Integer) Then
9880                          WriteCreateItem ArgumentosAll(0), ArgumentosAll(1)
9890                      Else
                              'No es numerico
9900                                                                   With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
9910                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9920                       End With
9930                      End If
9940                  Else
                          'Avisar que falta el parametro
9950                                                               With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
9960                       Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
9970                       End With
9980                  End If
                      
9990              Case "/DEST"
10000                 Call WriteDestroyItems
                      
10010             Case "/NOCAOS"
10020                 If notNullArguments Then
10030                     Call WriteChaosLegionKick(ArgumentosRaw)
10040                 Else
                          'Avisar que falta el parametro
10050                                                             With _
                                                                      FontTypes(FontTypeNames.FONTTYPE_INFO)
10060                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10070                      End With
10080                 End If
          
10090             Case "/NOREAL"
10100                 If notNullArguments Then
10110                     Call WriteRoyalArmyKick(ArgumentosRaw)
10120                 Else
                          'Avisar que falta el parametro
10130                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
10140                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10150                      End With
10160                 End If
          
10170             Case "/FORCEMIDI"
10180                 If notNullArguments Then
10190                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              Then
10200                         Call WriteForceMIDIAll(ArgumentosAll(0))
10210                     Else
                              'No es numerico
10220                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
10230                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10240                      End With
10250                     End If
10260                 Else
                          'Avisar que falta el parametro
10270                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
10280                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10290                      End With
10300                 End If
          
10310             Case "/FORCEWAV"
10320                 If notNullArguments Then
10330                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              Then
10340                         Call WriteForceWAVEAll(ArgumentosAll(0))
10350                     Else
                              'No es numerico
10360                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
10370                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10380                      End With
10390                     End If
10400                 Else
                          'Avisar que falta el parametro
10410                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
10420                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10430                      End With
10440                 End If
                      
10450             Case "/BORRARPENA"
10460                 If notNullArguments Then
10470                     tmpArr = Split(ArgumentosRaw, "@", 3)
10480                     If UBound(tmpArr) = 2 Then
10490                         Call WriteRemovePunishment(tmpArr(0), tmpArr(1), _
                                  tmpArr(2))
10500                     Else
                              'Faltan los parametros con el formato propio
10510                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
10520                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10530                      End With
10540                     End If
10550                 Else
                          'Avisar que falta el parametro
10560                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
10570                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10580                      End With
10590                 End If
                      
10600             Case "/BLOQ"
10610                 Call WriteTileBlockedToggle
                      
10620             Case "/MATA"
10630                 Call WriteKillNPCNoRespawn
              
10640             Case "/MASSKILL"
10650                 Call WriteKillAllNearbyNPCs
                      
10660             Case "/LASTIP"
10670                 If notNullArguments Then
10680                     Call WriteLastIP(ArgumentosRaw)
10690                 Else
                          'Avisar que falta el parametro
10700                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
10710                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10720                      End With
10730                 End If
                      
10740             Case "/SMSG"
10750                 If notNullArguments Then
10760                     Call WriteSystemMessage(ArgumentosRaw)
10770                 Else
                          'Avisar que falta el parametro
10780                     Call ShowConsoleMsg("Escriba un mensaje.")
10790                 End If
                      
10800             Case "/ACC"
10810                 If notNullArguments Then
10820                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              Then
10830                         Call WriteCreateNPC(ArgumentosAll(0))
10840                     Else
                              'No es numerico
10850                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
10860                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10870                      End With
10880                     End If
10890                 Else
                          'Avisar que falta el parametro
10900                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
10910                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
10920                      End With
10930                 End If
                      
10940             Case "/RACC"
10950                 If notNullArguments Then
10960                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) _
                              Then
10970                         Call WriteCreateNPCWithRespawn(ArgumentosAll(0))
10980                     Else
                              'No es numerico
10990                                                                   With _
                                                                            FontTypes(FontTypeNames.FONTTYPE_INFO)
11000                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11010                      End With
11020                     End If
11030                 Else
                          'Avisar que falta el parametro
11040                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
11050                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11060                      End With
11070                 End If
              
11080             Case "/AI" ' 1 - 4
11090                 If notNullArguments And CantidadArgumentos >= 2 Then
11100                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(1), _
                              eNumber_Types.ent_Integer) Then
11110                         Call WriteImperialArmour(ArgumentosAll(0), _
                                  ArgumentosAll(1))
11120                     Else
                              'No es numerico
11130                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
11140                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11150                      End With
11160                     End If
11170                 Else
                          'Avisar que falta el parametro
11180                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
11190                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11200                      End With
11210                 End If
                      
11220             Case "/AC" ' 1 - 4
11230                 If notNullArguments And CantidadArgumentos >= 2 Then
11240                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(1), _
                              eNumber_Types.ent_Integer) Then
11250                         Call WriteChaosArmour(ArgumentosAll(0), _
                                  ArgumentosAll(1))
11260                     Else
                              'No es numerico
11270                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
11280                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11290                      End With
11300                     End If
11310                 Else
                          'Avisar que falta el parametro
11320                                                             With _
                                                                      FontTypes(FontTypeNames.FONTTYPE_INFO)
11330                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11340                      End With
11350                 End If
                      
11360             Case "/NAVE"
11370                 Call WriteNavigateToggle
              
11380             Case "/HABILITAR"
11390                 Call WriteServerOpenToUsersToggle
                  
11400             Case "/APAGAR"
11410                 Call WriteTurnOffServer
                      
11420             Case "/CONDEN"
11430                 If notNullArguments Then
11440                     Call WriteTurnCriminal(ArgumentosRaw)
11450                 Else
                          'Avisar que falta el parametro
11460                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
11470                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11480                      End With
11490                 End If
                      
11500                             Case "/PERDONARCAOS"
11510                 If notNullArguments Then
11520                     Call WriteResetFactionCaos(ArgumentosRaw)
11530                 Else
                          'Avisar que falta el parametro
11540                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
11550                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11560                      End With
11570                 End If
                      
11580             Case "/PERDONARREAL"
11590                 If notNullArguments Then
11600                     Call WriteResetFactionReal(ArgumentosRaw)
11610                 Else
                          'Avisar que falta el parametro
11620                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
11630                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11640                      End With
11650                 End If
                      
11660             Case "/RAJARCLAN"
11670                 If notNullArguments Then
11680                     Call WriteRemoveCharFromGuild(ArgumentosRaw)
11690                 Else
                          'Avisar que falta el parametro
11700                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
11710                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11720                      End With
11730                 End If
                      
11740             Case "/LASTEMAIL"
11750                 If notNullArguments Then
11760                     Call WriteRequestCharMail(ArgumentosRaw)
11770                 Else
                          'Avisar que falta el parametro
11780                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
11790                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11800                      End With
11810                 End If
                      
11820             Case "/APASS"
11830                 If notNullArguments Then
11840                     tmpArr = Split(ArgumentosRaw, "@", 2)
11850                     If UBound(tmpArr) = 1 Then
11860                         Call WriteAlterPassword(tmpArr(0), tmpArr(1))
11870                     Else
                              'Faltan los parametros con el formato propio
11880                                                                   With _
                                                                            FontTypes(FontTypeNames.FONTTYPE_INFO)
11890                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11900                      End With
11910                     End If
11920                 Else
                          'Avisar que falta el parametro
11930                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
11940                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
11950                      End With
11960                 End If
                      
11970             Case "/AEMAIL"
11980                 If notNullArguments Then
11990                     tmpArr = AEMAILSplit(ArgumentosRaw)
12000                     If LenB(tmpArr(0)) = 0 Then
                              'Faltan los parametros con el formato propio
12010                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
12020                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12030                      End With
12040                     Else
12050                         Call WriteAlterMail(tmpArr(0), tmpArr(1))
12060                     End If
12070                 Else
                          'Avisar que falta el parametro
12080                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
12090                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12100                      End With
12110                 End If
                      
12120             Case "/NOMBRE"
12130                 If Not notNullArguments Then
                          'Avisar que falta el parametro
12140                     With FontTypes(FontTypeNames.FONTTYPE_INFO)
12150                         Call ShowConsoleMsg("Utiliza /NOMBRE nuevonick", .red, _
                                  .green, .blue, .bold, .italic)
12160                      End With
                           
12170                 Else
12180                     WriteChangeNick ArgumentosRaw
12190                 End If
                      
12200             Case "/ANAME"
12210                 If notNullArguments Then
12220                     tmpArr = Split(ArgumentosRaw, "@", 2)
12230                     If UBound(tmpArr) = 1 Then
12240                         Call WriteAlterName(tmpArr(0), tmpArr(1))
12250                     Else
                              'Faltan los parametros con el formato propio
12260                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
12270                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12280                      End With
12290                     End If
12300                 Else
                          'Avisar que falta el parametro
12310                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
12320                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12330                      End With
12340                 End If
                      
12350             Case "/SLOT"
12360                 If notNullArguments Then
12370                     tmpArr = Split(ArgumentosRaw, "@", 2)
12380                     If UBound(tmpArr) = 1 Then
12390                         If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
12400                             Call WriteCheckSlot(tmpArr(0), tmpArr(1))
12410                         Else
                                  'Faltan o sobran los parametros con el formato propio
12420                             Call _
                                      ShowConsoleMsg("Formato incorrecto. Utilice /slot NICK@SLOT.")
12430                         End If
12440                     Else
                              'Faltan o sobran los parametros con el formato propio
12450                                                                   With _
                                                                            FontTypes(FontTypeNames.FONTTYPE_INFO)
12460                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12470                      End With
12480                     End If
12490                 Else
                          'Avisar que falta el parametro
12500                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
12510                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12520                      End With
12530                 End If
                      
12540             Case "/CENTINELAACTIVADO"
12550                 Call WriteToggleCentinelActivated
                      
12560             Case "/DOBACKUP"
12570                 Call WriteDoBackup
                      
12580             Case "/SHOWCMSG"
12590                 If notNullArguments Then
12600                     Call WriteShowGuildMessages(ArgumentosRaw)
12610                 Else
                          'Avisar que falta el parametro
12620                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
12630                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12640                      End With
12650                 End If
                      
12660             Case "/GUARDAMAPA"
12670                 Call WriteSaveMap
                      
12680             Case "/MODMAPINFO" ' PK, BACKUP
12690                 If CantidadArgumentos > 1 Then
12700                     Select Case UCase$(ArgumentosAll(0))
                              Case "PK" ' "/MODMAPINFO PK"
12710                             Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                              
12720                         Case "BACKUP" ' "/MODMAPINFO BACKUP"
12730                             Call WriteChangeMapInfoBackup(ArgumentosAll(1) = _
                                      "1")
                              
12740                         Case "RESTRINGIR" '/MODMAPINFO RESTRINGIR
12750                             Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                              
12760                         Case "MAGIASINEFECTO" '/MODMAPINFO MAGIASINEFECTO
12770                             Call WriteChangeMapInfoNoMagic(ArgumentosAll(1))
                              
12780                         Case "INVISINEFECTO" '/MODMAPINFO INVISINEFECTO
12790                             Call WriteChangeMapInfoNoInvi(ArgumentosAll(1))
                              
12800                         Case "RESUSINEFECTO" '/MODMAPINFO RESUSINEFECTO
12810                             Call WriteChangeMapInfoNoResu(ArgumentosAll(1))
                              
12820                         Case "TERRENO" '/MODMAPINFO TERRENO
12830                             Call WriteChangeMapInfoLand(ArgumentosAll(1))
                              
12840                         Case "ZONA" '/MODMAPINFO ZONA
12850                             Call WriteChangeMapInfoZone(ArgumentosAll(1))
                                  
12860                         Case "ROBONPC" '/MODMAPINFO ROBONPC
12870                             Call WriteChangeMapInfoStealNpc(ArgumentosAll(1) = _
                                      "1")
                                  
12880                         Case "OCULTARSINEFECTO" '/MODMAPINFO OCULTARSINEFECTO
12890                             Call WriteChangeMapInfoNoOcultar(ArgumentosAll(1) = _
                                      "1")
                                  
12900                         Case "INVOCARSINEFECTO" '/MODMAPINFO INVOCARSINEFECTO
12910                             Call WriteChangeMapInfoNoInvocar(ArgumentosAll(1) = _
                                      "1")
                                  
12920                     End Select
12930                 Else
                          'Avisar que falta el parametro
12940                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
12950                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
12960                      End With
12970                 End If
                      
12980             Case "/GRABAR"
12990                 Call WriteSaveChars
                      
13000             Case "/BORRAR"
13010                 If notNullArguments Then
13020                     Select Case UCase(ArgumentosAll(0))
                              Case "SOS" ' "/BORRAR SOS"
13030                             Call WriteCleanSOS
                                  
13040                     End Select
13050                 End If
                      
13060             Case "/NOCHE"
13070                 Call WriteNight
                      
13080             Case "/ECHARTODOSPJS"
13090                 Call WriteKickAllChars
                      
13100             Case "/RELOADNPCS"
13110                 Call WriteReloadNPCs
                      
13120             Case "/RELOADSINI"
13130                 Call WriteReloadServerIni
                      
13140             Case "/RELOADHECHIZOS"
13150                 Call WriteReloadSpells
                      
13160             Case "/RELOADOBJ"
13170                 Call WriteReloadObjects
                       
13180             Case "/REINICIAR"
13190                 Call WriteRestart
                      
13200             Case "/AUTOUPDATE"
13210                 Call WriteResetAutoUpdate
                  
13220             Case "/IMPERSONAR"
13230                 Call WriteImpersonate
                      
13240             Case "/MIMETIZAR"
13250                 Call WriteImitate
                  
                  
13260             Case "/CHATCOLOR"
13270                 If notNullArguments And CantidadArgumentos >= 3 Then
13280                     If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) _
                              And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) _
                              Then
13290                         Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), _
                                  ArgumentosAll(2))
13300                     Else
                              'No es numerico
13310                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
13320                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
13330                      End With
13340                     End If
13350                 ElseIf Not notNullArguments Then    'Go back to default!
13360                     Call WriteChatColor(0, 255, 0)
13370                 Else
                          'Avisar que falta el parametro
13380                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
13390                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
13400                      End With
13410                 End If
                  
13420             Case "/IGNORADO"
13430                 Call WriteIgnored
                  
13440             Case "/SALIRRETO"
13450                 Call WriteByeFight
                      
13460             Case "/CVC"
13470                 If notNullArguments Then
13480                     WriteAcceptFightClan ArgumentosRaw
13490                 Else
                          'Avisar que falta el parametro
13500                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
13510                         Call _
                                  ShowConsoleMsg("Escribe /CVC seguido del nombre del clan que deseas retar.", _
                                  .red, .green, .blue, .bold, .italic)
13520                      End With
13530                 End If
                      
                      
                Case "/SUBIRCANJE" ' ASI ? SE
                    Call WriteSubirCanje
                    
13540           Case "/PING"
13550                 Call WritePing

                    
                  Case "/CERRARINVASION"
                    If esGM(UserCharIndex) Then
                        WriteTerminateInvasion
                    End If
                    
13560             Case "/PANELAPUESTAS"
13570                 If esGM(UserCharIndex) Then
13580                     FrmApuestasGM.Show vbModeless, frmMain
13590                 End If
                      
13600             Case "/APUESTAS"
13610                 WriteRequestApuestas
                      'FrmApuestas.Show vbModeless, frmMain
                  
13620             Case "/INFOEVENTO"
13630                 Call WriteRequestInfoEvent
                      
13640             Case "/DISOLVERCLAN"
13650                 Call WriteDisolverGuild
                      
13660             Case "/REANUDARCLAN"
13670                 If notNullArguments Then
13680                     WriteReanudarGuild ArgumentosRaw
13690                 Else
                          'Avisar que falta el parametro
13700                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
13710                         Call _
                                  ShowConsoleMsg("Escribe /REANUDARCLAN seguido del nombre del clan que deseas reanudar.", _
                                  .red, .green, .blue, .bold, .italic)
13720                      End With
13730                 End If
                      
13740             Case "/INGRESAR"
13750                 If notNullArguments Then
13760                     Call WriteParticipeEvent(ArgumentosRaw)
13770                 Else
                          'Avisar que falta el parametro
13780                      With FontTypes(FontTypeNames.FONTTYPE_INFO)
13790                         Call _
                                  ShowConsoleMsg("Escribe /INGRESAR y separado el nombre del evento que aparece en consola. /INFOEVENTO si tenes dudas.", _
                                  .red, .green, .blue, .bold, .italic)
13800                      End With
13810                 End If
                      
13820             Case "/SALIREVENTO"
13830                 Call WriteAbandonateEvent
                      
                  
13840             Case "/EVENTOS"
13850                 If esGM(UserCharIndex) Then
13860                     Call frmPanelTorneo.Show(vbModeless, frmMain)
13870                 End If
                      
13880             Case "/QUITARPJ"
13890                 Call WriteQuitarPj
                      
13900             Case "/PODER"
13910                 Call WriteWherePower

13920             Case "/LARRY"
13930                 If notNullArguments Then
13940                     tmpArr = Split(ArgumentosRaw, "@", 2)
13950                     If UBound(tmpArr) = 1 Then
13960                         If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
13970                             Call WriteLarryMataNiños(tmpArr(0), tmpArr(1))
13980                         Else
                                  'Faltan o sobran los parametros con el formato propio
13990                             Call _
                                      ShowConsoleMsg("Formato incorrecto. Utilice /LARRY NICK@TIPO.")
14000                         End If
14010                     Else
                              'Faltan o sobran los parametros con el formato propio
14020                         With FontTypes(FontTypeNames.FONTTYPE_INFO)
14030                             Call ShowConsoleMsg("Comando desconocido.", .red, _
                                      .green, .blue, .bold, .italic)
14040                         End With
14050                     End If
14060                 Else
                          'Avisar que falta el parametro
14070                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
14080                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14090                      End With
14100                 End If
                      
14110             Case "/PUNTOS"
14120                 If notNullArguments Then
14130                     tmpArr = Split(ArgumentosRaw, "@", 2)
14140                     If UBound(tmpArr) = 1 Then
14150                         Call WriteDarPoints(tmpArr(0), tmpArr(1))
14160                     Else
                              'Faltan o sobran los parametros con el formato propio
14170                         With FontTypes(FontTypeNames.FONTTYPE_INFO)
14180                             Call ShowConsoleMsg("Comando desconocido.", .red, _
                                      .green, .blue, .bold, .italic)
14190                         End With
14200                     End If
14210                 Else
                          'Avisar que falta el parametro
14220                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
14230                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14240                      End With
14250                 End If
                      
14260             Case "/BANDIAS"
14270                 If notNullArguments Then
14280                     tmpArr = Split(ArgumentosRaw, "@", 2)
14290                     If UBound(tmpArr) = 1 Then
14300                         Call WriteComandoParaDias(tmpArr(0), tmpArr(1), 0)
14310                     Else
                              'Faltan o sobran los parametros con el formato propio
14320                         With FontTypes(FontTypeNames.FONTTYPE_INFO)
14330                             Call ShowConsoleMsg("Comando desconocido.", .red, _
                                      .green, .blue, .bold, .italic)
14340                         End With
14350                     End If
14360                 Else
                          'Avisar que falta el parametro
14370                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
14380                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14390                      End With
14400                 End If
                      
14410             Case "/DARDIOS"
14420                 If notNullArguments Then
14430                     tmpArr = Split(ArgumentosRaw, "@", 2)
14440                     If UBound(tmpArr) = 1 Then
14450                         Call WriteComandoParaDias(tmpArr(0), tmpArr(1), 1)
14460                     Else
                              'Faltan o sobran los parametros con el formato propio
14470                         With FontTypes(FontTypeNames.FONTTYPE_INFO)
14480                             Call ShowConsoleMsg("Comando desconocido.", .red, _
                                      .green, .blue, .bold, .italic)
14490                         End With
14500                     End If
14510                 Else
                          'Avisar que falta el parametro
14520                                                              With _
                                                                       FontTypes(FontTypeNames.FONTTYPE_INFO)
14530                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14540                      End With
14550                 End If
14560             Case "/ROSTRO"
14570                 Call WriteHead(-1)
                      
14580             Case "/QUEST"
14590                 Call WriteQuest
       
14600             Case "/INFOQUEST"
14610               Call WriteQuestListRequest
                      
14620             Case "/SETINIVAR"
14630                 If CantidadArgumentos = 3 Then
14640                     ArgumentosAll(2) = Replace(ArgumentosAll(2), "+", " ")
14650                     Call WriteSetIniVar(ArgumentosAll(0), ArgumentosAll(1), _
                              ArgumentosAll(2))
14660                 Else
14670                     Call _
                              ShowConsoleMsg("Prámetros incorrectos. Utilice /SETINIVAR LLAVE CLAVE VALOR")
14680                 End If
                  
14690             Case "/HOGAR"
14700             Call WriteHome
                      
14710                                 Case "/INTERCAMBIAR"
14720                 If notNullArguments Then
14730                     tmpArr = Split(ArgumentosRaw, "@", 2)
14740                     If UBound(tmpArr) = 1 Then
14750                         Call WriteCambioPj(tmpArr(0), tmpArr(1))
14760                     Else
                              'Faltan los parametros con el formato propio
14770                                                                  With _
                                                                           FontTypes(FontTypeNames.FONTTYPE_INFO)
14780                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14790                      End With
14800                     End If
14810                 Else
                          'Avisar que falta el parametro
14820                                                               With _
                                                                        FontTypes(FontTypeNames.FONTTYPE_INFO)
14830                      Call ShowConsoleMsg("Comando desconocido.", .red, .green, _
                               .blue, .bold, .italic)
14840                      End With
14850                 End If
                      
14860         End Select
              
14870     ElseIf Left$(Comando, 1) = "\" Then
14880         If UserEstado = 1 Then 'Muerto
14890             With FontTypes(FontTypeNames.FONTTYPE_INFO)
14900                 Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
14910             End With
14920             Exit Sub
14930         End If
              ' Mensaje Privado
14940         Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
              
14950     ElseIf Left$(Comando, 1) = "-" Then
14960         If UserEstado = 1 Then 'Muerto
14970             With FontTypes(FontTypeNames.FONTTYPE_INFO)
14980                 Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, _
                          .bold, .italic)
14990             End With
15000             Exit Sub
15010         End If
              ' Gritar
15020         Call WriteYell(mid$(RawCommand, 2))
              
15030     Else
              ' Hablar
15040         Call WriteTalk(RawCommand)
15050     End If
End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As _
    Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As _
    Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As _
    Boolean = False)
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 01/03/07
      '
      '***************************************************
10        Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, _
              italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal Tipo As _
    eNumber_Types) As Boolean
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 01/06/07
      '
      '***************************************************
          Dim Minimo As Long
          Dim Maximo As Long
          
10        If Not IsNumeric(Numero) Then Exit Function
          
20        Select Case Tipo
              Case eNumber_Types.ent_Byte
30                Minimo = 0
40                Maximo = 255

50            Case eNumber_Types.ent_Integer
60                Minimo = -32768
70                Maximo = 32767

80            Case eNumber_Types.ent_Long
90                Minimo = -2147483648#
100               Maximo = 2147483647
              
110           Case eNumber_Types.ent_Trigger
120               Minimo = 0
130               Maximo = 6
140       End Select
          
150       If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal Ip As String) As Boolean
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 01/06/07
      '
      '***************************************************
          Dim tmpArr() As String
          
10        tmpArr = Split(Ip, ".")
          
20        If UBound(tmpArr) <> 3 Then Exit Function

30        If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or Not _
              ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(2), _
              eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(3), _
              eNumber_Types.ent_Byte) Then Exit Function
          
40        validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal Ip As String) As Byte()
      '***************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modification: 07/26/07
      'Last Modified By: Rapsodius
      'Specify Return Type as Array of Bytes
      'Otherwise, the default is a Variant or Array of Variants, that slows down
      'the function
      '***************************************************
          Dim tmpArr() As String
          Dim bArr(3) As Byte
          
10        tmpArr = Split(Ip, ".")
          
20        bArr(0) = CByte(tmpArr(0))
30        bArr(1) = CByte(tmpArr(1))
40        bArr(2) = CByte(tmpArr(2))
50        bArr(3) = CByte(tmpArr(3))

60        str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
      '***************************************************
      'Author: Lucas Tavolaro Ortuz (Tavo)
      'Useful for AEMAIL BUG FIX
      'Last Modification: 07/26/07
      'Last Modified By: Rapsodius
      'Specify Return Type as Array of Strings
      'Otherwise, the default is a Variant or Array of Variants, that slows down
      'the function
      '***************************************************
          Dim tmpArr(0 To 1) As String
          Dim Pos As Byte
          
10        Pos = InStr(1, Text, "-")
          
20        If Pos <> 0 Then
30            tmpArr(0) = mid$(Text, 1, Pos - 1)
40            tmpArr(1) = mid$(Text, Pos + 1)
50        Else
60            tmpArr(0) = vbNullString
70        End If
          
80        AEMAILSplit = tmpArr
End Function
