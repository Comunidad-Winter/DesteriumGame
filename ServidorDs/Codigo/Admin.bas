Attribute VB_Name = "Admin"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Type tApuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tApuestas

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public MinutosGuardarUsuarios As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        VersionOK = (Ver = ULTIMAVERSION)
End Function

Sub ReSpawnOrigPosNpcs()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


          Dim i As Integer
          Dim MiNPC As Npc
             
   On Error GoTo ReSpawnOrigPosNpcs_Error

20        For i = 1 To LastNPC
             'OJO
30           If Npclist(i).flags.NPCActive Then
                  
40                If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
50                        MiNPC = Npclist(i)
60                        Call QuitarNPC(i)
70                        Call ReSpawnNpc(MiNPC)
80                End If
                  
                  'tildada por sugerencia de yind
                  'If Npclist(i).Contadores.TiempoExistencia > 0 Then
                  '        Call MuereNpc(i, 0)
                  'End If
90           End If
             
100       Next i

   On Error GoTo 0
   Exit Sub

ReSpawnOrigPosNpcs_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ReSpawnOrigPosNpcs of Módulo Admin in line " & Erl
          
End Sub

Sub WorldSave()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


          Dim LoopX As Integer
          Dim Porc As Long
          Dim hFile As Integer
          
          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave y limpieza del mundo en 10 segundos.", FontTypeNames.FONTTYPE_SERVER))
   On Error GoTo WorldSave_Error

20        Call SendData(SendTarget.ToAll, 0, PrepareMessageShortMsj(14, FontTypeNames.FONTTYPE_SERVER))
          
          'Call LimpiarM
30        CountDownLimpieza = 10
          
40        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
          
          Dim j As Integer, k As Integer
          
50        For j = 1 To NumMaps
60            If MapInfo(j).BackUp = 1 Then k = k + 1
70        Next j
          
80        FrmStat.ProgressBar1.min = 0
90        FrmStat.ProgressBar1.max = k
100       FrmStat.ProgressBar1.value = 0
          
110       FrmStat.Visible = False
          
120       If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
          
130       hFile = FreeFile()
          
140       Open DatPath & "\bkNpcs.dat" For Output As hFile
          
150           For LoopX = 1 To LastNPC
160               If Npclist(LoopX).flags.BackUp = 1 Then
170                   Call BackUPnPc(LoopX, hFile)
180               End If
190           Next LoopX
              
200       Close hFile

          
210       Call SaveForums
          
          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> ¡No te pierdas las increíbles promociones 2x1 de este mes! Entrá a http://www.ds-ao.com.ar/donaciones", FontTypeNames.FONTTYPE_GM))
220       Call SendData(SendTarget.ToAll, 0, PrepareMessageShortMsj(15, FontTypeNames.FONTTYPE_GM))
          
          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
230       Call SendData(SendTarget.ToAll, 0, PrepareMessageShortMsj(16, FontTypeNames.FONTTYPE_SERVER))

          'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))
240       Call SendData(SendTarget.ToAll, 0, PrepareMessageShortMsj(17, FontTypeNames.FONTTYPE_SERVER))

   On Error GoTo 0
   Exit Sub

WorldSave_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WorldSave of Módulo Admin in line " & Erl

End Sub

Public Sub PurgarPenas()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Long
              
   On Error GoTo PurgarPenas_Error

10        If GambleSystem.TimeFinish > 0 Then
20            GambleSystem.TimeFinish = GambleSystem.TimeFinish - 1
              
30            If GambleSystem.TimeFinish = 0 Then
40                SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Las apuestas han sido cerradas. Ya no hay más tiempo para apostar. Mucha suerte a los usuarios que llegaron a tiempo.", FontTypeNames.FONTTYPE_GUILD)
50            End If
60        End If
          
70        For i = 1 To LastUser

80            If UserList(i).flags.UserLogged Then
90                If UserList(i).Counters.Pena > 0 Then
100                   UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                      
110                   If UserList(i).Counters.Pena < 1 Then
120                       UserList(i).Counters.Pena = 0
130                       Call WarpUserChar(i, Libertad.map, Libertad.X, Libertad.Y, True)
                          'Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
140                       Call WriteShortMsj(i, 18, FontTypeNames.FONTTYPE_INFO)
                          
150                       Call FlushBuffer(i)
160                   End If
170               End If
180           End If
190       Next i

   On Error GoTo 0
   Exit Sub

PurgarPenas_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PurgarPenas of Módulo Admin in line " & Erl
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo Encarcelar_Error

10        UserList(UserIndex).Counters.Pena = Minutos
          
20        Call WarpUserChar(UserIndex, Prision.map, Prision.X, Prision.Y, True)
          
30        If LenB(GmName) = 0 Then
              'Call WriteConsoleMsg(Userindex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
40            Call WriteShortMsj(UserIndex, 19, FontTypeNames.FONTTYPE_INFO, Minutos)

50        Else
              'Call WriteConsoleMsg(Userindex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
60            Call WriteShortMsj(UserIndex, 20, FontTypeNames.FONTTYPE_INFO, Minutos, , , , GmName)

70        End If
80        If UserList(UserIndex).flags.Traveling = 1 Then
90            UserList(UserIndex).flags.Traveling = 0
100           UserList(UserIndex).Counters.goHome = 0
110           Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
120       End If

   On Error GoTo 0
   Exit Sub

Encarcelar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Encarcelar of Módulo Admin in line " & Erl
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo BorrarUsuario_Error

20        If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
30            Kill CharPath & UCase$(UserName) & ".chr"
40        End If

   On Error GoTo 0
   Exit Sub

BorrarUsuario_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BorrarUsuario of Módulo Admin in line " & Erl
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo BANCheck_Error

10        BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

   On Error GoTo 0
   Exit Function

BANCheck_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BANCheck of Módulo Admin in line " & Erl

End Function
Public Function ban_Reason(ByVal lName As String) As String
       
          '
          ' @ maTih.-
       
          Dim last_P As Byte
           
          'leemos la última pena
   On Error GoTo ban_Reason_Error

10        last_P = val(GetVar(CharPath & lName & ".chr", "PENAS", "Cant"))
           
20        If (last_P <> 0) Then
30            ban_Reason = GetVar(CharPath & lName & ".chr", "PENAS", "P" & CStr(last_P))
40        End If

   On Error GoTo 0
   Exit Function

ban_Reason_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ban_Reason of Módulo Admin in line " & Erl
       
End Function
Public Function PersonajeExiste(ByVal Name As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo PersonajeExiste_Error

10        PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

   On Error GoTo 0
   Exit Function

PersonajeExiste_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PersonajeExiste of Módulo Admin in line " & Erl

End Function

Public Function UnBan(ByVal Name As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'Unban the character
   On Error GoTo UnBan_Error

10        Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
          
          'Remove it from the banned people database
20        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
30        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")

   On Error GoTo 0
   Exit Function

UnBan_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UnBan of Módulo Admin in line " & Erl
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Integer
          
   On Error GoTo MD5ok_Error

10        If MD5ClientesActivado = 1 Then
20            For i = 0 To UBound(MD5s)
30                If (md5formateado = MD5s(i)) Then
40                    MD5ok = True
50                    Exit Function
60                End If
70            Next i
80            MD5ok = False
90        Else
100           MD5ok = True
110       End If

   On Error GoTo 0
   Exit Function

MD5ok_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure MD5ok of Módulo Admin in line " & Erl

End Function

Public Sub MD5sCarga()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Integer
          
   On Error GoTo MD5sCarga_Error

10        MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
          
20        If MD5ClientesActivado = 1 Then
30            ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
40            For LoopC = 0 To UBound(MD5s)
50                MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
60               MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
70            Next LoopC
80        End If

   On Error GoTo 0
   Exit Sub

MD5sCarga_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure MD5sCarga of Módulo Admin in line " & Erl

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        BanIps.Add ip
          
20        Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Dale As Boolean
          Dim LoopC As Long
          
   On Error GoTo BanIpBuscar_Error

10        Dale = True
20        LoopC = 1
30        Do While LoopC <= BanIps.Count And Dale
40            Dale = (BanIps.Item(LoopC) <> ip)
50            LoopC = LoopC + 1
60        Loop
          
70        If Dale Then
80            BanIpBuscar = 0
90        Else
100           BanIpBuscar = LoopC - 1
110       End If

   On Error GoTo 0
   Exit Function

BanIpBuscar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanIpBuscar of Módulo Admin in line " & Erl
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


          Dim N As Long
          
   On Error GoTo BanIpQuita_Error

20        N = BanIpBuscar(ip)
30        If N > 0 Then
40            BanIps.Remove N
50            BanIpGuardar
60            BanIpQuita = True
70        Else
80            BanIpQuita = False
90        End If

   On Error GoTo 0
   Exit Function

BanIpQuita_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanIpQuita of Módulo Admin in line " & Erl

End Function

Public Sub BanIpGuardar()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ArchivoBanIp As String
          Dim ArchN As Long
          Dim LoopC As Long
          
   On Error GoTo BanIpGuardar_Error

10        ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
          
20        ArchN = FreeFile()
30        Open ArchivoBanIp For Output As #ArchN
          
40        For LoopC = 1 To BanIps.Count
50            Print #ArchN, BanIps.Item(LoopC)
60        Next LoopC
          
70        Close #ArchN

   On Error GoTo 0
   Exit Sub

BanIpGuardar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanIpGuardar of Módulo Admin in line " & Erl

End Sub

Public Sub BanIpCargar()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ArchN As Long
          Dim tmp As String
          Dim ArchivoBanIp As String
          
   On Error GoTo BanIpCargar_Error

10        ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
          
20        Do While BanIps.Count > 0
30            BanIps.Remove 1
40        Loop
          
50        ArchN = FreeFile()
60        Open ArchivoBanIp For Input As #ArchN
          
70        Do While Not EOF(ArchN)
80            Line Input #ArchN, tmp
90            BanIps.Add tmp
100       Loop
          
110       Close #ArchN

   On Error GoTo 0
   Exit Sub

BanIpCargar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanIpCargar of Módulo Admin in line " & Erl

End Sub

Public Sub ActualizaEstadisticasWeb()
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Static Andando As Boolean
          Static Contador As Long
          Dim tmp As Boolean
          
   On Error GoTo ActualizaEstadisticasWeb_Error

10        Contador = Contador + 1
          
20        If Contador >= 10 Then
30            Contador = 0
40            tmp = EstadisticasWeb.EstadisticasAndando()
              
50            If Andando = False And tmp = True Then
60                Call InicializaEstadisticas
70            End If
              
80            Andando = tmp
90        End If

   On Error GoTo 0
   Exit Sub

ActualizaEstadisticasWeb_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ActualizaEstadisticasWeb of Módulo Admin in line " & Erl

End Sub
Public Function RemoverRegistroHD(ByVal HD As String) As Boolean '//Disco.
       
          Dim N As Long
         
   On Error GoTo RemoverRegistroHD_Error

20        N = BuscarRegistroHD(HD)
30        If N > 0 Then
40            BanHDs.Remove N
50            RegistroBanHD
60            RemoverRegistroHD = True
70        Else
80            RemoverRegistroHD = False
90        End If

   On Error GoTo 0
   Exit Function

RemoverRegistroHD_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure RemoverRegistroHD of Módulo Admin in line " & Erl
         
End Function
Public Sub AgregarRegistroHD(ByVal HD As String)
10    BanHDs.Add HD
       
20    Call RegistroBanHD
End Sub
Public Function BuscarRegistroHD(ByVal HD As String) As Long '//Disco.
          Dim Dale As Boolean
          Dim LoopC As Long
         
   On Error GoTo BuscarRegistroHD_Error

10        Dale = True
20        LoopC = 1
30        Do While LoopC <= BanHDs.Count And Dale
40            Dale = (BanHDs.Item(LoopC) <> HD)
50            LoopC = LoopC + 1
60        Loop
         
70        If Dale Then
80            BuscarRegistroHD = 0
90        Else
100           BuscarRegistroHD = LoopC - 1
110       End If

   On Error GoTo 0
   Exit Function

BuscarRegistroHD_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BuscarRegistroHD of Módulo Admin in line " & Erl
End Function
Public Sub RegistroBanHD() '//Disco.
          Dim ArchivoBanHD As String
          Dim ArchN As Long
          Dim LoopC As Long
         
   On Error GoTo RegistroBanHD_Error

10        ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
             
20        ArchN = FreeFile()
30        Open ArchivoBanHD For Output As #ArchN
         
40        For LoopC = 1 To BanHDs.Count
50            Print #ArchN, BanHDs.Item(LoopC)
60        Next LoopC
         
70        Close #ArchN

   On Error GoTo 0
   Exit Sub

RegistroBanHD_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure RegistroBanHD of Módulo Admin in line " & Erl
         
End Sub
Public Sub BanHDCargar() '//Disco.
          Dim ArchN As Long
          Dim tmp As String
          Dim ArchivoBanHD As String
         
   On Error GoTo BanHDCargar_Error

10        ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
         
20        Do While BanHDs.Count > 0
30            BanHDs.Remove 1
40        Loop
         
50        ArchN = FreeFile()
60        Open ArchivoBanHD For Input As #ArchN
         
70        Do While Not EOF(ArchN)
80            Line Input #ArchN, tmp
90            BanHDs.Add tmp
100       Loop
         
110       Close #ArchN

   On Error GoTo 0
   Exit Sub

BanHDCargar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanHDCargar of Módulo Admin in line " & Erl
End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
      '***************************************************
      'Author: Unknown
      'Last Modification: 03/02/07
      'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
      '***************************************************

10        If EsAdmin(Name) Then
20            UserDarPrivilegioLevel = PlayerType.Admin
30        ElseIf EsDios(Name) Then
40            UserDarPrivilegioLevel = PlayerType.Dios
50        ElseIf EsSemiDios(Name) Then
60            UserDarPrivilegioLevel = PlayerType.SemiDios
70        ElseIf EsConsejero(Name) Then
80            UserDarPrivilegioLevel = PlayerType.Consejero
90        Else
100           UserDarPrivilegioLevel = PlayerType.User
110       End If
End Function


Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 03/02/07
      '22/05/2010: Ya no se peude banear admins de mayor rango si estan online.
      '***************************************************
       
          Dim tUser As Integer
          Dim UserPriv As Byte
          Dim cantPenas As Byte
          Dim Rank As Integer
         
10        If InStrB(UserName, "+") Then
20            UserName = Replace(UserName, "+", " ")
30        End If
         
40        tUser = NameIndex(UserName)
         
50        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
         
60        With UserList(bannerUserIndex)
70            If tUser <= 0 Then
                  'Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
80                Call WriteShortMsj(bannerUserIndex, 21, FontTypeNames.FONTTYPE_TALK)
                  
90                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
100                   UserPriv = UserDarPrivilegioLevel(UserName)
                     
110                   If (UserPriv And Rank) > (.flags.Privilegios And Rank) Then
                          'Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
120                       Call WriteShortMsj(bannerUserIndex, 22, FontTypeNames.FONTTYPE_TALK)
130                   Else
140                       If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
150                           Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
160                           Call WriteShortMsj(bannerUserIndex, 23, FontTypeNames.FONTTYPE_TALK)
170                       Else
180                           Call LogBanFromName(UserName, bannerUserIndex, Reason)
190                           Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                             
                              'ponemos el flag de ban a 1
200                           Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                              'ponemos la pena
210                           cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
220                           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
230                           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                             
240                           If (UserPriv And Rank) = (.flags.Privilegios And Rank) Then
250                               .flags.Ban = 1
260                               Call SendData(SendTarget.ToAdmins, 0, PrepareMessageShortMsj(25, FontTypeNames.FONTTYPE_FIGHT, , , , , .Name))
270                               Call CloseSocket(bannerUserIndex)
280                           End If
                             
290                           Call LogGM(.Name, "BAN a " & UserName)
300                       End If
310                   End If
320               Else
330                   Call WriteShortMsj(bannerUserIndex, 24, FontTypeNames.FONTTYPE_TALK)
                      'Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
340               End If
350           Else
360               If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
370                   Call WriteShortMsj(bannerUserIndex, 26, FontTypeNames.FONTTYPE_INFO)
380               Else
                 
390                   Call LogBan(tUser, bannerUserIndex, Reason)
400                   Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                     
                      'Ponemos el flag de ban a 1
410                   UserList(tUser).flags.Ban = 1
                     
420                   If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
430                       .flags.Ban = 1
440                       Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
450                       Call CloseSocket(bannerUserIndex)
460                   End If
                     
470                   Call LogGM(.Name, "BAN a " & UserName)
                     
                      'ponemos el flag de ban a 1
480                   Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                      'ponemos la pena
490                   cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
500                   Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
510                   Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                     
520                   Call CloseSocket(tUser)
530               End If
540           End If
550       End With
End Sub
