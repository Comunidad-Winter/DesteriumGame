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

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

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

    VersionOK = (Ver = ULTIMAVERSION)
End Function

Sub ReSpawnOrigPosNpcs()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim i As Integer
    Dim MiNPC As npc
       
    For i = 1 To LastNPC
       'OJO
       If Npclist(i).flags.NPCActive Then
            
            If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                    MiNPC = Npclist(i)
                    Call QuitarNPC(i)
                    Call ReSpawnNpc(MiNPC)
            End If
            
            'tildada por sugerencia de yind
            'If Npclist(i).Contadores.TiempoExistencia > 0 Then
            '        Call MuereNpc(i, 0)
            'End If
       End If
       
    Next i
    
End Sub

Sub WorldSave()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim loopx As Integer
    Dim Porc As Long
    Dim hFile As Integer
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave y limpieza del mundo.", FontTypeNames.FONTTYPE_SERVER))
    
    Call LimpiarM
    
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    
    Dim j As Integer, k As Integer
    
    For j = 1 To NumMaps
        If MapInfo(j).BackUp = 1 Then k = k + 1
    Next j
    
    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = k
    FrmStat.ProgressBar1.Value = 0
    
    FrmStat.Visible = False


        For loopx = 1 To NumMaps
        'DoEvents
        If MapInfo(loopx).BackUp = 1 And MapInfo(loopx).MapVersion <> -1 Then
            ' GSZAO - Si el mapa tiene MapInfo(loopX).MapVersion = -1, es que esta fallado, por tanto "no se guarda"
            Call GrabarMapa(loopx, App.Path & "\WorldBackUp\Mapa" & loopx)
             
        End If

    Next loopx


    If FileExist(DatPath & "\bkNpcs.dat") Then Kill (DatPath & "bkNpcs.dat")
    
    hFile = FreeFile()
    
    Open DatPath & "\bkNpcs.dat" For Output As hFile
    
        For loopx = 1 To LastNPC
            If Npclist(loopx).flags.BackUp = 1 Then
                Call BackUPnPc(loopx, hFile)
            End If
        Next loopx
        
    Close hFile

    
    Call SaveForums
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> ¡No te pierdas las increíbles promociones 2x1 de este mes! Entrá a https://Ds-AO.jimdo.com/promociones-vigentes/ ", FontTypeNames.FONTTYPE_GM))
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Mundo limpiado.", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal Userindex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    UserList(Userindex).Counters.Pena = Minutos
    
    
    Call WarpUserChar(Userindex, Prision.map, Prision.X, Prision.Y, True)
    
    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(Userindex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(Userindex).flags.Traveling = 1 Then
        UserList(Userindex).flags.Traveling = 0
        UserList(Userindex).Counters.goHome = 0
        Call WriteMultiMessage(Userindex, eMessages.CancelHome)
    End If
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next
    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Function BANCheck(ByVal name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BANCheck = (val(GetVar(App.Path & "\charfile\" & name & ".chr", "FLAGS", "Ban")) = 1)

End Function
Public Function ban_Reason(ByVal lName As String) As String
 
    '
    ' @ maTih.-
 
    Dim last_P As Byte
     
    'leemos la última pena
    last_P = val(GetVar(CharPath & lName & ".chr", "PENAS", "Cant"))
     
    If (last_P <> 0) Then
        ban_Reason = GetVar(CharPath & lName & ".chr", "PENAS", "P" & CStr(last_P))
    End If
 
End Function
Public Function PersonajeExiste(ByVal name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PersonajeExiste = FileExist(CharPath & UCase$(name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & name & ".chr", "FLAGS", "Ban", "0")
    
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    If MD5ClientesActivado = 1 Then
        For i = 0 To UBound(MD5s)
            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function
            End If
        Next i
        MD5ok = False
    Else
        MD5ok = True
    End If

End Function

Public Sub MD5sCarga()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Integer
    
    MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))
    
    If MD5ClientesActivado = 1 Then
        ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
           MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC
    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    BanIps.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Dale As Boolean
    Dim LoopC As Long
    
    Dale = True
    LoopC = 1
    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop
    
    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    Dim N As Long
    
    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

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
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN
    
    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC
    
    Close #ArchN

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
    
    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"
    
    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop
    
    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN
    
    Do While Not EOF(ArchN)
        Line Input #ArchN, tmp
        BanIps.Add tmp
    Loop
    
    Close #ArchN

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
    
    Contador = Contador + 1
    
    If Contador >= 10 Then
        Contador = 0
        tmp = EstadisticasWeb.EstadisticasAndando()
        
        If Andando = False And tmp = True Then
            Call InicializaEstadisticas
        End If
        
        Andando = tmp
    End If

End Sub
Public Function RemoverRegistroHD(ByVal HD As String) As Boolean '//Disco.
   On Error Resume Next
 
    Dim N As Long
   
    N = BuscarRegistroHD(HD)
    If N > 0 Then
        BanHDs.Remove N
        RegistroBanHD
        RemoverRegistroHD = True
    Else
        RemoverRegistroHD = False
    End If
   
End Function
Public Sub AgregarRegistroHD(ByVal HD As String)
BanHDs.Add HD
 
Call RegistroBanHD
End Sub
Public Function BuscarRegistroHD(ByVal HD As String) As Long '//Disco.
    Dim Dale As Boolean
    Dim LoopC As Long
   
    Dale = True
    LoopC = 1
    Do While LoopC <= BanHDs.Count And Dale
        Dale = (BanHDs.Item(LoopC) <> HD)
        LoopC = LoopC + 1
    Loop
   
    If Dale Then
        BuscarRegistroHD = 0
    Else
        BuscarRegistroHD = LoopC - 1
    End If
End Function
Public Sub RegistroBanHD() '//Disco.
    Dim ArchivoBanHD As String
    Dim ArchN As Long
    Dim LoopC As Long
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
       
    ArchN = FreeFile()
    Open ArchivoBanHD For Output As #ArchN
   
    For LoopC = 1 To BanHDs.Count
        Print #ArchN, BanHDs.Item(LoopC)
    Next LoopC
   
    Close #ArchN
   
End Sub
Public Sub BanHDCargar() '//Disco.
    Dim ArchN As Long
    Dim tmp As String
    Dim ArchivoBanHD As String
   
    ArchivoBanHD = App.Path & "\Dat\BanHDs.dat"
   
    Do While BanHDs.Count > 0
        BanHDs.Remove 1
    Loop
   
    ArchN = FreeFile()
    Open ArchivoBanHD For Input As #ArchN
   
    Do While Not EOF(ArchN)
        Line Input #ArchN, tmp
        BanHDs.Add tmp
    Loop
   
    Close #ArchN
End Sub

Public Function UserDarPrivilegioLevel(ByVal name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************

    If EsAdmin(name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    ElseIf EsUsuario(name) Then
        UserDarPrivilegioLevel = PlayerType.User
    End If
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
   
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
   
    tUser = NameIndex(UserName)
   
    Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
   
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)
           
            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                UserPriv = UserDarPrivilegioLevel(UserName)
               
                If (UserPriv And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                       
                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
                       
                        If (UserPriv And Rank) = (.flags.Privilegios And Rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                       
                        Call LogGM(.name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            Else
           
                Call LogBan(tUser, bannerUserIndex, Reason)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha baneado a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_SERVER))
               
                'Ponemos el flag de ban a 1
                UserList(tUser).flags.Ban = 1
               
                If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
                    .flags.Ban = 1
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                    Call CloseSocket(bannerUserIndex)
                End If
               
                Call LogGM(.name, "BAN a " & UserName)
               
                'ponemos el flag de ban a 1
                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                'ponemos la pena
                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)
               
                Call CloseSocket(tUser)
            End If
        End If
    End With
End Sub
Public Function RemoverRegistroT0(ByVal CPU_ID As String) As Boolean '//Disco.
   On Error Resume Next
 
    Dim N As Long
   
    N = BuscarRegistroT0(CPU_ID)
    If N > 0 Then
        BanT0s.Remove N
        RegistroBanT0
        RemoverRegistroT0 = True
    Else
        RemoverRegistroT0 = False
    End If
   
End Function
Public Sub AgregarRegistroT0(ByVal CPU_ID As String)
BanT0s.Add CPU_ID
 
Call RegistroBanT0
End Sub

Public Function BuscarRegistroT0(ByVal CPU_ID As String) As Long '//Disco.
    Dim Dale As Boolean
    Dim LoopC As Long
   
    Dale = True
    LoopC = 1
    Do While LoopC <= BanT0s.Count And Dale
        Dale = (BanT0s.Item(LoopC) <> CPU_ID)
        LoopC = LoopC + 1
    Loop
   
    If Dale Then
        BuscarRegistroT0 = 0
    Else
        BuscarRegistroT0 = LoopC - 1
    End If
End Function
Public Sub RegistroBanT0() '//Disco.
    Dim ArchivoBanT0 As String
    Dim ArchN As Long
    Dim LoopC As Long
   
    ArchivoBanT0 = App.Path & "\Dat\BanT0s.dat"
       
    ArchN = FreeFile()
    Open ArchivoBanT0 For Output As #ArchN
   
    For LoopC = 1 To BanT0s.Count
        Print #ArchN, BanT0s.Item(LoopC)
    Next LoopC
   
    Close #ArchN
   
End Sub
Public Sub BanT0Cargar() '//Disco.
    Dim ArchN As Long
    Dim tmp As String
    Dim ArchivoBanT0 As String
   
    ArchivoBanT0 = App.Path & "\Dat\BanT0s.dat"
   
    Do While BanT0s.Count > 0
        BanT0s.Remove 1
    Loop
   
    ArchN = FreeFile()
    Open ArchivoBanT0 For Input As #ArchN
   
    Do While Not EOF(ArchN)
        Line Input #ArchN, tmp
        BanT0s.Add tmp
    Loop
   
    Close #ArchN
End Sub

