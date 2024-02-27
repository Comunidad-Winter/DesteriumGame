Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 11/03/2010
      '11/03/2010: ZaMa - Ahora no te vuelve cirminal por matar un atacable
      '***************************************************

          Dim DaExp As Integer
          Dim EraCriminal As Boolean
          
10        DaExp = CInt(UserList(VictimIndex).Stats.ELV) * 2
          
20        With UserList(AttackerIndex)
30            .Stats.Exp = .Stats.Exp + DaExp
40            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
              
50            If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
              
                  ' Es legal matarlo si estaba en atacable
60                If UserList(VictimIndex).flags.AtacablePor <> AttackerIndex Then
70                    EraCriminal = criminal(AttackerIndex)
                      
80                    With .Reputacion
90                        If Not criminal(VictimIndex) Then
100                           .AsesinoRep = .AsesinoRep + vlASESINO * 2
                              'If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
110                           .BurguesRep = 0
120                           .NobleRep = 0
130                           .PlebeRep = 0
140                       Else
150                           .NobleRep = .NobleRep + vlNoble
160                           If .NobleRep > MAXREP Then .NobleRep = MAXREP
170                       End If
180                   End With
                      
190                   If criminal(AttackerIndex) Then
200                       If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
210                   Else
220                       If EraCriminal Then Call RefreshCharStatus(AttackerIndex)
230                   End If
240               End If
250           End If
              
              
              'Lo mata
              'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
              'Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
              'Call WriteConsoleMsg(VictimIndex, "¡" & .name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
260           Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, VictimIndex, DaExp)
270           Call WriteMultiMessage(VictimIndex, eMessages.UserKill, AttackerIndex)

              'Call UserDie(VictimIndex)
280           Call FlushBuffer(VictimIndex)
              
              'Log
290           Call LogAsesinato(.Name & " asesino a " & UserList(VictimIndex).Name)
300       End With
End Sub
Public Sub RevivirUsuario(ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(userIndex)
              
20            .flags.Muerto = 0
30            .Stats.MinHp = .Stats.MaxHp
              
40            If .flags.Navegando = 1 Then
50                Call ToggleBoatBody(userIndex)
60            Else
70                Call DarCuerpoDesnudo(userIndex)
                  
80                .Char.Head = .OrigChar.Head
90            End If
              
100           If .flags.Traveling Then
110               .flags.Traveling = 0
120               .Counters.goHome = 0
130               Call WriteMultiMessage(userIndex, eMessages.CancelHome)
140           End If
              
150           Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
160           Call WriteUpdateUserStats(userIndex)
170       End With
End Sub



Public Sub ChangeUserChar(ByVal userIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal casco As Integer, Optional ByVal Transformation As Boolean = False)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If Not Transformation Then
              ' En caso de que recien transforme al usuario cambiamos de body.
20            If UserList(userIndex).flags.SlotEvent > 0 Then
30                If Events(UserList(userIndex).flags.SlotEvent).CharBody <> 0 Then
40                    Exit Sub
50                End If
60            End If
70        End If
              
80        With UserList(userIndex).Char
90            .body = body
100           .Head = Head
110           .Heading = Heading
120           .WeaponAnim = Arma
130           .ShieldAnim = Escudo
140           .CascoAnim = casco
              
150           Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCharacterChange(body, Head, Heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
160       End With
End Sub

Public Function GetWeaponAnim(ByVal userIndex As Integer, ByVal ObjIndex As Integer) As Integer
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 03/29/10
      '
      '***************************************************
          Dim tmp As Integer

10        With UserList(userIndex)
20            tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
                  
30            If tmp > 0 Then
40                If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
50                    GetWeaponAnim = tmp
60                    Exit Function
70                End If
80            End If
              
90            GetWeaponAnim = ObjData(ObjIndex).WeaponAnim
100       End With
End Function

Public Sub EnviarFama(ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim L As Long
          
10        With UserList(userIndex).Reputacion
20            L = (-.AsesinoRep) + _
                  (-.BandidoRep) + _
                  .BurguesRep + _
                  (-.LadronesRep) + _
                  .NobleRep + _
                  .PlebeRep
30            L = Round(L / 6)
              
40            .Promedio = L
50        End With
          
60        Call WriteFame(userIndex)
End Sub

Public Sub EraseUserChar(ByVal userIndex As Integer, ByVal IsAdminInvisible As Boolean)
      '*************************************************
      'Author: Unknown
      'Last modified: 08/01/2009
      '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
      '*************************************************

10    On Error GoTo ErrorHandler
          
20        With UserList(userIndex)
30            CharList(.Char.CharIndex) = 0
              
40            If .Char.CharIndex = LastChar Then
50                Do Until CharList(LastChar) > 0
60                    LastChar = LastChar - 1
70                    If LastChar <= 1 Then Exit Do
80                Loop
90            End If
              
              ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
100           If IsAdminInvisible Then
110               Call EnviarDatosASlot(userIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
120           Else
                  'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
130               Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
140           End If
              
150           Call QuitarUser(userIndex, .Pos.map)
              
160           MapData(.Pos.map, .Pos.X, .Pos.Y).userIndex = 0
170           .Char.CharIndex = 0
180       End With
          
190       NumChars = NumChars - 1
200   Exit Sub
          
ErrorHandler:
210       Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)
End Sub

Public Sub RefreshCharStatus(ByVal userIndex As Integer)
      '*************************************************
      'Author: Tararira
      'Last modified: 04/07/2009
      'Refreshes the status and tag of UserIndex.
      '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
      '*************************************************
          Dim ClanTag As String
          Dim NickColor As Byte
          
10        With UserList(userIndex)
20            If .GuildIndex > 0 Then
30                ClanTag = modGuilds.GuildName(.GuildIndex)
40                ClanTag = " <" & ClanTag & ">"
50            End If
              
60            NickColor = GetNickColor(userIndex)
              
70            If .showName Then
80                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageUpdateTagAndStatus(userIndex, NickColor, .Name & ClanTag))
90            Else
100               Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageUpdateTagAndStatus(userIndex, NickColor, vbNullString))
110           End If
              
             'Si esta navengando, se cambia la barca.
120           If .flags.Navegando Then
130               If .flags.Muerto = 1 Then
140                   .Char.body = iFragataFantasmal
150               Else
160                   Call ToggleBoatBody(userIndex)
170               End If
                  
180               Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
190           End If
200       End With
End Sub

Public Function GetNickColor(ByVal userIndex As Integer) As Byte
      '*************************************************
      'Author: ZaMa
      'Last modified: 15/01/2010
      '
      '*************************************************
          
   On Error GoTo GetNickColor_Error

10        With UserList(userIndex)
              
20            If criminal(userIndex) Then
30                GetNickColor = eNickColor.ieCriminal
40            Else
50                GetNickColor = eNickColor.ieCiudadano
60            End If
              
70            If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
              
80            If .flags.SlotReto > 0 Then
90                If UserList(userIndex).flags.FightTeam = 1 Then
100                   GetNickColor = eNickColor.ieTeamUno
110               ElseIf UserList(userIndex).flags.FightTeam = 2 Then
120                   GetNickColor = eNickColor.ieTeamDos
130               End If
140           End If
              
150           If .flags.SlotEvent > 0 Then
160               With Events(.flags.SlotEvent)
170                   If .Modality = CastleMode Then
180                       If .Users(UserList(userIndex).flags.SlotUserEvent).Team = 1 Then
190                           GetNickColor = eNickColor.ieTeamUno
200                       ElseIf .Users(UserList(userIndex).flags.SlotUserEvent).Team = 2 Then
210                           GetNickColor = eNickColor.ieTeamDos
220                       End If
230                   End If
                      
240                   If .Modality = Enfrentamientos Then
250                       If UserList(userIndex).flags.FightTeam = 1 Then
260                           GetNickColor = eNickColor.ieTeamUno
270                       ElseIf UserList(userIndex).flags.FightTeam = 2 Then
280                           GetNickColor = eNickColor.ieTeamDos
290                       End If
300                   End If
                      
310               End With
320           End If
330       End With

   On Error GoTo 0
   Exit Function

GetNickColor_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure GetNickColor of Módulo UsUaRiOs in line " & Erl
          
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal userIndex As Integer, _
        ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ButIndex As Boolean = False)
      '*************************************************
      'Author: Unknown
      'Last modified: 15/01/2010
      '23/07/2009: Budi - Ahora se envía el nick
      '15/01/2010: ZaMa - Ahora se envia el color del nick.
      '*************************************************

10    On Error GoTo Errhandler

          Dim CharIndex As Integer
          Dim ClanTag As String
          Dim NickColor As Byte
          Dim UserName As String
          Dim Privileges As Byte
          
20        With UserList(userIndex)
          
30            If InMapBounds(map, X, Y) Then
                  'If needed make a new character in list
40                If .Char.CharIndex = 0 Then
50                    CharIndex = NextOpenCharIndex
60                    .Char.CharIndex = CharIndex
70                    CharList(CharIndex) = userIndex
80                End If
                  
                  'Place character on map if needed
90                If toMap Then MapData(map, X, Y).userIndex = userIndex
                  
                  'Send make character command to clients
100               If Not toMap Then
110                   If .GuildIndex > 0 Then
120                       ClanTag = modGuilds.GuildName(.GuildIndex)
130                   End If
                      
140                   NickColor = GetNickColor(userIndex)
150                   Privileges = .flags.Privilegios
                      
                      'Preparo el nick
160                   If .showName Then
170                       UserName = .Name
                          
180                       If .flags.EnConsulta Then
190                           UserName = UserName & " " & TAG_CONSULT_MODE
200                       Else
210                           If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
220                               If LenB(ClanTag) <> 0 Then _
                                      UserName = UserName & " <" & ClanTag & ">"
230                           Else
240                               If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
250                                   UserName = UserName & " " & TAG_USER_INVISIBLE
260                               Else
270                                   If LenB(ClanTag) <> 0 Then _
                                          UserName = UserName & " <" & ClanTag & ">"
280                               End If
290                           End If
300                       End If
310                   End If
                      
                      
320                   Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.Heading, _
                              .Char.CharIndex, X, Y, _
                              .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, _
                              UserName, NickColor, Privileges)

330               Else
                      'Hide the name and clan - set privs as normal user
340                    Call AgregarUser(userIndex, .Pos.map, ButIndex)
350               End If
360           End If
370       End With
380   Exit Sub

Errhandler:
390       LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
          'Resume Next
400       Call CloseSocket(userIndex)
End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal userIndex As Integer)
10        On Error GoTo CheckUserLevel_Error
      '*************************************************
      'Author: Unknown
      'Last modified: 11/19/2009
      'Chequea que el usuario no halla alcanzado el siguiente nivel,
      'de lo contrario le da la vida, mana, etc, correspodiente.
      '07/08/2006 Integer - Modificacion de los valores
      '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV

      'kUserLevel - Error : 70 - Description : Permission denied
      'Tenes que abrir el Server ne el vps como admin...
      '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
      '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
      '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
      '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
      '12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
      '02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
      '11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
      '02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
      '*************************************************
          Dim Pts As Integer
          Dim AumentoHIT As Integer
          Dim AumentoMANA As Integer
          Dim AumentoSTA As Integer
          Dim AumentoHP As Integer
          Dim WasNewbie As Boolean
             Dim WasQuince As Boolean
             Dim WasSiete As Boolean
             Dim WasOcho As Boolean
             Dim WasNueve As Boolean
             Dim WasQuinceElv As Boolean
             Dim WasVeinte As Boolean
             Dim WasVeinticinco As Boolean
             Dim WasQuinceM As Boolean
             Dim WasTreintaM As Boolean
             Dim WasHM As Boolean
             Dim WasUM As Boolean
             Dim WasMM As Boolean
             Dim WasVip As Boolean
             Dim WasVipp As Boolean
             Dim WasVipb As Boolean
             Dim WasNoUM As Boolean
      Dim waspremium As Boolean
          Dim Promedio As Double
          Dim aux As Integer
          Dim DistVida(1 To 5) As Integer
          Dim GI As Integer 'Guild Index
          
20        WasNewbie = EsNewbie(userIndex)
30        WasQuince = EsQuince(userIndex)
40        WasSiete = EsSiete(userIndex)
50        WasOcho = EsOcho(userIndex)
60        WasNueve = EsNueve(userIndex)
70        WasQuince = EsQuince(userIndex)
80        WasVeinte = EsVeinte(userIndex)
90        WasVeinticinco = EsVeinticinco(userIndex)
100       WasQuinceM = EsQuinceM(userIndex)
110       WasTreintaM = EsTreintaM(userIndex)
120       WasHM = EsHM(userIndex)
130       WasUM = EsUM(userIndex)
140       WasMM = EsMM(userIndex)
150       WasVip = EsVip(userIndex)
160       WasVipp = EsVipp(userIndex)
170       WasVipb = EsVipb(userIndex)
180       WasNoUM = NoEsUM(userIndex)
190       waspremium = EsPremium(userIndex)
          
200       With UserList(userIndex)
210           Do While .Stats.Exp >= .Stats.ELU
                  
                  'Checkea si alcanzó el máximo nivel
220               If .Stats.ELV >= STAT_MAXELV Then
230                   .Stats.Exp = 0
240                   .Stats.ELU = 0
250                   Exit Sub
260               End If
                  
                  'Store it! VER LAUTARO
                  'Call Statistics.UserLevelUp(Userindex)
                  
                  'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
                  'Call WriteConsoleMsg(Userindex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
                  
270               If .Stats.ELV = 1 Then
280                   Pts = 10
290               Else
                      'For multiple levels being rised at once
300                   Pts = Pts + 5
310               End If
                  
320               .Stats.ELV = .Stats.ELV + 1
                  
330               .Stats.Exp = .Stats.Exp - .Stats.ELU
                  
                  'Nueva subida de exp x lvl. Pablo (ToxicWaste)
340               If .Stats.ELV = 2 Then
350                   .Stats.ELU = 450
360                   ElseIf .Stats.ELV = 3 Then
370                   .Stats.ELU = 675
380                   ElseIf .Stats.ELV = 4 Then
390                   .Stats.ELU = 1012
400                   ElseIf .Stats.ELV = 5 Then
410                   .Stats.ELU = 1518
420                   ElseIf .Stats.ELV = 6 Then
430                   .Stats.ELU = 2277
440                   ElseIf .Stats.ELV = 7 Then
450                   .Stats.ELU = 3416
460                   ElseIf .Stats.ELV = 8 Then
470                   .Stats.ELU = 5124
480                   ElseIf .Stats.ELV = 9 Then
490                   .Stats.ELU = 7886
500                   ElseIf .Stats.ELV = 10 Then
510                   .Stats.ELU = 11529
520                   ElseIf .Stats.ELV = 11 Then
530                   .Stats.ELU = 14988
540                   ElseIf .Stats.ELV = 12 Then
550                   .Stats.ELU = 19484
560                   ElseIf .Stats.ELV = 13 Then
570                   .Stats.ELU = 25329
580                   ElseIf .Stats.ELV = 14 Then
590                   .Stats.ELU = 32928
600                   ElseIf .Stats.ELV = 15 Then
610                   .Stats.ELU = 42806
620                   ElseIf .Stats.ELV = 16 Then
630                   .Stats.ELU = 55648
640                   ElseIf .Stats.ELV = 17 Then
650                   .Stats.ELU = 72342
660                   ElseIf .Stats.ELV = 18 Then
670                   .Stats.ELU = 94045
680                   ElseIf .Stats.ELV = 19 Then
690                   .Stats.ELU = 122259
700                   ElseIf .Stats.ELV = 20 Then
710                   .Stats.ELU = 158937
720                   ElseIf .Stats.ELV = 21 Then
730                   .Stats.ELU = 206618
740                   ElseIf .Stats.ELV = 22 Then
750                   .Stats.ELU = 268603
760                   ElseIf .Stats.ELV = 23 Then
770                   .Stats.ELU = 349184
780                   ElseIf .Stats.ELV = 24 Then
790                   .Stats.ELU = 453939
800                   ElseIf .Stats.ELV = 25 Then
810                   .Stats.ELU = 544727
820                   ElseIf .Stats.ELV = 26 Then
830                   .Stats.ELU = 667632
840                   ElseIf .Stats.ELV = 27 Then
850                   .Stats.ELU = 784406
860                   ElseIf .Stats.ELV = 28 Then
870                   .Stats.ELU = 941287
880                   ElseIf .Stats.ELV = 29 Then
890                   .Stats.ELU = 1129544
900                   ElseIf .Stats.ELV = 30 Then
910                   .Stats.ELU = 1355453
920                   ElseIf .Stats.ELV = 31 Then
930                   .Stats.ELU = 1626544
940                   ElseIf .Stats.ELV = 32 Then
950                   .Stats.ELU = 1951853
960                   ElseIf .Stats.ELV = 33 Then
970                   .Stats.ELU = 2342224
980                   ElseIf .Stats.ELV = 34 Then
990                   .Stats.ELU = 3372803
1000                  ElseIf .Stats.ELV = 35 Then
1010                  .Stats.ELU = 4047364
1020                  ElseIf .Stats.ELV = 36 Then
1030                  .Stats.ELU = 5828204
1040                  ElseIf .Stats.ELV = 37 Then
1050                  .Stats.ELU = 6993845
1060                  ElseIf .Stats.ELV = 38 Then
1070                  .Stats.ELU = 8392614
1080                  ElseIf .Stats.ELV = 39 Then
1090                  .Stats.ELU = 10071137
1100                  ElseIf .Stats.ELV = 40 Then
1110                  .Stats.ELU = 120853640
1120                  ElseIf .Stats.ELV = 41 Then
1130                  .Stats.ELU = 145024370
1140                  ElseIf .Stats.ELV = 42 Then
1150                  .Stats.ELU = 174029240
1160                  ElseIf .Stats.ELV = 43 Then
1170                  .Stats.ELU = 208835090
1180                  ElseIf .Stats.ELV = 44 Then
1190                  .Stats.ELU = 417670180
1200                  ElseIf .Stats.ELV = 45 Then
1210                  .Stats.ELU = 835340360
1220                  ElseIf .Stats.ELV = 46 Then
1230                  .Stats.ELU = 1670680720
1240                  Else
1250                  .Stats.ELU = 0
1260                  End If
                  
1270         Select Case .clase
                 Case eClass.Warrior
1280                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1290                      AumentoHP = RandomNumber(9, 12)
1300                      Case 20
1310                      AumentoHP = RandomNumber(8, 12)
1320                      Case 19
1330                      AumentoHP = RandomNumber(8, 11)
1340                      Case 18
1350                      AumentoHP = RandomNumber(7, 11)
1360                      Case Else
1370                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
1380                      End Select
                          
                          
1390                      If (.Stats.ELV < 48) Then
1400                      AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
1410                      Else
1420                      AumentoHIT = 1
1430                      End If
                          
1440                      AumentoSTA = AumentoSTDef
                      
1450                  Case eClass.Hunter
1460                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1470                      AumentoHP = RandomNumber(9, 11)
1480                      Case 20
1490                      AumentoHP = RandomNumber(8, 11)
1500                      Case 19
1510                      AumentoHP = RandomNumber(7, 11)
1520                      Case 18
1530                      AumentoHP = RandomNumber(6, 10)
1540                      Case Else
1550                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
1560                      End Select
                      
                      
1570                  If (.Stats.ELV < 48) Then
1580                      AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
1590                      Else
1600                      AumentoHIT = 1
1610                      End If
                          
1620                      AumentoSTA = AumentoSTDef
                      
1630                  Case eClass.Pirat
1640                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1650                      AumentoHP = RandomNumber(9, 11)
1660                      Case 20
1670                      AumentoHP = RandomNumber(8, 11)
1680                      Case 19
1690                      AumentoHP = RandomNumber(7, 11)
1700                      Case 18
1710                      AumentoHP = RandomNumber(6, 11)
1720                      Case Else
1730                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
1740                      End Select
                      
1750                  If (.Stats.ELV < 48) Then
1760                      AumentoHIT = 3
1770                      Else
1780                      AumentoHIT = 2
1790                      End If
                          
1800                      AumentoSTA = AumentoSTDef
                      
1810                  Case eClass.Paladin
1820                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
1830                      AumentoHP = RandomNumber(9, 11)
1840                      Case 20
1850                      AumentoHP = RandomNumber(8, 11)
1860                      Case 19
1870                      AumentoHP = RandomNumber(7, 11)
1880                      Case 18
1890                      AumentoHP = RandomNumber(6, 11)
1900                      Case Else
1910                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
1920                      End Select
                      
                  
1930              If (.Stats.ELV > 47) Then
1940              AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4 + AdicionalHPCazador
1950              End If
                  
1960                 If (.Stats.ELV < 48) Then
1970                      AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
1980                      Else
1990                      AumentoHIT = 1
2000                      End If
                          
2010                If (.Stats.ELV < 48) Then
2020                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
2030                Else
2040                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) \ 2
2050                End If
                    
2060                      AumentoSTA = AumentoSTDef
                      
2070                  Case eClass.Thief
2080                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2090                      AumentoHP = RandomNumber(6, 9)
2100                      Case 20
2110                      AumentoHP = RandomNumber(5, 9)
2120                      Case 19
2130                      AumentoHP = RandomNumber(4, 9)
2140                      Case 18
2150                      AumentoHP = RandomNumber(4, 8)
2160                      Case Else
2170                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2180                      End Select
                      
2190                      AumentoHIT = 2
2200                      AumentoSTA = AumentoSTLadron
                      
2210                  Case eClass.Mage
2220                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2230                      AumentoHP = RandomNumber(6, 9)
2240                      Case 20
2250                      AumentoHP = RandomNumber(5, 8)
2260                      Case 19
2270                      AumentoHP = RandomNumber(4, 8)
2280                      Case 18
2290                      AumentoHP = RandomNumber(3, 8)
2300                      Case Else
2310                      AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
2320                      End Select
2330                      If AumentoHP < 1 Then AumentoHP = 4
                          
2340                      If (.Stats.ELV > 47) Then
2350                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4 - AdicionalHPCazador
2360                      End If
                          
2370                      AumentoHIT = 1
                          'AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
2380                      AumentoSTA = AumentoSTMago
                          
2390                      If (.Stats.MaxMAN >= 2000) Then
2400                      AumentoMANA = (3 * .Stats.UserAtributos(eAtributos.Inteligencia)) / 2
2410                      Else
2420                      AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
2430                      End If
                                    
2440                  Case eClass.Worker
2450                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2460                      AumentoHP = RandomNumber(9, 12)
2470                      Case 20
2480                      AumentoHP = RandomNumber(8, 12)
2490                      Case 19
2500                      AumentoHP = RandomNumber(7, 12)
2510                      Case 18
2520                      AumentoHP = RandomNumber(6, 11)
2530                      Case Else
2540                      AumentoHP = RandomNumber(6, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
2550                      End Select
                      
2560                      AumentoHIT = 1
2570                      AumentoSTA = AumentoSTTrabajador
                      
                   
2580                  Case eClass.Cleric
2590                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2600                      AumentoHP = RandomNumber(7, 10)
2610                      Case 20
2620                      AumentoHP = RandomNumber(6, 10)
2630                      Case 19
2640                      AumentoHP = RandomNumber(6, 9)
2650                      Case 18
2660                      AumentoHP = RandomNumber(5, 9)
2670                      Case Else
2680                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2690                      End Select
                      
2700                  If (.Stats.ELV > 47) Then
2710                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
2720                      End If
                      
2730                                 If (.Stats.ELV < 48) Then
2740                      AumentoHIT = 2
2750                      Else
2760                      AumentoHIT = 1
2770                      End If
                          
2780                If (.Stats.ELV < 48) Then
2790                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
2800                Else
2810                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
2820                End If
                      
2830                      AumentoSTA = AumentoSTDef
                      
2840                  Case eClass.Druid
2850                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
2860                      AumentoHP = RandomNumber(7, 10)
2870                      Case 20
2880                      AumentoHP = RandomNumber(6, 10)
2890                      Case 19
2900                      AumentoHP = RandomNumber(6, 9)
2910                      Case 18
2920                      AumentoHP = RandomNumber(5, 9)
2930                      Case Else
2940                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
2950                      End Select
                  
2960                  If (.Stats.ELV > 47) Then
2970                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
2980                      End If
                      
2990                                 If (.Stats.ELV < 48) Then
3000                      AumentoHIT = 2
3010                      Else
3020                      AumentoHIT = 1
3030                      End If
                          
3040                If (.Stats.ELV < 48) Then
3050                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
3060                Else
3070                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
3080                End If
                    
3090                      AumentoSTA = AumentoSTDef
                      
3100                  Case eClass.Assasin
3110                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3120                      AumentoHP = RandomNumber(7, 10)
3130                      Case 20
3140                      AumentoHP = RandomNumber(6, 10)
3150                      Case 19
3160                      AumentoHP = RandomNumber(6, 9)
3170                      Case 18
3180                      AumentoHP = RandomNumber(5, 9)
3190                      Case Else
3200                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
3210                      End Select
                      
3220                                  If (.Stats.ELV > 47) Then
3230                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
3240                      End If
                      
3250                                 If (.Stats.ELV < 48) Then
3260                      AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
3270                      Else
3280                      AumentoHIT = 1
3290                      End If
                          
3300                If (.Stats.ELV < 48) Then
3310                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
3320                Else
3330                AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
3340                End If
                      
3350                      AumentoSTA = AumentoSTDef
3360                  Case eClass.Bard
3370                      Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3380                      AumentoHP = RandomNumber(7, 10)
3390                      Case 20
3400                      AumentoHP = RandomNumber(6, 10)
3410                      Case 19
3420                      AumentoHP = RandomNumber(6, 9)
3430                      Case 18
3440                      AumentoHP = RandomNumber(5, 9)
3450                      Case Else
3460                      AumentoHP = RandomNumber(4, .Stats.UserAtributos(eAtributos.Constitucion) \ 2)
3470                      End Select
                      
3480                  If (.Stats.ELV > 47) Then
3490                      AumentoHP = .Stats.UserAtributos(eAtributos.Constitucion) \ 4
3500                      End If
                      
3510                                 If (.Stats.ELV < 48) Then
3520                      AumentoHIT = 2
3530                      Else
3540                      AumentoHIT = 1
3550                      End If
                          
3560                If (.Stats.ELV < 48) Then
3570                AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
3580                Else
3590                AumentoMANA = 1 * .Stats.UserAtributos(eAtributos.Inteligencia)
3600                End If
                    
3610                      AumentoSTA = AumentoSTDef
                                      
3620                  Case Else
3630                     Select Case .Stats.UserAtributos(eAtributos.Constitucion)
                          Case 21
3640                      AumentoHP = RandomNumber(6, 8)
3650                      Case 20
3660                      AumentoHP = RandomNumber(5, 8)
3670                      Case 19
3680                      AumentoHP = RandomNumber(4, 8)
3690                      Case 18
3700                      AumentoHP = RandomNumber(3, 8)
3710                      Case Else
3720                      AumentoHP = RandomNumber(5, .Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
3730                      End Select
                      
3740                      AumentoHIT = 2
3750                      AumentoSTA = AumentoSTDef
3760              End Select
                  
                  'Actualizamos HitPoints
3770              .Stats.MaxHp = .Stats.MaxHp + AumentoHP
3780              If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                  
                  'Actualizamos Stamina
3790              .Stats.MaxSta = .Stats.MaxSta + AumentoSTA
3800              If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                  
                  'Actualizamos Mana
3810              .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA
3820              If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
                  
                  'Actualizamos Golpe Máximo
3830              .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT
3840              If .Stats.ELV < 36 Then
3850                  If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
                          .Stats.MaxHIT = STAT_MAXHIT_UNDER36
3860              Else
3870                  If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
                          .Stats.MaxHIT = STAT_MAXHIT_OVER36
3880              End If
                  
                  'Actualizamos Golpe Mínimo
3890              .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
3900              If .Stats.ELV < 36 Then
3910                  If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
                          .Stats.MinHIT = STAT_MAXHIT_UNDER36
3920              Else
3930                  If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
                          .Stats.MinHIT = STAT_MAXHIT_OVER36
3940              End If

                  Dim strTemp As String
                  
                  
                  'Notificamos al user
3950              If AumentoHP > 0 Then
3960                  strTemp = "Vida: " & AumentoHP
                      'Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
3970              End If

3980              If AumentoSTA > 0 Then
3990                  strTemp = strTemp & ", Energía: " & AumentoSTA
                      'Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoSTA & " puntos de energía.", FontTypeNames.FONTTYPE_INFO)
4000              End If
4010              If AumentoMANA > 0 Then
4020                  strTemp = strTemp & ", Maná: " & AumentoMANA
                      'Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoMANA & " puntos de maná.", FontTypeNames.FONTTYPE_INFO)
4030              End If
4040              If AumentoHIT > 0 Then
4050                  strTemp = strTemp & ", Hit aumentado: +" & AumentoHIT
                      'Call WriteConsoleMsg(Userindex, "Tu golpe máximo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                      'Call WriteConsoleMsg(Userindex, "Tu golpe mínimo aumentó en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
4060              End If
                  
                  'Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
                  
4070              .Stats.MinHp = .Stats.MaxHp

4090              If .Stats.ELV = 25 Then
4100                  GI = .GuildIndex
4110                  If GI > 0 Then
4120                      If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                              'We get here, so guild has factionary alignment, we have to expulse the user
4130                          Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
4140                          Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
4150                          Call WriteConsoleMsg(userIndex, "¡Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearás! Por esta razón, hasta tanto no te enlistes en la facción bajo la cual tu clan está alineado, estarás excluído del mismo.", FontTypeNames.FONTTYPE_GUILD)
4160                      End If
4170                  End If
4180              End If

4190          Loop
              
              'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
4200          If Not EsNewbie(userIndex) And WasNewbie Then
4210              Call QuitarNewbieObj(userIndex)
4220              If UCase$(MapInfo(.Pos.map).Restringir) = "NEWBIE" Then
4230                  Call WarpUserChar(userIndex, 1, 50, 50, True)
4240                  Call WriteConsoleMsg(userIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
4250              End If
4260          End If
              
              'Send all gained skill points at once (if any)
4270          If Pts > 0 Then
4280              Call WriteLevelUp(userIndex, Pts)
                  
4290              .Stats.SkillPts = .Stats.SkillPts + Pts
4300              strTemp = strTemp & vbCrLf & "Has ganado: " & Pts & " skillpoints."
4310          End If
              
4320          WriteConsoleMsg userIndex, strTemp, FontTypeNames.FONTTYPE_INFO
4330      End With
          
4340      Call WriteUpdateUserStats(userIndex)
    
4350      On Error GoTo 0
4360      Exit Sub

CheckUserLevel_Error:

4370      LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckUserLevel, line " & Erl & "."

End Sub
Public Sub UserLevelEditation(ByVal userIndex As Integer)
    ' Procedimiento creado para entrenamiento de personajes de nivel 1 a 15. (Editar con f1)
    
    Dim LoopC As Integer
    Dim NewHp As Integer
    Dim NewMan As Integer
    Dim NewSta As Integer
    Dim NewHit As Integer
    
    With UserList(userIndex)
        If Not .Stats.ELV = 1 Then
            WriteConsoleMsg userIndex, "Tienes que ser nivel 1 para comenzar el entrenameinto automático.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If

        
        'Nivel 2 a 15
        For LoopC = 2 To 15
            NewHp = NewHp + AumentoHP(userIndex)
            NewMan = NewMan + AumentoMANA(userIndex)
            NewSta = NewSta + AumentoSTA(userIndex)
            NewHit = NewHit + AumentoHIT(userIndex)
        Next LoopC
        
        ' Nueva vida
        .Stats.MaxHp = .Stats.MaxHp + NewHp
        If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
        
        ' Nueva energía
        .Stats.MaxSta = .Stats.MaxSta + NewSta
        If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

        ' Nueva maná
        .Stats.MaxMAN = .Stats.MaxMAN + NewMan
        If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
        
        ' Nuevo golpe máximo y mínimo
        .Stats.MaxHIT = .Stats.MaxHIT + NewHit
        .Stats.MinHIT = .Stats.MinHIT + NewHit

        If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then .Stats.MaxHIT = STAT_MAXHIT_UNDER36
        If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then .Stats.MinHIT = STAT_MAXHIT_UNDER36

        
        
        .Stats.MinMAN = .Stats.MaxMAN
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinSta = .Stats.MaxSta
        .Stats.ELV = 15
        .Stats.ELU = 42806
        .Stats.SkillPts = 85
        
        WriteLevelUp userIndex, 85
        
        
        WriteConsoleMsg userIndex, "Tu personaje se incrementó a nivel 15. Podras realizar un RESET permanente del mismo con la tecla F2", FontTypeNames.FONTTYPE_INFO
        WriteUpdateUserStats userIndex
    
    End With
End Sub
Private Function AumentoHP(ByVal userIndex As Integer) As Integer
    
    Dim UserConstitucion As Byte
    
    ' AumentoHP por nivel
    With UserList(userIndex)
            UserConstitucion = .Stats.UserAtributos(eAtributos.Constitucion)
            
            Select Case .clase
                Case eClass.Warrior
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(9, 12)
                        Case 20: AumentoHP = RandomNumber(8, 12)
                        Case 19: AumentoHP = RandomNumber(8, 11)
                        Case 18: AumentoHP = RandomNumber(7, 11)
                        Case Else: AumentoHP = RandomNumber(6, UserConstitucion \ 2) + AdicionalHPGuerrero
                    End Select
                     
                     
                    'AumentoHit = IIf(.Stats.ELV > 35, 2, 3)
                    'AumentoSta = AumentoSTDef
                Case eClass.Hunter
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(9, 11)
                        Case 20: AumentoHP = RandomNumber(8, 11)
                        Case 19: AumentoHP = RandomNumber(7, 11)
                        Case 18: AumentoHP = RandomNumber(6, 10)
                        Case Else: AumentoHP = RandomNumber(6, UserConstitucion \ 2)
                    End Select
                     
                    'AumentoHit = IIf(.Stats.ELV > 35, 2, 3)
                    'AumentoSta = AumentoSTDef
                Case eClass.Pirat
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(9, 11)
                        Case 20: AumentoHP = RandomNumber(8, 11)
                        Case 19: AumentoHP = RandomNumber(7, 11)
                        Case 18: AumentoHP = RandomNumber(6, 10)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2) + AdicionalHPGuerrero
                    End Select
                     
                    'AumentoHit = 3
                    'AumentoSta = AumentoSTDef
                    
                Case eClass.Paladin
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(9, 11)
                        Case 20: AumentoHP = RandomNumber(8, 11)
                        Case 19: AumentoHP = RandomNumber(7, 11)
                        Case 18: AumentoHP = RandomNumber(6, 11)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2) + AdicionalHPCazador
                    End Select
                     
                    'AumentoHit = IIf(.Stats.ELV > 35, 1, 3)
                    'AumentoHit = 1
                    'AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    'AumentoSta = AumentoSTDef
                    
                Case eClass.Thief
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(6, 9)
                        Case 20: AumentoHP = RandomNumber(5, 9)
                        Case 19: AumentoHP = RandomNumber(4, 9)
                        Case 18: AumentoHP = RandomNumber(4, 8)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2)
                    End Select
                     
                    'AumentoHit = 2
                    'AumentoSta = AumentoSTLadron
                         
                Case eClass.Mage
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(6, 9)
                        Case 20: AumentoHP = RandomNumber(5, 8)
                        Case 19: AumentoHP = RandomNumber(4, 8)
                        Case 18: AumentoHP = RandomNumber(3, 8)
                        Case Else: AumentoHP = RandomNumber(5, UserConstitucion \ 2) - AdicionalHPCazador
                    End Select
                    
                    If AumentoHP < 1 Then AumentoHP = 4
                    'AumentoHit = 1
                    'AumentoSta = AumentoSTMago
                    
                    'If (.Stats.MaxMAN >= 2000) Then
                        'AumentoMANA = (3 * .Stats.UserAtributos(eAtributos.Inteligencia)) / 2
                    'Else
                        'AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    'End If
                    
                Case eClass.Worker
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(9, 12)
                        Case 20: AumentoHP = RandomNumber(8, 12)
                        Case 19: AumentoHP = RandomNumber(7, 12)
                        Case 18: AumentoHP = RandomNumber(6, 11)
                        Case Else: AumentoHP = RandomNumber(6, UserConstitucion \ 2) - AdicionalHPCazador
                    End Select
                    
                    'AumentoHit = 1
                    'AumentoSta = AumentoSTTrabajador
                   
                Case eClass.Cleric
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(7, 10)
                        Case 20: AumentoHP = RandomNumber(6, 10)
                        Case 19: AumentoHP = RandomNumber(6, 9)
                        Case 18: AumentoHP = RandomNumber(5, 9)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2)
                    End Select
                     
                    'AumentoHit = 2
                    'AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    'AumentoSta = AumentoSTDef
                    
                Case eClass.Druid
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(7, 10)
                        Case 20: AumentoHP = RandomNumber(6, 10)
                        Case 19: AumentoHP = RandomNumber(6, 9)
                        Case 18: AumentoHP = RandomNumber(5, 9)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2)
                    End Select
                     
                    'AumentoHit = 2
                    'AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    'AumentoSta = AumentoSTDef
                     
                Case eClass.Assasin
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(7, 10)
                        Case 20: AumentoHP = RandomNumber(6, 10)
                        Case 19: AumentoHP = RandomNumber(6, 9)
                        Case 18: AumentoHP = RandomNumber(5, 9)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2)
                    End Select
                    
                    
                    'AumentoHit = IIf(.Stats.ELV > 35, 1, 3)
                    'AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    'AumentoSta = AumentoSTDef

                Case eClass.Bard
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(7, 10)
                        Case 20: AumentoHP = RandomNumber(6, 10)
                        Case 19: AumentoHP = RandomNumber(6, 9)
                        Case 18: AumentoHP = RandomNumber(5, 9)
                        Case Else: AumentoHP = RandomNumber(4, UserConstitucion \ 2)
                    End Select
                    
                    'AumentoHit = 2
                    'AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    'AumentoSta = AumentoSTDef
                    
                Case Else
                    Select Case UserConstitucion
                        Case 21: AumentoHP = RandomNumber(6, 8)
                        Case 20: AumentoHP = RandomNumber(5, 8)
                        Case 19: AumentoHP = RandomNumber(4, 8)
                        Case 18: AumentoHP = RandomNumber(6, 8)
                        Case Else: AumentoHP = RandomNumber(5, UserConstitucion \ 2) - AdicionalHPCazador
                    End Select
                    
                    'AumentoHit = 2
                    'AumentoSta = AumentoSTDef
                    
            End Select
    End With
End Function
Private Function AumentoSTA(ByVal userIndex As Integer) As Integer
    ' Aumento de energía
    
    With UserList(userIndex)
            
            Select Case .clase
                Case eClass.Thief
                    AumentoSTA = AumentoSTLadron
                Case eClass.Mage
                    AumentoSTA = AumentoSTMago
                Case eClass.Worker
                    AumentoSTA = AumentoSTTrabajador
                Case Else
                    AumentoSTA = AumentoSTDef
            End Select
    End With
End Function
Private Function AumentoHIT(ByVal userIndex As Integer) As Integer
    ' Aumento de HIT por nivel
    With UserList(userIndex)
            
            Select Case .clase
                Case eClass.Warrior, eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    
                Case eClass.Pirat
                    AumentoHIT = 3
                    
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    
                Case eClass.Thief
                    AumentoHIT = 2
                         
                Case eClass.Mage
                    AumentoHIT = 1
                    
                Case eClass.Worker
                    AumentoHIT = 1
                   
                Case eClass.Cleric
                    AumentoHIT = 2
                    
                Case eClass.Druid
                    AumentoHIT = 2
                     
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)

                Case eClass.Bard
                    AumentoHIT = 2
                    
                Case Else
                    AumentoHIT = 2
            End Select
    End With
End Function
Private Function AumentoMANA(ByVal userIndex As Integer) As Integer
    ' Aumento de maná según clase
    
    Dim UserInteligencia As Byte

    With UserList(userIndex)
            UserInteligencia = .Stats.UserAtributos(eAtributos.Inteligencia)
            
            Select Case .clase
                    
                Case eClass.Paladin
                    AumentoMANA = UserInteligencia
                         
                Case eClass.Mage
                    If (.Stats.MaxMAN >= 2000) Then
                        AumentoMANA = (3 * UserInteligencia) / 2
                    Else
                        AumentoMANA = 3 * UserInteligencia
                    End If
                   
                Case eClass.Cleric
                    AumentoMANA = 2 * UserInteligencia
                    
                Case eClass.Druid
                    AumentoMANA = 2 * UserInteligencia
                     
                Case eClass.Assasin
                    AumentoMANA = UserInteligencia

                Case eClass.Bard
                    AumentoMANA = 2 * UserInteligencia
                    
                Case Else
                    AumentoMANA = 0
            End Select
    End With
End Function
Public Function PuedeAtravesarAgua(ByVal userIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        PuedeAtravesarAgua = UserList(userIndex).flags.Navegando = 1 _
                          Or UserList(userIndex).flags.Vuela = 1
End Function

Sub MoveUserChar(ByVal userIndex As Integer, ByVal nHeading As eHeading)
      '*************************************************
      'Author: Unknown
      'Last modified: 13/07/2009
      'Moves the char, sending the message to everyone in range.
      '30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
      '28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
      '13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
      '13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
      '*************************************************
          Dim nPos As WorldPos
          Dim sailing As Boolean
          Dim CasperIndex As Integer
          Dim CasperHeading As eHeading
          Dim isAdminInvi As Boolean
          
   On Error GoTo MoveUserChar_Error

10        sailing = PuedeAtravesarAgua(userIndex)
20        nPos = UserList(userIndex).Pos
30        Call HeadtoPos(nHeading, nPos)
              
40        isAdminInvi = (UserList(userIndex).flags.AdminInvisible = 1)
          
50        If MoveToLegalPos(UserList(userIndex).Pos.map, nPos.X, nPos.Y, sailing, Not sailing) Then
              'si no estoy solo en el mapa...
60            If MapInfo(UserList(userIndex).Pos.map).NumUsers > 1 Then
                     
70                CasperIndex = MapData(UserList(userIndex).Pos.map, nPos.X, nPos.Y).userIndex
                  'Si hay un usuario, y paso la validacion, entonces es un casper
80                If CasperIndex > 0 Then
                      ' Los admins invisibles no pueden patear caspers
90                    If Not isAdminInvi Then
                          
100                       If TriggerZonaPelea(userIndex, CasperIndex) = TRIGGER6_PROHIBE Then
110                           If UserList(CasperIndex).flags.SeguroResu = False Then
120                               UserList(CasperIndex).flags.SeguroResu = True
130                               Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)
140                           End If
150                       End If
          
160                       With UserList(CasperIndex)
170                           CasperHeading = InvertHeading(nHeading)
180                           Call HeadtoPos(CasperHeading, .Pos)
                          
                              ' Si es un admin invisible, no se avisa a los demas clientes
190                           If Not .flags.AdminInvisible = 1 Then _
                                  Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                              
200                           Call WriteForceCharMove(CasperIndex, CasperHeading)
                                  
                              'Update map and char
210                           .Char.Heading = CasperHeading
220                           MapData(.Pos.map, .Pos.X, .Pos.Y).userIndex = CasperIndex
230                       End With
                      
                          'Actualizamos las áreas de ser necesario
240                       Call ModAreas.CheckUpdateNeededUser(CasperIndex, CasperHeading)
250                   End If
260               End If
                  
                  ' Si es un admin invisible, no se avisa a los demas clientes
270               If Not isAdminInvi Then _
                      Call SendData(SendTarget.ToPCAreaButIndex, userIndex, PrepareMessageCharacterMove(UserList(userIndex).Char.CharIndex, nPos.X, nPos.Y))
                  
280           End If
              
              ' Los admins invisibles no pueden patear caspers
290           If Not (isAdminInvi And (CasperIndex <> 0)) Then
                  Dim oldUserIndex As Integer
                  
300               With UserList(userIndex)
310                   oldUserIndex = MapData(.Pos.map, .Pos.X, .Pos.Y).userIndex
                      
                      ' Si no hay intercambio de pos con nadie
320                   If oldUserIndex = userIndex Then
330                       MapData(.Pos.map, .Pos.X, .Pos.Y).userIndex = 0
340                   End If
                      
350                   .Pos = nPos
360                   .Char.Heading = nHeading
370                   MapData(.Pos.map, .Pos.X, .Pos.Y).userIndex = userIndex
                      
380                   If HayCura(userIndex) Then Call Autoresurreccion(userIndex)
                      
390                   Call DoTileEvents(userIndex, .Pos.map, .Pos.X, .Pos.Y)
400               End With
                  
                  'Actualizamos las áreas de ser necesario
410               Call ModAreas.CheckUpdateNeededUser(userIndex, nHeading)
                  
                  ' Invocaciones
420               If PuedeRealizarInvocacion(userIndex) Then
                          
                      Dim Inv As Byte
430                    Inv = Invocation.InvocacionIndex(UserList(userIndex).Pos.map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y)
                          
440                   If Inv > 0 Then
450                       If Invocaciones(Inv).Activo = 0 Then
460                           Invocation.RealizarInvocacion userIndex, Inv
470                       End If
480                   End If
490               End If
                  
500               If UserList(userIndex).flags.SlotEvent > 0 Then
510                   If Events(UserList(userIndex).flags.SlotEvent).Modality = Busqueda Then
520                       If MapData(UserList(userIndex).Pos.map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).ObjEvent = 1 Then
530                           Call EventosDS.Busqueda_GetObj(UserList(userIndex).flags.SlotEvent, UserList(userIndex).flags.SlotUserEvent)
540                           MapData(UserList(userIndex).Pos.map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).ObjEvent = 0
550                           EraseObj 10000, UserList(userIndex).Pos.map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y
                              
560                       End If
570                   End If
580               End If
590           Else
600               Call WritePosUpdate(userIndex)

610           End If
620       Else
630           Call WritePosUpdate(userIndex)
640       End If
          
650       If UserList(userIndex).Counters.Trabajando Then _
              UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando - 1

660       If UserList(userIndex).Counters.Ocultando Then _
              UserList(userIndex).Counters.Ocultando = UserList(userIndex).Counters.Ocultando - 1

          'Montando
670       If UserList(userIndex).flags.Montando Then SendData SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(RandomNumber(215, 219), UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y)

   On Error GoTo 0
   Exit Sub

MoveUserChar_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure MoveUserChar of Módulo UsUaRiOs in line " & Erl
          
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
      '*************************************************
      'Author: ZaMa
      'Last modified: 30/03/2009
      'Returns the heading opposite to the one passed by val.
      '*************************************************
10        Select Case nHeading
              Case eHeading.EAST
20                InvertHeading = WEST
30            Case eHeading.WEST
40                InvertHeading = EAST
50            Case eHeading.SOUTH
60                InvertHeading = NORTH
70            Case eHeading.NORTH
80                InvertHeading = SOUTH
90        End Select
End Function

Sub ChangeUserInv(ByVal userIndex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        UserList(userIndex).Invent.Object(Slot) = Object
20        Call WriteChangeInventorySlot(userIndex, Slot)
End Sub
Sub Autoresurreccion(ByVal userIndex As Integer)
      '******************************
      'Adaptacion a 13.0: Kaneidra
      'Last Modification: 15/05/2012
      '******************************
10        If UserList(userIndex).flags.Muerto = 1 Then
20            Call RevivirUsuario(userIndex)
30            UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MaxMAN
40            Call WriteUpdateMana(userIndex)
50            Call WriteUpdateFollow(userIndex)
60            UserList(userIndex).Stats.MinHp = UserList(userIndex).Stats.MaxHp
70            Call WriteUpdateHP(userIndex)
80            Call WriteUpdateFollow(userIndex)
90            Call WriteConsoleMsg(userIndex, "El sacerdote te ha resucitado y curado.", FontTypeNames.FONTTYPE_INFO)
100       End If
       
110       If UserList(userIndex).Stats.MinHp < UserList(userIndex).Stats.MaxHp Then
120           UserList(userIndex).Stats.MinHp = UserList(userIndex).Stats.MaxHp
130           Call WriteUpdateHP(userIndex)
140           Call WriteUpdateFollow(userIndex)
150           Call WriteConsoleMsg(userIndex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)
160       End If
       
170       If UserList(userIndex).flags.Envenenado = 1 Then UserList(userIndex).flags.Envenenado = 0
       
End Sub

Function NextOpenCharIndex() As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          
10        For LoopC = 1 To MAXCHARS
20            If CharList(LoopC) = 0 Then
30                NextOpenCharIndex = LoopC
40                NumChars = NumChars + 1
                  
50                If LoopC > LastChar Then _
                      LastChar = LoopC
                  
60                Exit Function
70            End If
80        Next LoopC
End Function

Function NextOpenUser() As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim LoopC As Long
          
10        For LoopC = 1 To MaxUsers + 1
20            If LoopC > MaxUsers Then Exit For
30            If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
40        Next LoopC
          
50        NextOpenUser = LoopC
End Function

Public Sub SendUserStatsTxt(ByVal SendIndex As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim GuildI As Integer
          
10        With UserList(userIndex)
20            Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
30            Call WriteConsoleMsg(SendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
40            Call WriteConsoleMsg(SendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
              
50            If .Invent.WeaponEqpObjIndex > 0 Then
60                Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
70            Else
80                Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
90            End If
              
100           If .Invent.ArmourEqpObjIndex > 0 Then
110               If .Invent.EscudoEqpObjIndex > 0 Then
120                   Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
130               Else
140                   Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
150               End If
160           Else
170               Call WriteConsoleMsg(SendIndex, "(CUERPO) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
180           End If
              
190           If .Invent.CascoEqpObjIndex > 0 Then
200               Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
210           Else
220               Call WriteConsoleMsg(SendIndex, "(CABEZA) Mín Def/Máx Def: 0", FontTypeNames.FONTTYPE_INFO)
230           End If
              
240           GuildI = .GuildIndex
250           If GuildI > 0 Then
260               Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
270               If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
280                   Call WriteConsoleMsg(SendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)
290               End If
                  'guildpts no tienen objeto
300           End If
              
#If ConUpTime Then
              Dim TempDate As Date
              Dim TempSecs As Long
              Dim tempStr As String
310           TempDate = Now - .LogOnTime
320           TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
330           tempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
340           Call WriteConsoleMsg(SendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
350           Call WriteConsoleMsg(SendIndex, "Total: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
              
360           Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.Gld & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
370           Call WriteConsoleMsg(SendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
380           Call WriteConsoleMsg(SendIndex, "Retos Ganados: " & .Stats.RetosGanados & "", FontTypeNames.FONTTYPE_INFO)
390           Call WriteConsoleMsg(SendIndex, "Retos Perdidos: " & .Stats.RetosPerdidos & "", FontTypeNames.FONTTYPE_INFO)
400           Call WriteConsoleMsg(SendIndex, "Oro Ganado: " & .Stats.OroGanado & "", FontTypeNames.FONTTYPE_INFO)
410           Call WriteConsoleMsg(SendIndex, "Oro Perdido: " & .Stats.OroPerdido & "", FontTypeNames.FONTTYPE_INFO)
420           Call WriteConsoleMsg(SendIndex, "Torneos Ganados: " & .Stats.TorneosGanados & "", FontTypeNames.FONTTYPE_INFO)
              
              If .Counters.TimeTelep > 0 Then
                    If .Counters.TimeTelep < 60 Then
                        Call WriteConsoleMsg(SendIndex, "Tiempo restante en el mapa: " & .Counters.TimeTelep & " segundos.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .Counters.TimeTelep = 60 Then
                        Call WriteConsoleMsg(SendIndex, "Tiempo restante en el mapa: 1 minuto.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(SendIndex, "Tiempo restante en el mapa: " & Int(.Counters.TimeTelep / 60) & " minutos.", FontTypeNames.FONTTYPE_INFO)
                    End If
              End If
430       End With
End Sub

Sub SendUserMiniStatsTxt(ByVal SendIndex As Integer, ByVal userIndex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 23/01/2007
      'Shows the users Stats when the user is online.
      '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
      '*************************************************
10        With UserList(userIndex)
20            Call WriteConsoleMsg(SendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
30            Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
40            Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
60            Call WriteConsoleMsg(SendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
              
70            If .Faccion.ArmadaReal = 1 Then
80                Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
90                Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
100               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
110           ElseIf .Faccion.FuerzasCaos = 1 Then
120               Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
130               Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
140               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
150           ElseIf .Faccion.RecibioExpInicialReal = 1 Then
160               Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
170               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
180           ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
190               Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
200               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
210           End If
              
220           Call WriteConsoleMsg(SendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
230           Call WriteConsoleMsg(SendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
              
240           If .GuildIndex > 0 Then
250               Call WriteConsoleMsg(SendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
260           End If
270       End With
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
      '*************************************************
      'Author: Unknown
      'Last modified: 23/01/2007
      'Shows the users Stats when the user is offline.
      '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
      '*************************************************
          Dim CharFile As String
          Dim Ban As String
          Dim BanDetailPath As String
          
10        BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
20        CharFile = CharPath & charName & ".chr"
          
30        If FileExist(CharFile) Then
40            Call WriteConsoleMsg(SendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
60            Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
70            Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
80            Call WriteConsoleMsg(SendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
              
              
90            If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
100               Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
110               Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
120               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
              
130           ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
140               Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
150               Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
160               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
              
170           ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
180               Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
190               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
              
200           ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
210               Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
220               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
230           End If

              
240           Call WriteConsoleMsg(SendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
250           Call WriteConsoleMsg(SendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
              
260           If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
270               Call WriteConsoleMsg(SendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
280           End If
              
290           Ban = GetVar(CharFile, "FLAGS", "Ban")
300           Call WriteConsoleMsg(SendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
              
310           If Ban = "1" Then
320               Call WriteConsoleMsg(SendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
330           End If
340       Else
350           Call WriteConsoleMsg(SendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
360       End If
End Sub

Sub SendUserInvTxt(ByVal SendIndex As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next

          Dim j As Long
          
20        With UserList(userIndex)
30            Call WriteConsoleMsg(SendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
40            Call WriteConsoleMsg(SendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
              
50            For j = 1 To .CurrentInventorySlots
60                If .Invent.Object(j).ObjIndex > 0 Then
70                    Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)
80                End If
90            Next j
100       End With
End Sub

Sub SendUserInvTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next

          Dim j As Long
          Dim CharFile As String, tmp As String
          Dim ObjInd As Long, ObjCant As Long
          
20        CharFile = CharPath & charName & ".chr"
          
30        If FileExist(CharFile, vbNormal) Then
40            Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
              
60            For j = 1 To MAX_INVENTORY_SLOTS
70                tmp = GetVar(CharFile, "Inventory", "Obj" & j)
80                ObjInd = ReadField(1, tmp, Asc("-"))
90                ObjCant = ReadField(2, tmp, Asc("-"))
100               If ObjInd > 0 Then
110                   Call WriteConsoleMsg(SendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
120               End If
130           Next j
140       Else
150           Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
160       End If
End Sub

Sub SendUserSkillsTxt(ByVal SendIndex As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next
          Dim j As Integer
          
20        Call WriteConsoleMsg(SendIndex, UserList(userIndex).Name, FontTypeNames.FONTTYPE_INFO)
          
30        For j = 1 To NUMSKILLS
40            Call WriteConsoleMsg(SendIndex, SkillsNames(j) & " = " & UserList(userIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
50        Next j
          
60        Call WriteConsoleMsg(SendIndex, "SkillLibres:" & UserList(userIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal userIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If Npclist(NpcIndex).MaestroUser > 0 Then
20            EsMascotaCiudadano = Not criminal(Npclist(NpcIndex).MaestroUser)
30            If EsMascotaCiudadano Then
40                Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(userIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
50            End If
60        End If
End Function
Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal userIndex As Integer)
      '**********************************************
      'Author: Unknown
      'Last Modification: 02/04/2010
      '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
      '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
      '06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran más al lado de él sin hacer nada.
      '02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
      '**********************************************
          Dim EraCriminal As Boolean
          
          'Guardamos el usuario que ataco el npc.
10        Npclist(NpcIndex).flags.AttackedBy = UserList(userIndex).Name
          
          'Npc que estabas atacando.
          Dim LastNpcHit As Integer
20        LastNpcHit = UserList(userIndex).flags.NPCAtacado
          'Guarda el NPC que estas atacando ahora.
30        UserList(userIndex).flags.NPCAtacado = NpcIndex
          
          'Revisamos robo de npc.
          'Guarda el primer nick que lo ataca.
40        If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
              'El que le pegabas antes ya no es tuyo
50            If LastNpcHit <> 0 Then
60                If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userIndex).Name Then
70                    Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
80                End If
90            End If
100           Npclist(NpcIndex).flags.AttackedFirstBy = UserList(userIndex).Name
110       ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(userIndex).Name Then
              'Estas robando NPC
              'El que le pegabas antes ya no es tuyo
120           If LastNpcHit <> 0 Then
130               If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userIndex).Name Then
140                   Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
150               End If
160           End If
170       End If
          
180       If Npclist(NpcIndex).MaestroUser > 0 Then
190           If Npclist(NpcIndex).MaestroUser <> userIndex Then
200               Call AllMascotasAtacanUser(userIndex, Npclist(NpcIndex).MaestroUser)
210           End If
220       End If
          
230       If EsMascotaCiudadano(NpcIndex, userIndex) Then
240           Call VolverCriminal(userIndex)
250           Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
260           Npclist(NpcIndex).Hostile = 1
270       Else
280           EraCriminal = criminal(userIndex)
              
              'Reputacion
290           If Npclist(NpcIndex).Stats.Alineacion = 0 Then
300              If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
310                   Call VolverCriminal(userIndex)
320              End If
              
330           ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
340              UserList(userIndex).Reputacion.PlebeRep = UserList(userIndex).Reputacion.PlebeRep + vlCAZADOR / 2
350              If UserList(userIndex).Reputacion.PlebeRep > MAXREP Then _
                  UserList(userIndex).Reputacion.PlebeRep = MAXREP
360           End If
              
370           If Npclist(NpcIndex).MaestroUser <> userIndex Then
                  'hacemos que el npc se defienda
380               Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
390               Npclist(NpcIndex).Hostile = 1
400           End If
              
410           If EraCriminal And Not criminal(userIndex) Then
420               Call VolverCiudadano(userIndex)
430           End If
440       End If
End Sub
Public Function PuedeApuñalar(ByVal userIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
20            If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Apuñala = 1 Then
30                PuedeApuñalar = UserList(userIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR _
                              Or UserList(userIndex).clase = eClass.Assasin
40            End If
50        End If
End Function

Public Function PuedeAcuchillar(ByVal userIndex As Integer) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 25/01/2010 (ZaMa)
      '
      '***************************************************
          
10        With UserList(userIndex)
20            If .clase = eClass.Pirat Then
30                If .Invent.WeaponEqpObjIndex > 0 Then
40                    PuedeAcuchillar = (ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1)
50                End If
60            End If
70        End With
          
End Function

Sub SubirSkill(ByVal userIndex As Integer, ByVal Skill As Integer, ByVal Acerto As Boolean)
      '*************************************************
      'Author: Unknown
      'Last modified: 11/19/2009
      '11/19/2009 Pato - Implement the new system to train the skills.
      '*************************************************
10        With UserList(userIndex)
20            If .flags.Hambre = 0 And .flags.Sed = 0 Then

30                With .Stats
40                    If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                      
                      Dim Lvl As Integer
50                    Lvl = .ELV
                      
60                    If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
                      
70                    If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
                      
80                    If Acerto Then
90                        .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_ACIERTO_SKILL
100                   Else
110                       .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_FALLO_SKILL
120                   End If
                      
130                   If .ExpSkills(Skill) >= .EluSkills(Skill) Then
140                       .UserSkills(Skill) = .UserSkills(Skill) + 1
150                       Call WriteConsoleMsg(userIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                          
160                       .Exp = .Exp + 50
170                       If .Exp > MAXEXP Then .Exp = MAXEXP
                          
180                       Call WriteConsoleMsg(userIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
                          
190                       Call WriteUpdateExp(userIndex)
200                       Call CheckUserLevel(userIndex)
210                       Call CheckEluSkill(userIndex, Skill, False)
220                   End If
230               End With
240           End If
250       End With
End Sub


''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'
Sub UserDie(ByVal userIndex As Integer, Optional ByVal AttackerIndex As Integer = 0)
        '************************************************
        'Author: Uknown
        'Last Modified: 12/01/2010 (ZaMa)
        '04/15/2008: NicoNZ - Ahora se resetea el counter del invi
        '13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
        '27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
        '21/07/2009: Marco - Al morir se desactiva el comercio seguro.
        '16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
        '27/11/2009: Budi - Al morir envia los atributos originales.
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
        '************************************************
       
        '<EhHeader>
10      On Error GoTo UserDie_Err
 
        '</EhHeader>
        Dim i  As Long
        Dim aN As Integer
 
20      With UserList(userIndex)
 
            'Sonido
30          If .Genero = eGenero.Mujer Then
40              Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MUERTE_MUJER)
50          Else
60              Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MUERTE_HOMBRE)
 
70          End If
     
            'Quitar el dialogo del user muerto
80          Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
     
90          .Stats.MinHp = 0
100         .Stats.MinSta = 0
110         .flags.AtacadoPorUser = 0
120         .flags.Envenenado = 0
130         .flags.Muerto = 1
     
140         .Counters.Trabajando = 0
     
            ' No se activa en arenas
150         If TriggerZonaPelea(userIndex, userIndex) <> TRIGGER6_PERMITE Then
160             .flags.SeguroResu = True
170             Call WriteMultiMessage(userIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
180         Else
190             .flags.SeguroResu = False
200             Call WriteMultiMessage(userIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
 
210         End If
     
220         aN = .flags.AtacadoPorNpc
 
230         If aN > 0 Then
240             Npclist(aN).Movement = Npclist(aN).flags.OldMovement
250             Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
260             Npclist(aN).flags.AttackedBy = vbNullString
 
270         End If
     
280         aN = .flags.NPCAtacado
 
290         If aN > 0 Then
300             If Npclist(aN).flags.AttackedFirstBy = .Name Then
310                 Npclist(aN).flags.AttackedFirstBy = vbNullString
 
320             End If
 
330         End If
 
340         .flags.AtacadoPorNpc = 0
350         .flags.NPCAtacado = 0
     
360         Call PerdioNpc(userIndex, False)
     
370         If .flags.Montando = 1 Then
380             .flags.Montando = 0
390             Call WriteMontateToggle(userIndex)
400         End If
     
            '<<<< Atacable >>>>
410         If .flags.AtacablePor > 0 Then
420             .flags.AtacablePor = 0
430             Call RefreshCharStatus(userIndex)
 
440         End If

            '<<<< Paralisis >>>>
450         If .flags.Paralizado = 1 Then
460             .flags.Paralizado = 0
470             Call WriteParalizeOK(userIndex)
 
480         End If
     
            '<<< Estupidez >>>
490         If .flags.Estupidez = 1 Then
500             .flags.Estupidez = 0
510             Call WriteDumbNoMore(userIndex)
 
520         End If
     
            '<<<< Descansando >>>>
530         If .flags.Descansar Then
540             .flags.Descansar = False
550             Call WriteRestOK(userIndex)
 
560         End If
     
            '<<<< Meditando >>>>
570         If .flags.Meditando Then
580             .flags.Meditando = False
590             Call WriteMeditateToggle(userIndex)
 
600         End If
     
            '<<<< Invisible >>>>
610         If .flags.invisible = 1 Or .flags.Oculto = 1 Then
620             .flags.Oculto = 0
630             .flags.invisible = 0
640             .Counters.TiempoOculto = 0
650             .Counters.Invisibilidad = 0
         
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
660             Call SetInvisible(userIndex, UserList(userIndex).Char.CharIndex, False)
 
670         End If
     
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
680         Call SetInvisible(userIndex, UserList(userIndex).Char.CharIndex, False)

690         If TriggerZonaPelea(userIndex, userIndex) <> eTrigger6.TRIGGER6_PERMITE Then
 
                ' << Si es newbie no pierde el inventario >>
700             If Not EsNewbie(userIndex) Then
710                 If MapInfo(.Pos.map).SeCaenItems = False Then
720                     Call TirarTodo(userIndex)
 
730                 End If
 
740             Else
 
750                 If EsNewbie(userIndex) And MapInfo(.Pos.map).SeCaenItems = False Then
760                     Call TirarTodosLosItemsNoNewbies(userIndex)
 
770                 End If
 
780             End If
 
790         End If


            
            
            ' Desquipar anillos
800         If .Invent.AnilloNpcObjIndex > 0 Then
810             Call Desequipar(userIndex, .Invent.AnilloNpcSlot)
                
820         End If
            
            ' DESEQUIPA TODOS LOS OBJETOS
            'desequipar armadura
830         If .Invent.ArmourEqpObjIndex > 0 Then
840             Call Desequipar(userIndex, .Invent.ArmourEqpSlot)
 
850         End If
     
            'desequipar arma
860         If .Invent.WeaponEqpObjIndex > 0 Then
870             Call Desequipar(userIndex, .Invent.WeaponEqpSlot)
 
880         End If
     
            'desequipar casco
890         If .Invent.CascoEqpObjIndex > 0 Then
900             Call Desequipar(userIndex, .Invent.CascoEqpSlot)
 
910         End If
     
            'desequipar herramienta
920         If .Invent.AnilloEqpSlot > 0 Then
930             Call Desequipar(userIndex, .Invent.AnilloEqpSlot)
 
940         End If
     
            'desequipar municiones
950         If .Invent.MunicionEqpObjIndex > 0 Then
960             Call Desequipar(userIndex, .Invent.MunicionEqpSlot)
 
970         End If
     
            'desequipar escudo
980         If .Invent.EscudoEqpObjIndex > 0 Then
990             Call Desequipar(userIndex, .Invent.EscudoEqpSlot)
 
1000        End If
     
            ' << Reseteamos los posibles FX sobre el personaje >>
1010        If .Char.loops = INFINITE_LOOPS Then
1020            .Char.FX = 0
1030            .Char.loops = 0
 
1040        End If
     
            ' << Restauramos el mimetismo
1050        If .flags.Mimetizado = 1 Then
1060            .Char.body = .CharMimetizado.body
1070            .Char.Head = .CharMimetizado.Head
1080            .Char.CascoAnim = .CharMimetizado.CascoAnim
1090            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
1100            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
1110            .Counters.Mimetismo = 0
1120            .flags.Mimetizado = 0
                ' Puede ser atacado por npcs (cuando resucite)
1130            .flags.Ignorado = False
 
1140        End If
     
            ' << Restauramos los atributos >>
1150        If .flags.TomoPocion = True Then
 
1160            For i = 1 To 5
1170                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
1180            Next i
 
1190        End If
     
            '<< Cambiamos la apariencia del char >>
1200        If .flags.Navegando = 0 Then
1210            If UserList(userIndex).Faccion.FuerzasCaos = 1 Then
1220                .Char.body = iCuerpoMuertoCrimi
1230                .Char.Head = iCabezaMuertoCrimi
1240                .Char.ShieldAnim = NingunEscudo
1250                .Char.WeaponAnim = NingunArma
1260                .Char.CascoAnim = NingunCasco
1270            Else
1280                .Char.body = iCuerpoMuerto
1290                .Char.Head = iCabezaMuerto
1300                .Char.ShieldAnim = NingunEscudo
1310                .Char.WeaponAnim = NingunArma
1320                .Char.CascoAnim = NingunCasco
 
1330            End If
 
1340        Else
1350            .Char.body = iFragataFantasmal
 
1360        End If
     
1370        For i = 1 To MAXMASCOTAS
 
1380            If .MascotasIndex(i) > 0 Then
1390                Call MuereNpc(.MascotasIndex(i), 0)
                    ' Si estan en agua o zona segura
1400            Else
1410                .MascotasType(i) = 0
 
1420            End If
 
1430        Next i
     
1440        .NroMascotas = 0
     
    
      
            '<< Actualizamos clientes >>
1450        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
1460        Call WriteUpdateUserStats(userIndex)
1470        Call WriteUpdateStrenghtAndDexterity(userIndex)
     
            '<<Cerramos comercio seguro>>
1510        Call LimpiarComercioSeguro(userIndex)

            If Not EsGm(userIndex) Then
                ' Duelos 1vs1, 2vs2, 3vs3 y ClanvsClan
1540            If .flags.SlotReto > 0 Then
1550                Call mRetos.UserdieFight(userIndex, AttackerIndex, False)
1560            End If
                    
                Dim SlotEvent As Byte
1570            SlotEvent = .flags.SlotEvent
                
1580            If SlotEvent > 0 Then
1590                Select Case Events(SlotEvent).Modality
                        Case eModalityEvent.CastleMode
1600                            .Counters.TimeCastleMode = 3
1610                    Case eModalityEvent.DeathMatch
1620                            Call DeathMatch_UserDie(SlotEvent, userIndex)
1630                    Case eModalityEvent.Aracnus, eModalityEvent.HombreLobo, eModalityEvent.Minotauro
1640                            Transformation_UserDie userIndex, AttackerIndex
1650                    Case eModalityEvent.Unstoppable
1660                            Unstoppable_Userdie SlotEvent, .flags.SlotUserEvent, UserList(AttackerIndex).flags.SlotUserEvent
1670                    Case eModalityEvent.Enfrentamientos
1680                            Fight_UserDie SlotEvent, .flags.SlotUserEvent, AttackerIndex
1690                End Select
1700            End If
            
            
1710            'If .flags.InCVC Then
1720                'UserdieCVC UserIndex
1730            'End If
            End If
1740    End With
 
        '<EhFooter>
1750    Exit Sub
 
UserDie_Err:
1760    LogError Err.Description & vbCrLf & "UserDie AttackerIndex " & AttackerIndex & "at line " & Erl
        '</EhFooter>
End Sub
Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 13/07/2010
      '13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
      '***************************************************
       
10        If EsNewbie(Muerto) Then Exit Sub
             
20        With UserList(Atacante)
30            If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
             
40            If criminal(Muerto) Then
50                If .flags.LastCrimMatado <> UserList(Muerto).Name Then
60                    .flags.LastCrimMatado = UserList(Muerto).Name
70                    If .Faccion.CriminalesMatados < MAXUSERMATADOS Then _
                          .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + 1
80                End If
                 
90                If .Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
100                   .Faccion.Reenlistadas = 200  'jaja que trucho
                     
                      'con esto evitamos que se vuelva a reenlistar
110               End If
120           Else
130               If .flags.LastCiudMatado <> UserList(Muerto).Name Then
140                   .flags.LastCiudMatado = UserList(Muerto).Name
150                   If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then _
                          .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + 1
160               End If
170           End If
             
180           If .Stats.UsuariosMatados < MAXUSERMATADOS Then _
                  .Stats.UsuariosMatados = .Stats.UsuariosMatados + 1
190       End With
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, ByRef PuedeTierra As Boolean)
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 18/09/2010
      '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
      '18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
      '**************************************************************
10    On Error GoTo Errhandler

          Dim Found As Boolean
          Dim LoopC As Integer
          Dim tX As Long
          Dim tY As Long
          
20        nPos = Pos
30        tX = Pos.X
40        tY = Pos.Y
          
50        LoopC = 1
          
          ' La primera posicion es valida?
60        If LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
              
70            If Not HayObjeto(Pos.map, nPos.X, nPos.Y, Obj.ObjIndex, Obj.Amount) Then
80                Found = True
90            End If
              
100       End If
          
          ' Busca en las demas posiciones, en forma de "rombo"
110       If Not Found Then
120           While (Not Found) And LoopC <= 16
130               If RhombLegalTilePos(Pos, tX, tY, LoopC, Obj.ObjIndex, Obj.Amount, PuedeAgua, PuedeTierra) Then
140                   nPos.X = tX
150                   nPos.Y = tY
160                   Found = True
170               End If
              
180               LoopC = LoopC + 1
190           Wend
              
200       End If
          
210       If Not Found Then
220           nPos.X = 0
230           nPos.Y = 0
240       End If
          
250       Exit Sub
          
Errhandler:
260       Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.Description)
End Sub
Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean
      '***************************************************
      'Author: Marco Vanotti (Marco)
      'Last Modification: -
      ' walks all the perimeter of a rhomb of side  "distance + 1",
      ' which starts at Pos.x - Distance and Pos.y
      '***************************************************

          Dim i As Long
          
10        vX = Pos.X - Distance
20        vY = Pos.Y
          
30        For i = 0 To Distance - 1
40            If (LegalPos(Pos.map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
50                vX = vX + i
60                vY = vY - i
70                RhombLegalPos = True
80                Exit Function
90            End If
100       Next
          
110       vX = Pos.X
120       vY = Pos.Y - Distance
          
130       For i = 0 To Distance - 1
140           If (LegalPos(Pos.map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
150               vX = vX + i
160               vY = vY + i
170               RhombLegalPos = True
180               Exit Function
190           End If
200       Next
          
210       vX = Pos.X + Distance
220       vY = Pos.Y
          
230       For i = 0 To Distance - 1
240           If (LegalPos(Pos.map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
250               vX = vX - i
260               vY = vY + i
270               RhombLegalPos = True
280               Exit Function
290           End If
300       Next
          
310       vX = Pos.X
320       vY = Pos.Y + Distance
          
330       For i = 0 To Distance - 1
340           If (LegalPos(Pos.map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
350               vX = vX - i
360               vY = vY - i
370               RhombLegalPos = True
380               Exit Function
390           End If
400       Next
          
410       RhombLegalPos = False
          
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: -
      ' walks all the perimeter of a rhomb of side  "distance + 1",
      ' which starts at Pos.x - Distance and Pos.y
      ' and searchs for a valid position to drop items
      '***************************************************
10    On Error GoTo Errhandler

          Dim i As Long
          Dim HayObj As Boolean
          
          Dim X As Integer
          Dim Y As Integer
          Dim MapObjIndex As Integer
          
20        vX = Pos.X - Distance
30        vY = Pos.Y
          
40        For i = 0 To Distance - 1
              
50            X = vX + i
60            Y = vY - i
              
70            If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
80                If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
90                    vX = X
100                   vY = Y
                      
110                   RhombLegalTilePos = True
120                   Exit Function
130               End If
                  
140           End If
150       Next
          
160       vX = Pos.X
170       vY = Pos.Y - Distance
          
180       For i = 0 To Distance - 1
              
190           X = vX + i
200           Y = vY + i
              
210           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
220               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
230                   vX = X
240                   vY = Y
                      
250                   RhombLegalTilePos = True
260                   Exit Function
270               End If
280           End If
290       Next
          
300       vX = Pos.X + Distance
310       vY = Pos.Y
          
320       For i = 0 To Distance - 1
              
330           X = vX - i
340           Y = vY + i
          
350           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
              
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
360               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
370                   vX = X
380                   vY = Y
                      
390                   RhombLegalTilePos = True
400                   Exit Function
410               End If
420           End If
430       Next
          
440       vX = Pos.X
450       vY = Pos.Y + Distance
          
460       For i = 0 To Distance - 1
              
470           X = vX - i
480           Y = vY - i
          
490           If (LegalPos(Pos.map, X, Y, PuedeAgua, PuedeTierra, True)) Then
                  ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
500               If Not HayObjeto(Pos.map, X, Y, ObjIndex, ObjAmount) Then
510                   vX = X
520                   vY = Y
                      
530                   RhombLegalTilePos = True
540                   Exit Function
550               End If
560           End If
570       Next
          
580       RhombLegalTilePos = False
          
590       Exit Function
          
Errhandler:
600       Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.Description)
End Function

Sub WarpUserChar(ByVal userIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, _
ByVal FX As Boolean, Optional ByVal Teletransported As Boolean)
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 11/23/2010
      '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
      '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
      '16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
      '11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
      '**************************************************************
          Dim OldMap As Integer
          Dim OldX As Integer
          Dim OldY As Integer
          
10        With UserList(userIndex)
              'Quitar el dialogo
20            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
              
30            OldMap = .Pos.map
40            OldX = .Pos.X
50            OldY = .Pos.Y

60            Call EraseUserChar(userIndex, .flags.AdminInvisible = 1)
              
70            If OldMap <> map Then
80                Call WriteChangeMap(userIndex, map, MapInfo(.Pos.map).MapVersion)
                  
                  
                  If .flags.Montando > 0 Then
                    If Not MapInfo(map).Pk Then
                        .Char.Head = UserList(userIndex).OrigChar.Head
                        
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            .Char.body = ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).Ropaje
                        Else
                            Call DarCuerpoDesnudo(userIndex)
                        End If
                        
                        If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).ShieldAnim
                        If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).WeaponAnim
                        If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).CascoAnim
                            
                          .flags.Montando = 0

                        Call WriteMontateToggle(userIndex)
                        WriteConsoleMsg userIndex, "El mapa no permite monturas. Has bajado de tu montura.", FontTypeNames.FONTTYPE_INFO
                    End If
                  End If
                  
                  
90                If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                      Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)
                      Dim WasInvi As Boolean
                      'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
100                   If MapInfo(map).InviSinEfecto > 0 And .flags.invisible = 1 Then
110                       .flags.invisible = 0
120                       .Counters.Invisibilidad = 0
130                       AhoraVisible = True
140                       WasInvi = True 'si era invi, para el string
150                   End If
                      'Chequeo de flags de mapa por ocultar (C4b3z0n)
160                   If MapInfo(map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
170                       AhoraVisible = True
180                       .flags.Oculto = 0
190                       .Counters.TiempoOculto = 0
200                   End If
                      
210                   If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
220                       Call SetInvisible(userIndex, .Char.CharIndex, False)
230                       If WasInvi Then 'era invi
240                           Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
250                       Else 'estaba oculto
260                           Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)
270                       End If
280                   End If
290               End If
                  
300               Call WritePlayMidi(userIndex, val(ReadField(1, MapInfo(map).Music, 45)))
                  
          
                  'Update new Map Users
310               MapInfo(map).NumUsers = MapInfo(map).NumUsers + 1
                  
                  'Update old Map Users
320               MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
330               If MapInfo(OldMap).NumUsers < 0 Then
340                   MapInfo(OldMap).NumUsers = 0
350               End If
              
                  'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
                  Dim nextMap, previousMap As Boolean
360               nextMap = IIf(distanceToCities(map).distanceToCity(.Hogar) >= 0, True, False)
370               previousMap = IIf(distanceToCities(.Pos.map).distanceToCity(.Hogar) >= 0, True, False)

380               If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                      'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
390               ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
400                   .flags.lastMap = .Pos.map
410               ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
420                   .flags.lastMap = 0
430               ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
440                   .flags.lastMap = .flags.lastMap
450               End If
                  
460               Call WriteRemoveAllDialogs(userIndex)
470           End If
              
480           .Pos.X = X
490           .Pos.Y = Y
500           .Pos.map = map
              
540           Call MakeUserChar(True, map, userIndex, map, X, Y)
550           Call WriteUserCharIndexInServer(userIndex)
              
560           Call DoTileEvents(userIndex, map, X, Y)
              
              'Force a flush, so user index is in there before it's destroyed for teleporting
570           Call FlushBuffer(userIndex)
              
              'Seguis invisible al pasar de mapa
580           If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
                  
                  ' No si estas navegando
590               If .flags.Navegando = 0 Then
600                   Call SetInvisible(userIndex, .Char.CharIndex, True)
610               End If
620           End If
              
630           If Teletransported Then
640               If .flags.Traveling = 1 Then
650                   .flags.Traveling = 0
660                   .Counters.goHome = 0
670                   Call WriteMultiMessage(userIndex, eMessages.CancelHome)
680               End If
690           End If
              
700           If FX And .flags.AdminInvisible = 0 Then 'FX
710               Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
720               Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
730           End If
              
740           If .NroMascotas Then Call WarpMascotas(userIndex)

              ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
750           Call IntervaloPermiteSerAtacado(userIndex, True)
              
              ' Perdes el npc al cambiar de mapa
760           Call PerdioNpc(userIndex, False)
              
              ' Automatic toogle navigate
770           If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
780               If HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then
790                   If .flags.Navegando = 0 Then
800                       .flags.Navegando = 1
                              
                          'Tell the client that we are navigating.
810                       Call WriteNavigateToggle(userIndex)
820                   End If
830               Else
840                   If .flags.Navegando = 1 Then
850                       .flags.Navegando = 0
                                  
                          'Tell the client that we are navigating.
860                       Call WriteNavigateToggle(userIndex)
870                   End If
880               End If
890           End If
            
900       End With
End Sub

Public Sub WarpMascotas(ByVal userIndex As Integer, Optional CanWarp As Boolean)
      '************************************************
      'Author: Uknown
      'Last Modified: 11/05/2009
      '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
      '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
      '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
      '************************************************
          Dim i As Integer
          Dim petType As Integer
          Dim PetRespawn As Boolean
          Dim PetTiempoDeVida As Integer
          Dim NroPets As Integer
          Dim InvocadosMatados As Integer
       '   Dim CanWarp As Boolean
          Dim Index As Integer
          Dim iMinHP As Integer
          
10        NroPets = UserList(userIndex).NroMascotas
         ' CanWarp = (MapInfo(UserList(UserIndex).Pos.Map).Pk = True)
          
20        For i = 1 To MAXMASCOTAS
30            Index = UserList(userIndex).MascotasIndex(i)
              
40            If Index > 0 Then
                  ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
50                If Npclist(Index).Contadores.TiempoExistencia > 0 Then
60                    Call QuitarNPC(Index)
70                    UserList(userIndex).MascotasIndex(i) = 0
80                    InvocadosMatados = InvocadosMatados + 1
90                    NroPets = NroPets - 1
                      
100                   petType = 0
110                   UserList(userIndex).NroMascotas = UserList(userIndex).NroMascotas - 1
120               Else
                      'Store data and remove NPC to recreate it after warp
                      'PetRespawn = Npclist(index).flags.Respawn = 0
130                   petType = UserList(userIndex).MascotasType(i)
                      'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                      
                      ' Guardamos el hp, para restaurarlo uando se cree el npc
140                   iMinHP = Npclist(Index).Stats.MinHp
                      
150                   Call QuitarNPC(Index)
                      
                      ' Restauramos el valor de la variable
160                   UserList(userIndex).MascotasType(i) = petType

170               End If
180           ElseIf UserList(userIndex).MascotasType(i) > 0 Then
                  'Store data and remove NPC to recreate it after warp
190               PetRespawn = True
200               petType = UserList(userIndex).MascotasType(i)
210               PetTiempoDeVida = 0
220           Else
230               petType = 0
240           End If
              
250           If petType > 0 And CanWarp Then
260               Index = SpawnNpc(petType, UserList(userIndex).Pos, True, PetRespawn)
                  
                  'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                  ' Exception: Pets don't spawn in water if they can't swim
270               If Index = 0 Then
280                   Call WriteConsoleMsg(userIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
290               Else
300                   UserList(userIndex).MascotasIndex(i) = Index

                      ' Nos aseguramos de que conserve el hp, si estaba dañado
310                   Npclist(Index).Stats.MinHp = IIf(iMinHP = 0, Npclist(Index).Stats.MinHp, iMinHP)
                  
320                   Npclist(Index).MaestroUser = userIndex
330                   Npclist(Index).Movement = TipoAI.SigueAmo
340                   Npclist(Index).Target = 0
350                   Npclist(Index).TargetNPC = 0
360                   Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
370                   Call FollowAmo(Index)
380               End If
390           End If
400       Next i
          
410       If InvocadosMatados > 0 Then
420           Call WriteConsoleMsg(userIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
430       End If
          
440       If Not CanWarp Then
           '   Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
450       End If
          
460       UserList(userIndex).NroMascotas = NroPets
End Sub

Public Sub WarpMascota(ByVal userIndex As Integer, ByVal PetIndex As Integer)
      '************************************************
      'Author: ZaMa
      'Last Modified: 18/11/2009
      'Warps a pet without changing its stats
      '************************************************
          Dim petType As Integer
          Dim NpcIndex As Integer
          Dim iMinHP As Integer
          Dim TargetPos As WorldPos
          
10           With UserList(userIndex)
              
20            TargetPos.map = .flags.TargetMap
30            TargetPos.X = .flags.TargetX
40            TargetPos.Y = .flags.TargetY
              
50            NpcIndex = .MascotasIndex(PetIndex)
                  
              'Store data and remove NPC to recreate it after warp
60            petType = .MascotasType(PetIndex)
              
              ' Guardamos el hp, para restaurarlo cuando se cree el npc
70            iMinHP = Npclist(NpcIndex).Stats.MinHp
              
80            Call QuitarNPC(NpcIndex)
              
              ' Restauramos el valor de la variable
90            .MascotasType(PetIndex) = petType
100           .NroMascotas = .NroMascotas + 1
110           NpcIndex = SpawnNpc(petType, TargetPos, False, False)
              
              'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
              ' Exception: Pets don't spawn in water if they can't swim
120           If NpcIndex = 0 Then
130               Call WriteConsoleMsg(userIndex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
140           Else
150               .MascotasIndex(PetIndex) = NpcIndex

160               With Npclist(NpcIndex)
                      ' Nos aseguramos de que conserve el hp, si estaba dañado
170                   .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
                  
180                   .MaestroUser = userIndex
190                   .Movement = TipoAI.SigueAmo
200                   .Target = 0
210                   .TargetNPC = 0
220               End With
                  
230               Call FollowAmo(NpcIndex)
240           End If
250       End With
End Sub


''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 16/09/2010
      '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
      '***************************************************
          Dim isNotVisible As Boolean
          Dim HiddenPirat As Boolean
          
10        With UserList(userIndex)
20            If .flags.UserLogged And Not .Counters.Saliendo Then
30                .Counters.Saliendo = True
40                .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
                  
50                isNotVisible = (.flags.Oculto Or .flags.invisible)
60                If isNotVisible Then
70                    .flags.invisible = 0
80                    .Counters.Invisibilidad = 0
                      
90                    If .flags.Oculto Then
100                       If .flags.Navegando = 1 Then
110                           If .clase = eClass.Pirat Then
                                  ' Pierde la apariencia de fragata fantasmal
120                               Call ToggleBoatBody(userIndex)
130                               Call WriteConsoleMsg(userIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
140                               Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                                      NingunEscudo, NingunCasco)
150                               HiddenPirat = True
160                           End If
170                       End If
180                   End If
                      
190                   .flags.Oculto = 0
                      
                      
                      ' Para no repetir mensajes
200                   If Not HiddenPirat Then Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                      
                      ' Si esta navegando ya esta visible
210                   If .flags.Navegando = 0 Then
220                       Call SetInvisible(userIndex, .Char.CharIndex, False)
230                   End If
240               End If
                  
250               If .flags.Traveling = 1 Then
260                   Call WriteMultiMessage(userIndex, eMessages.CancelHome)
270                   .flags.Traveling = 0
280                   .Counters.goHome = 0
290               End If
                  
                  
                  'Call WriteConsoleMsg(UserIndex, "Gracias por jugar Desterium AO.", FontTypeNames.FONTTYPE_INFO)
300               Call WriteConsoleMsg(userIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
310           End If
320       End With
End Sub


''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal userIndex As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/02/08
      '
      '***************************************************
10        If UserList(userIndex).Counters.Saliendo Then
              ' Is the user still connected?
20            If UserList(userIndex).ConnIDValida Then
30                UserList(userIndex).Counters.Saliendo = False
40                UserList(userIndex).Counters.Salir = 0
50                Call WriteConsoleMsg(userIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
60            Else
                  'Simply reset
70                UserList(userIndex).Counters.Salir = IIf((UserList(userIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(userIndex).Pos.map).Pk, IntervaloCerrarConexion, 0)
80            End If
90        End If
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ViejoNick As String
          Dim ViejoCharBackup As String
          
10        If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
20        ViejoNick = UserList(UserIndexDestino).Name
          
30        If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
              'hace un backup del char
40            ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
50            Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
60        End If
End Sub

Sub SendUserStatsTxtOFF(ByVal SendIndex As Integer, ByVal Nombre As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
20            Call WriteConsoleMsg(SendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
30        Else
40            Call WriteConsoleMsg(SendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
60            Call WriteConsoleMsg(SendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
70            Call WriteConsoleMsg(SendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
              
80            Call WriteConsoleMsg(SendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)

90            Call WriteConsoleMsg(SendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
              
#If ConUpTime Then
              Dim TempSecs As Long
              Dim tempStr As String
100           TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
110           tempStr = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
120           Call WriteConsoleMsg(SendIndex, "Tiempo Logeado: " & tempStr, FontTypeNames.FONTTYPE_INFO)
#End If
          
130       End If
End Sub

Sub SendUserOROTxtFromChar(ByVal SendIndex As Integer, ByVal charName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim CharFile As String
          
10    On Error Resume Next
20        CharFile = CharPath & charName & ".chr"
          
30        If FileExist(CharFile, vbNormal) Then
40            Call WriteConsoleMsg(SendIndex, charName, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
60        Else
70            Call WriteConsoleMsg(SendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
80        End If
End Sub

Sub VolverCriminal(ByVal userIndex As Integer)
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 21/02/2010
      'Nacho: Actualiza el tag al cliente
      '21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
      '**************************************************************
10        With UserList(userIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
              
30            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
40                .Reputacion.BurguesRep = 0
50                .Reputacion.NobleRep = 0
60                .Reputacion.PlebeRep = 0
70                .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO
80                If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
90                If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userIndex)
                  
100               If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

110           End If
120       End With
          
130       Call RefreshCharStatus(userIndex)
End Sub

Sub VolverCiudadano(ByVal userIndex As Integer)
      '**************************************************************
      'Author: Unknown
      'Last Modify Date: 21/06/2006
      'Nacho: Actualiza el tag al cliente.
      '**************************************************************
10        With UserList(userIndex)
20            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
              
30            .Reputacion.LadronesRep = 0
40            .Reputacion.BandidoRep = 0
50            .Reputacion.AsesinoRep = 0
60            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO
70            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
80        End With
          
90        Call RefreshCharStatus(userIndex)
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 10/07/2008
      'Checks if a given body index is a boat
      '**************************************************************
      'TODO : This should be checked somehow else. This is nasty....
10        If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
                  body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
                  body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
20            BodyIsBoat = True
30        End If
End Function

Public Sub SetInvisible(ByVal userIndex As Integer, ByVal userCharIndex As Integer, ByVal invisible As Boolean)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim sndNick As String

10    With UserList(userIndex)
20        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, userIndex, PrepareMessageSetInvisible(userCharIndex, invisible))
          
30        sndNick = .Name
          
40        If invisible Then
50            sndNick = sndNick & " " & TAG_USER_INVISIBLE
60        Else
70            If .GuildIndex > 0 Then
80                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
90            End If
100       End If
          
110       Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, userIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))
120   End With
End Sub

Public Sub SetConsulatMode(ByVal userIndex As Integer)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 05/06/10
      '
      '***************************************************

      Dim sndNick As String

10    With UserList(userIndex)
20        sndNick = .Name
          
30        If .flags.EnConsulta Then
40            sndNick = sndNick & " " & TAG_CONSULT_MODE
50        Else
60            If .GuildIndex > 0 Then
70                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
80            End If
90        End If
          
100       Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))
110   End With
End Sub

Public Function IsArena(ByVal userIndex As Integer) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 10/11/2009
      'Returns true if the user is in an Arena
      '**************************************************************
10        IsArena = (TriggerZonaPelea(userIndex, userIndex) = TRIGGER6_PERMITE)
End Function

Public Sub PerdioNpc(ByVal userIndex As Integer, Optional ByVal CheckPets As Boolean = True)
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 11/07/2010 (ZaMa)
      'The user loses his owned npc
      '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
      '11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
      '13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
      '**************************************************************

          Dim PetCounter As Long
          Dim PetIndex As Integer
          Dim NpcIndex As Integer
          
10        With UserList(userIndex)
              
20            NpcIndex = .flags.OwnedNpc
30            If NpcIndex > 0 Then
                  
40                If CheckPets Then
                      ' Dejan de atacar las mascotas
50                    If .NroMascotas > 0 Then
60                        For PetCounter = 1 To MAXMASCOTAS
                          
70                            PetIndex = .MascotasIndex(PetCounter)
                              
80                            If PetIndex > 0 Then
                                  ' Si esta atacando al npc deja de hacerlo
90                                If Npclist(PetIndex).TargetNPC = NpcIndex Then
100                                   Call FollowAmo(PetIndex)
110                               End If
120                           End If
                              
130                       Next PetCounter
140                   End If
150               End If
                  
                  ' Reset flags
160               Npclist(NpcIndex).Owner = 0
170               .flags.OwnedNpc = 0

180           End If
190       End With
End Sub

Public Sub ApropioNpc(ByVal userIndex As Integer, ByVal NpcIndex As Integer)
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 18/01/2010 (zaMa)
      'The user owns a new npc
      '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
      '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
      '**************************************************************

10        With UserList(userIndex)
              ' Los admins no se pueden apropiar de npcs
20            If EsGm(userIndex) Then Exit Sub
              
              'No aplica a zonas seguras
30            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
              
              ' No aplica a algunos mapas que permiten el robo de npcs
40            If MapInfo(.Pos.map).RoboNpcsPermitido = 1 Then Exit Sub
              
              ' Pierde el npc anterior
50            If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
              
              ' Si tenia otro dueño, lo perdio aca
60            Npclist(NpcIndex).Owner = userIndex
70            .flags.OwnedNpc = NpcIndex
80        End With
          
          ' Inicializo o actualizo el timer de pertenencia
90        Call IntervaloPerdioNpc(userIndex, True)
End Sub

Public Function GetDireccion(ByVal userIndex As Integer, ByVal OtherUserIndex As Integer) As String
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 17/11/2009
      'Devuelve la direccion hacia donde esta el usuario
      '**************************************************************
          Dim X As Integer
          Dim Y As Integer
          
10        X = UserList(userIndex).Pos.X - UserList(OtherUserIndex).Pos.X
20        Y = UserList(userIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
          
30        If X = 0 And Y > 0 Then
40            GetDireccion = "Sur"
50        ElseIf X = 0 And Y < 0 Then
60            GetDireccion = "Norte"
70        ElseIf X > 0 And Y = 0 Then
80            GetDireccion = "Este"
90        ElseIf X < 0 And Y = 0 Then
100           GetDireccion = "Oeste"
110       ElseIf X > 0 And Y < 0 Then
120           GetDireccion = "NorEste"
130       ElseIf X < 0 And Y < 0 Then
140           GetDireccion = "NorOeste"
150       ElseIf X > 0 And Y > 0 Then
160           GetDireccion = "SurEste"
170       ElseIf X < 0 And Y > 0 Then
180           GetDireccion = "SurOeste"
190       End If

End Function

Public Function SameFaccion(ByVal userIndex As Integer, ByVal OtherUserIndex As Integer) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 17/11/2009
      'Devuelve True si son de la misma faccion
      '**************************************************************
10        SameFaccion = (esCaos(userIndex) And esCaos(OtherUserIndex)) Or _
                          (esArmada(userIndex) And esArmada(OtherUserIndex))
End Function

Public Function FarthestPet(ByVal userIndex As Integer) As Integer
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 18/11/2009
      'Devuelve el indice de la mascota mas lejana.
      '**************************************************************
10    On Error GoTo Errhandler
          
          Dim PetIndex As Integer
          Dim Distancia As Integer
          Dim OtraDistancia As Integer
          
20        With UserList(userIndex)
30            If .NroMascotas = 0 Then Exit Function
          
40            For PetIndex = 1 To MAXMASCOTAS
                  ' Solo pos invocar criaturas que exitan!
50                If .MascotasIndex(PetIndex) > 0 Then
                      ' Solo aplica a mascota, nada de elementales..
60                    If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
70                        If FarthestPet = 0 Then
                              ' Por si tiene 1 sola mascota
80                            FarthestPet = PetIndex
90                            Distancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                          Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
100                       Else
                              ' La distancia de la proxima mascota
110                           OtraDistancia = Abs(.Pos.X - Npclist(.MascotasIndex(PetIndex)).Pos.X) + _
                                              Abs(.Pos.Y - Npclist(.MascotasIndex(PetIndex)).Pos.Y)
                              ' Esta mas lejos?
120                           If OtraDistancia > Distancia Then
130                               Distancia = OtraDistancia
140                               FarthestPet = PetIndex
150                           End If
160                       End If
170                   End If
180               End If
190           Next PetIndex
200       End With

210       Exit Function
          
Errhandler:
220       Call LogError("Error en FarthestPet")
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal userIndex As Integer, ByVal Skill As Byte, ByVal Allocation As Boolean)
      '*************************************************
      'Author: Torres Patricio (Pato)
      'Last modified: 11/20/2009
      '
      '*************************************************

10    With UserList(userIndex).Stats
20        If .UserSkills(Skill) < MAXSKILLPOINTS Then
30            If Allocation Then
40                .ExpSkills(Skill) = 0
50            Else
60                .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)
70            End If
              
80            .EluSkills(Skill) = ELU_SKILL_INICIAL * 1 ^ .UserSkills(Skill)
90        Else
100           .ExpSkills(Skill) = 0
110           .EluSkills(Skill) = 0
120       End If
130   End With

End Sub

Public Function HasEnoughItems(ByVal userIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long) As Boolean
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 25/11/2009
      'Cheks Wether the user has the required amount of items in the inventory or not
      '**************************************************************

          Dim Slot As Long
          Dim ItemInvAmount As Long
          
10        For Slot = 1 To UserList(userIndex).CurrentInventorySlots
              ' Si es el item que busco
20            If UserList(userIndex).Invent.Object(Slot).ObjIndex = ObjIndex Then
                  ' Lo sumo a la cantidad total
30                ItemInvAmount = ItemInvAmount + UserList(userIndex).Invent.Object(Slot).Amount
40            End If
50        Next Slot

60        HasEnoughItems = Amount <= ItemInvAmount
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, ByVal userIndex As Integer) As Long
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 25/11/2009
      'Cheks the amount of items the user has in offerSlots.
      '**************************************************************
          Dim Slot As Byte
          
10        For Slot = 1 To MAX_OFFER_SLOTS
                  ' Si es el item que busco
20            If UserList(userIndex).ComUsu.Objeto(Slot) = ObjIndex Then
                  ' Lo sumo a la cantidad total
30                TotalOfferItems = TotalOfferItems + UserList(userIndex).ComUsu.cant(Slot)
40            End If
50        Next Slot

End Function

Public Function getMaxInventorySlots(ByVal userIndex As Integer) As Byte
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    If UserList(userIndex).Invent.MochilaEqpObjIndex > 0 Then
20        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(userIndex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
30    Else
40        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
50    End If
End Function

Public Sub goHome(ByVal userIndex As Integer)
      Dim Distance As Integer
      Dim tiempo As Long

10    With UserList(userIndex)
20        If .flags.Muerto = 1 Then
30            If .flags.lastMap = 0 Then
40                Distance = distanceToCities(.Pos.map).distanceToCity(.Hogar)
50            Else
60                Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY
70            End If
              
80            tiempo = (Distance + 1) * 30 'segundos
              
90            .Counters.goHome = tiempo / 6 'Se va a chequear cada 6 segundos.
              
100           .flags.Traveling = 1

110           Call WriteMultiMessage(userIndex, eMessages.Home, Distance, tiempo, , MapInfo(Ciudades(.Hogar).map).Name)
120       Else
130           Call WriteConsoleMsg(userIndex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)
140       End If
150   End With
End Sub

Public Function ToogleToAtackable(ByVal userIndex As Integer, ByVal OwnerIndex As Integer, Optional ByVal StealingNpc As Boolean = True) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 15/01/2010
      'Change to Atackable mode.
      '***************************************************
          
          Dim AtacablePor As Integer
          
10        With UserList(userIndex)
              ' Inicializar el timer
20            Call IntervaloEstadoAtacable(userIndex, True)
              
30            ToogleToAtackable = True
              
40        End With
          
End Function

Public Sub setHome(ByVal userIndex As Integer, ByVal newHome As eCiudad, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Budi
      'Last Modification: 30/04/2010
      '30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
      '***************************************************
10        If newHome < eCiudad.cUllathorpe Or newHome > cArghal Then Exit Sub
20        UserList(userIndex).Hogar = newHome
          
30        Call WriteChatOverHead(userIndex, "¡¡¡Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
End Sub

Public Sub ToggleBoatBody(ByVal userIndex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 25/07/2010
      'Gives boat body depending on user alignment.
      '25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
      '***************************************************

          Dim Ropaje As Integer
          Dim EsFaccionario As Boolean
          Dim NewBody As Integer
          
10        With UserList(userIndex)
       
20            .Char.Head = 0
30            If .Invent.BarcoObjIndex = 0 Then Exit Sub
              
40            Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
              
              ' Criminales y caos
50            If criminal(userIndex) Then
                  
60                EsFaccionario = esCaos(userIndex)
                  
70                Select Case Ropaje
                      Case iBarca
80                        If EsFaccionario Then
90                            NewBody = iBarcaPk
100                       Else
110                           NewBody = iBarcaPk
120                       End If
                      
130                   Case iGalera
140                       If EsFaccionario Then
150                           NewBody = iGaleraPk
160                       Else
170                           NewBody = iGaleraPk
180                       End If
                          
190                   Case iGaleon
200                       If EsFaccionario Then
210                           NewBody = iGaleonPk
220                       Else
230                           NewBody = iGaleonPk
240                       End If
250               End Select
              
              ' Ciudas y Armadas
260           Else
                  
270               EsFaccionario = esArmada(userIndex)
                  
                  ' Atacable
280               If .flags.AtacablePor <> 0 Then
                      
290                   Select Case Ropaje
                          Case iBarca
300                           If EsFaccionario Then
310                               NewBody = iBarcaCiuda
320                           Else
330                               NewBody = iBarcaCiuda
340                           End If
                          
350                       Case iGalera
360                           If EsFaccionario Then
370                               NewBody = iGaleraCiuda
380                           Else
390                               NewBody = iGaleraCiuda
400                           End If
                              
410                       Case iGaleon
420                           If EsFaccionario Then
430                               NewBody = iGaleonCiuda
440                           Else
450                               NewBody = iGaleonCiuda
460                           End If
470                   End Select
                  
                  ' Normal
480               Else
                  
490                   Select Case Ropaje
                          Case iBarca
500                           If EsFaccionario Then
510                               NewBody = iBarcaCiuda
520                           Else
530                               NewBody = iBarcaCiuda
540                           End If
                          
550                       Case iGalera
560                           If EsFaccionario Then
570                               NewBody = iGaleraCiuda
580                           Else
590                               NewBody = iGaleraCiuda
600                           End If
                              
610                       Case iGaleon
620                           If EsFaccionario Then
630                               NewBody = iGaleonCiuda
640                           Else
650                               NewBody = iGaleonCiuda
660                           End If
670                   End Select
                  
680               End If
                  
690           End If
              
700           .Char.body = NewBody
710           .Char.ShieldAnim = NingunEscudo
720           .Char.WeaponAnim = NingunArma
730           .Char.CascoAnim = NingunCasco
740       End With

End Sub

Sub SendUserStatsMercado(ByVal SendIndex As Integer, ByVal userIndex As Integer)
      '*************************************************
      'Author: Unknown
      'Last modified: 23/01/2007
      'Shows the users Stats when the user is online.
      '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
      '*************************************************
10        With UserList(userIndex)
20            Call WriteConsoleMsg(SendIndex, "Nick: " & .Name, FontTypeNames.FONTTYPE_INFO)
30            Call WriteConsoleMsg(SendIndex, "Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
40            Call WriteConsoleMsg(SendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
50            Call WriteConsoleMsg(SendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
60            Call WriteConsoleMsg(SendIndex, "Penas: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
70             Call WriteConsoleMsg(SendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
80            Call WriteConsoleMsg(SendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Maná: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
              
90            If .Faccion.ArmadaReal = 1 Then
100               Call WriteConsoleMsg(SendIndex, "Ejército real desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
110               Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso & " con " & .Faccion.MatadosIngreso & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
120               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
130           ElseIf .Faccion.FuerzasCaos = 1 Then
140               Call WriteConsoleMsg(SendIndex, "Legión oscura desde: " & .Faccion.FechaIngreso, FontTypeNames.FONTTYPE_INFO)
150               Call WriteConsoleMsg(SendIndex, "Ingresó en nivel: " & .Faccion.NivelIngreso, FontTypeNames.FONTTYPE_INFO)
160               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
170           ElseIf .Faccion.RecibioExpInicialReal = 1 Then
180               Call WriteConsoleMsg(SendIndex, "Fue ejército real", FontTypeNames.FONTTYPE_INFO)
190               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
              
200           ElseIf .Faccion.RecibioExpInicialCaos = 1 Then
210               Call WriteConsoleMsg(SendIndex, "Fue legión oscura", FontTypeNames.FONTTYPE_INFO)
220               Call WriteConsoleMsg(SendIndex, "Veces que ingresó: " & .Faccion.Reenlistadas, FontTypeNames.FONTTYPE_INFO)
230           End If
              
240           Call WriteConsoleMsg(SendIndex, "Asesino: " & .Reputacion.AsesinoRep, FontTypeNames.FONTTYPE_INFO)
250           Call WriteConsoleMsg(SendIndex, "Noble: " & .Reputacion.NobleRep, FontTypeNames.FONTTYPE_INFO)
              
260           Call WriteConsoleMsg(SendIndex, "Oro: " & .Stats.Gld & "  Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
270           Call WriteConsoleMsg(SendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
280           Call WriteConsoleMsg(SendIndex, "Retos Ganados: " & .Stats.RetosGanados & "", FontTypeNames.FONTTYPE_INFO)
290           Call WriteConsoleMsg(SendIndex, "Retos Perdidos: " & .Stats.RetosPerdidos & "", FontTypeNames.FONTTYPE_INFO)
300           Call WriteConsoleMsg(SendIndex, "Oro Ganado: " & .Stats.OroGanado & "", FontTypeNames.FONTTYPE_INFO)
310           Call WriteConsoleMsg(SendIndex, "Oro Perdido: " & .Stats.OroPerdido & "", FontTypeNames.FONTTYPE_INFO)
320           Call WriteConsoleMsg(SendIndex, "Torneos Ganados: " & .Stats.TorneosGanados & "", FontTypeNames.FONTTYPE_INFO)
330           Call WriteConsoleMsg(SendIndex, "Famas asignadas: " & .flags.BonosHP & "", FontTypeNames.FONTTYPE_INFO)
              
              '?
340           If .GuildIndex > 0 Then
350               Call WriteConsoleMsg(SendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
360           End If
370       End With
End Sub
