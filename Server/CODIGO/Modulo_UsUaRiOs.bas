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
    '<EhHeader>
    On Error GoTo ActStats_Err

    '</EhHeader>

    Dim DaExp       As Integer

    Dim EraCriminal As Boolean
    
    DaExp = CInt(UserList(VictimIndex).Stats.Elv) * 2
    
    With UserList(AttackerIndex)
        .Stats.Exp = .Stats.Exp + DaExp

        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        Call CheckUserLevel(AttackerIndex)
        
        If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
        
            ' Es legal matarlo si estaba en atacable
            If UserList(VictimIndex).flags.AtacablePor <> AttackerIndex Then
                EraCriminal = Escriminal(AttackerIndex)
                
                With .Reputacion

                    If Not Escriminal(VictimIndex) Then
                        .AsesinoRep = .AsesinoRep + vlASESINO * 2

                        If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
                        .BurguesRep = 0
                        .NobleRep = 0
                        .PlebeRep = 0
                    Else
                        .NobleRep = .NobleRep + vlNoble

                        If .NobleRep > MAXREP Then .NobleRep = MAXREP

                    End If

                End With

                If EraCriminal <> Escriminal(AttackerIndex) Then
                    Call RefreshCharStatus(AttackerIndex)

                End If
                
            End If

        End If
        
        'Lo mata
        Call WriteMultiMessage(AttackerIndex, eMessages.HaveKilledUser, VictimIndex, DaExp)
        Call WriteMultiMessage(VictimIndex, eMessages.UserKill, AttackerIndex)
        Call FlushBuffer(VictimIndex)
        
        'Log
        'Call Logs_Security(eSecurity, eAntiFrags, .Name & " con IP: " & .Ip & " Y Cuenta: " & .Account.Email & " asesino a " & UserList(VictimIndex).Name & " con IP: " & UserList(VictimIndex).Ip & " Y Cuenta: " & UserList(VictimIndex).Account.Email)
        'Call Logs_User(.Name, eUser, eKill, .Name & " con IP: " & .Ip & " Y Cuenta: " & .Account.Email & " asesino a " & UserList(VictimIndex).Name & " con IP: " & UserList(VictimIndex).Ip & " Y Cuenta: " & UserList(VictimIndex).Account.Email)
    End With

    '<EhFooter>
    Exit Sub

ActStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ActStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub RevivirUsuario(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo RevivirUsuario_Err

    '</EhHeader>

    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion) * 5
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp

        End If
        
        If .flags.Navegando = 1 Then
            Call ToggleBoatBody(UserIndex)
        Else
            Call DarCuerpoDesnudo(UserIndex)
            
            .Char.Head = .OrigChar.Head

        End If
        
        If .flags.Traveling Then
            Call EndTravel(UserIndex, True)

        End If
        
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        Call WriteUpdateUserStats(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

RevivirUsuario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.RevivirUsuario " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'Gives boat body depending on user alignment.
    '25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
    '***************************************************
    '<EhHeader>
    On Error GoTo ToggleBoatBody_Err

    '</EhHeader>

    Dim Ropaje        As Integer

    Dim EsFaccionario As Boolean

    Dim NewBody       As Integer
    
    With UserList(UserIndex)
 
        .Char.Head = 0

        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
        Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
            
        If Ropaje = 0 Then
            
            ' Criminales y caos
            If Escriminal(UserIndex) Then
                
                EsFaccionario = esCaos(UserIndex)
                
                Select Case Ropaje

                    Case iBarca

                        If EsFaccionario Then
                            NewBody = iBarcaCaos
                        Else
                            NewBody = iBarcaPk

                        End If
                    
                    Case iGalera

                        If EsFaccionario Then
                            NewBody = iGaleraCaos
                        Else
                            NewBody = iGaleraPk

                        End If
                        
                    Case iGaleon

                        If EsFaccionario Then
                            NewBody = iGaleonCaos
                        Else
                            NewBody = iGaleonPk

                        End If

                End Select
            
                ' Ciudas y Armadas
            Else
                
                EsFaccionario = esArmada(UserIndex)
                
                Select Case Ropaje

                    Case iBarca

                        If EsFaccionario Then
                            NewBody = iBarcaReal
                        Else
                            NewBody = iBarcaCiuda

                        End If
                        
                    Case iGalera

                        If EsFaccionario Then
                            NewBody = iGaleraReal
                        Else
                            NewBody = iGaleraCiuda

                        End If
                            
                    Case iGaleon

                        If EsFaccionario Then
                            NewBody = iGaleonReal
                        Else
                            NewBody = iGaleonCiuda

                        End If

                End Select
                
            End If
                
        Else
            NewBody = Ropaje

        End If
            
        .Char.Body = NewBody
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
                
        Dim A As Long
              
        For A = 1 To MAX_AURAS
            .Char.AuraIndex(A) = NingunAura
        Next A
              
    End With

    '<EhFooter>
    Exit Sub

ToggleBoatBody_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ToggleBoatBody " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ChangeUserChar(ByVal UserIndex As Integer, _
                          ByVal Body As Integer, _
                          ByVal Head As Integer, _
                          ByVal Heading As Byte, _
                          ByVal Arma As Integer, _
                          ByVal Escudo As Integer, _
                          ByVal Casco As Integer, _
                          ByRef AuraIndex() As Byte)

    '<EhHeader>
    On Error GoTo ChangeUserChar_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    With UserList(UserIndex).Char
        .Body = Body
        .Head = Head
        .Heading = Heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = Casco
              
        Dim A As Long

        For A = 1 To MAX_AURAS
            .AuraIndex(A) = AuraIndex(A)
        Next A
             
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, 0, Head, Heading, .charindex, Arma, Escudo, .FX, .loops, Casco, AuraIndex, UserList(UserIndex).flags.ModoStream, False, UserList(UserIndex).flags.Navegando))

    End With

    '<EhFooter>
    Exit Sub

ChangeUserChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ChangeUserChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function GetArmourAnim_Bot(ByVal Slot As Long, _
                                  ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetArmourAnim_Bot_Err

    '</EhHeader>

    '***************************************************
    '
    '
    '
    '***************************************************
    Dim Tmp          As Integer

    Dim SkinSelected As Integer
        
    With BotIntelligence(Slot)
            
        ' If .Skins.Armour > 0 Then
        'ObjIndex = .Skins.Armour
        ' End If

        Tmp = ObjData(ObjIndex).RopajeEnano
                
        If Tmp > 0 Then
            If .Raze = eRaza.Enano Or .Raze = eRaza.Gnomo Then
                GetArmourAnim_Bot = Tmp

                Exit Function

            End If

        End If
        
        GetArmourAnim_Bot = ObjData(ObjIndex).Ropaje

    End With

    '<EhFooter>
    Exit Function

GetArmourAnim_Bot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetArmourAnim " & "at line " & Erl

    '</EhFooter>
End Function

Public Function GetArmourAnim(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetArmourAnim_Err

    '</EhHeader>

    '***************************************************
    '
    '
    '
    '***************************************************
    Dim Tmp          As Integer

    Dim SkinSelected As Integer
        
    With UserList(UserIndex)

        If .Skins.ArmourIndex > 0 Then
            ObjIndex = .Skins.ArmourIndex

        End If

        Tmp = ObjData(ObjIndex).RopajeEnano
                
        If Tmp > 0 Then
            If .Raza = eRaza.Enano Or .Raza = eRaza.Gnomo Then
                GetArmourAnim = Tmp

                Exit Function

            End If

        End If
        
        GetArmourAnim = ObjData(ObjIndex).Ropaje

    End With

    '<EhFooter>
    Exit Function

GetArmourAnim_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetArmourAnim " & "at line " & Erl

    '</EhFooter>
End Function

Public Function GetShieldAnim(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetShieldAnim_Err

    '</EhHeader>
    With UserList(UserIndex)

        If .Skins.ShieldIndex > 0 Then
            ObjIndex = .Skins.ShieldIndex

        End If
        
        GetShieldAnim = ObjData(ObjIndex).ShieldAnim

    End With
    
    '<EhFooter>
    Exit Function

GetShieldAnim_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetShieldAnim " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Function GetHelmAnim(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetHelmAnim_Err

    '</EhHeader>
    With UserList(UserIndex)

        If .Skins.HelmIndex > 0 Then
            ObjIndex = .Skins.HelmIndex

        End If
        
        GetHelmAnim = ObjData(ObjIndex).CascoAnim

    End With
    
    '<EhFooter>
    Exit Function

GetHelmAnim_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetHelmAnim " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Function GetWeaponAnim(ByVal UserIndex As Integer, _
                              ByVal UserRaza As Byte, _
                              ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetWeaponAnim_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/29/10
    '
    '***************************************************
    Dim Tmp As Integer

    With UserList(UserIndex)

        If ObjData(ObjIndex).Apuñala = 1 Then
            If .Skins.WeaponDagaIndex > 0 Then
                ObjIndex = .Skins.WeaponDagaIndex

            End If

        ElseIf ObjData(ObjIndex).proyectil > 0 Then

            If .Skins.WeaponArcoIndex > 0 Then
                ObjIndex = .Skins.WeaponArcoIndex

            End If

        Else

            If .Skins.WeaponIndex > 0 Then
                ObjIndex = .Skins.WeaponIndex

            End If

        End If
            
    End With

    Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
    If Tmp > 0 Then
        If UserRaza = eRaza.Enano Or UserRaza = eRaza.Gnomo Then
            GetWeaponAnim = Tmp

            Exit Function

        End If

    End If
        
    GetWeaponAnim = ObjData(ObjIndex).WeaponAnim

    '<EhFooter>
    Exit Function

GetWeaponAnim_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetWeaponAnim " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function GetWeaponAnimBot(ByVal Raza As Byte, ByVal ObjIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetWeaponAnimBot_Err

    '</EhHeader>

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/29/10
    '
    '***************************************************
    Dim Tmp As Integer

    Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
    If Tmp > 0 Then
        If Raza = eRaza.Enano Or Raza = eRaza.Gnomo Then
            GetWeaponAnimBot = Tmp

            Exit Function

        End If

    End If
        
    GetWeaponAnimBot = ObjData(ObjIndex).WeaponAnim

    '<EhFooter>
    Exit Function

GetWeaponAnimBot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetWeaponAnim " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub EraseUserChar(ByVal UserIndex As Integer, ByVal IsAdminInvisible As Boolean)

    '*************************************************
    'Author: Unknown
    'Last modified: 08/01/2009
    '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
    '*************************************************
    '<EhHeader>
    On Error GoTo EraseUserChar_Err

    '</EhHeader>

    With UserList(UserIndex)

        CharList(.Char.charindex) = 0
        
        If .Char.charindex > 0 And .Char.charindex <= LastChar Then
            CharList(.Char.charindex) = 0
            
            If .Char.charindex = LastChar Then

                Do Until CharList(LastChar) > 0
                    LastChar = LastChar - 1

                    If LastChar <= 1 Then Exit Do
                Loop

            End If

        End If
                
        Call ModAreas.DeleteEntity(UserIndex, ENTITY_TYPE_PLAYER)
        
        If MapaValido(.Pos.Map) Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0

        End If
                
        .Char.charindex = 0

    End With
    
    NumChars = NumChars - 1

    '<EhFooter>
    Exit Sub

EraseUserChar_Err:

    Dim UserName  As String

    Dim charindex As Integer
    
    If UserIndex > 0 Then
        UserName = UserList(UserIndex).Name
        charindex = UserList(UserIndex).Char.charindex

    End If

    Call LogError("Error en EraseUserchar " & Err.number & ": " & Err.description & ". User: " & UserName & "(UI: " & UserIndex & " - CI: " & charindex & ") en Linea: " & Erl)
        
    '</EhFooter>
End Sub

Public Sub RefreshCharStatus(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo RefreshCharStatus_Err

    '</EhHeader>

    '*************************************************
    'Author: Tararira
    'Last modified: 04/07/2009
    'Refreshes the status and tag of UserIndex.
    '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
    '*************************************************
    Dim ClanTag   As String

    Dim NickColor As Byte
    
    With UserList(UserIndex)

        If .GuildIndex > 0 Then
            ClanTag = GuildsInfo(.GuildIndex).Name
            ClanTag = " <" & ClanTag & ">"

        End If
        
        NickColor = GetNickColor(UserIndex)
        
        If .ShowName Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, NickColor, vbNullString))

        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.Body = iFragataFantasmal
            Else
                Call ToggleBoatBody(UserIndex)

            End If
            
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

RefreshCharStatus_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.RefreshCharStatus " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function GetNickColor(ByVal UserIndex As Integer) As Byte

    '*************************************************
    'Author: ZaMa
    'Last modified: 15/01/2010
    '
    '*************************************************
    '<EhHeader>
    On Error GoTo GetNickColor_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        If Escriminal(UserIndex) Then
            GetNickColor = eNickColor.ieCriminal
        ElseIf Not Escriminal(UserIndex) Then
            GetNickColor = eNickColor.ieCiudadano

        End If
        
        If .Faction.Status = r_Armada Then
            GetNickColor = eNickColor.ieArmada

        End If
        
        If .Faction.Status = r_Caos Then
            GetNickColor = eNickColor.ieCAOS

        End If
        
        If .flags.FightTeam = 1 Then
            GetNickColor = eNickColor.ieCriminal
        ElseIf .flags.FightTeam = 2 Then
            GetNickColor = eNickColor.ieCiudadano

        End If
        
        'If .flags.AtacablePor > 0 Then GetNickColor = GetNickColor Or eNickColor.ieAtacable
        
        If .Counters.Shield Then
            GetNickColor = eNickColor.ieShield

        End If
        
        If Power.UserIndex = UserIndex Then GetNickColor = eNickColor.ieAtacable
        
    End With
    
    '<EhFooter>
    Exit Function

GetNickColor_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetNickColor " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function MakeUserChar(ByVal toMap As Boolean, _
                             ByVal sndIndex As Integer, _
                             ByVal UserIndex As Integer, _
                             ByVal Map As Integer, _
                             ByVal X As Integer, _
                             ByVal Y As Integer, _
                             Optional ByVal ButIndex As Boolean = False, _
                             Optional ByVal IsInvi As Boolean = False) As Boolean
    '*************************************************
    'Author: Unknown
    'Last modified: 15/01/2010
    '23/07/2009: Budi - Ahora se envía el nick
    '15/01/2010: ZaMa - Ahora se envia el color del nick.
    '*************************************************

    On Error GoTo ErrHandler

    Dim charindex  As Integer

    Dim ClanTag    As String

    Dim NickColor  As Byte

    Dim UserName   As String

    Dim Privileges As Byte
    
    With UserList(UserIndex)
    
        If InMapBounds(Map, X, Y) Then

            'If needed make a new character in list
            If .Char.charindex = 0 Then
                charindex = NextOpenCharIndex
                .Char.charindex = charindex
                CharList(charindex) = UserIndex

            End If
            
            'Place character on map if needed
            If toMap Then MapData(Map, X, Y).UserIndex = UserIndex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = GuildsInfo(.GuildIndex).Name

                End If
                
                NickColor = GetNickColor(UserIndex)
                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .ShowName Then
                    UserName = .secName
                    
                    If EsSemiDios(UserName) Then
                        UserName = UserName & " " & TAG_GAME_MASTER

                    End If
                    
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else

                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User) Then
                            If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                        Else

                            If (.flags.Invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else

                                If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"

                            End If

                        End If

                    End If

                End If
                
                Call WriteCharacterCreate(sndIndex, .Char.Body, 0, .Char.Head, .Char.Heading, .Char.charindex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, UserName, NickColor, Privileges, .Char.AuraIndex, .Char.speeding, False)
                
                If IsInvi Then

                    'Actualizamos las áreas de ser necesario
                    'Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)
                End If

            Else
                ' Me lo mando a mi mismo
                Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y)
                
                ' Se lo mando a los demas
                Call ModAreas.CreateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, ModAreas.DEFAULT_ENTITY_WIDTH, ModAreas.DEFAULT_ENTITY_HEIGHT)

            End If

        End If

    End With

    MakeUserChar = True
    
    Exit Function

ErrHandler:

    Dim UserErrName As String

    Dim UserMap     As Integer

    If UserIndex > 0 Then
        UserErrName = UserList(UserIndex).Name
        UserMap = UserList(UserIndex).Pos.Map

    End If
    
    Dim sError As String

    sError = "MakeUserChar: num: " & Err.number & " desc: " & Err.description & ".User: " & UserErrName & "(" & UserIndex & "). UserMap: " & UserMap & ". Coor: " & Map & "," & X & "," & Y & ". toMap: " & toMap & ". sndIndex: " & sndIndex & ". CharIndex: " & charindex & ". ButIndex: " & ButIndex
    
    '
    Call CloseSocket(UserIndex)
    
    'Para ver si clona..
    sError = sError & ". MapUserIndex: " & MapData(Map, X, Y).UserIndex
    Call LogError(sError)

End Function

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo CheckUserLevel_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 08/04/2011
    'Chequea que el usuario no halla alcanzado el siguiente nivel,
    'de lo contrario le da la vida, mana, etc, correspodiente.
    '07/08/2006 Integer - Modificacion de los valores
    '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
    '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
    '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
    '12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y está en un clan, se lo expulsa para no sumar antifacción
    '02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
    '11/19/2009 Pato - Modifico la nueva fórmula de maná ganada para el bandido y se la limito a 499
    '02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
    '08/04/2011: Amraphen - Arreglada la distribución de probabilidades para la vida en el caso de promedio entero.
    '*************************************************
    Dim Pts              As Integer

    Dim WasNewbie        As Boolean

    Dim promedio         As Double

    Dim aux              As Integer

    Dim DistVida(1 To 5) As Integer

    Dim GI               As Integer 'Guild Index
    
    Dim aumentoHp        As Integer

    Dim AumentoMana      As Integer

    Dim AumentoSta       As Integer

    Dim AumentoHit       As Integer
    
    Dim pasoDeNivel      As Boolean
    
    WasNewbie = EsNewbie(UserIndex)
    
    With UserList(UserIndex)

        'Checkea si alcanzó el máximo nivel
        If .Stats.Elv >= STAT_MAXELV Then
            .Stats.Exp = 0
            .Stats.Elu = 0
                
            Exit Sub

        End If

        Do While .Stats.Exp >= .Stats.Elu And .Stats.Elv < STAT_MAXELV
            
            pasoDeNivel = True

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_NIVEL, .Pos.X, .Pos.Y, .Char.charindex))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, FXIDs.FX_LEVEL, 0))
            
            If .Stats.Elv = 1 Then
                Pts = 10
            Else
                'For multiple levels being rised at once
                Pts = Pts + 5

            End If
                
            Dim LastMap As Integer
                
            If MapInfo(.Pos.Map).LvlMax > 0 Then
                If .Stats.Elv >= MapInfo(.Pos.Map).LvlMax Then
                    Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
                    LastMap = .Pos.Map

                End If

            End If
            
            .Stats.Elv = .Stats.Elv + 1
                    
            If .Stats.Elv = 35 Then
                If .Hogar = eCiudad.cEsperanza Then
                    .Hogar = eCiudad.cUllathorpe
                    Call WriteConsoleMsg(UserIndex, "Tu nuevo hogar pasó a ser la Ciudad de Ullathorpe.", FontTypeNames.FONTTYPE_USERGOLD)

                End If

            End If
                    
            #If Classic = 1 Then
                    
                'Esta haciendo la mision newbie. Pasamos a la siguiente
                If .Stats.Elv = LimiteNewbie + 1 Then
                    If .QuestStats(1).QuestIndex > 0 Then
                        .QuestStats(1).QuestIndex = 0
                        Call Quest_SetUser(UserIndex, 2)

                    End If
                        
                End If
                        
            #End If
                
            If .Stats.Elv >= 35 Then

                ' # Envia un mensaje a discord
                Dim TextDiscord As String

                TextDiscord = "El personaje **'" & .Name & "'** pasó a **nivel " & .Stats.Elv & "**."
                WriteMessageDiscord CHANNEL_LEVEL, TextDiscord

            End If
                   
            If .Stats.Elv = STAT_MAXELV Then
                .Stats.SkillPts = .Stats.SkillPts + 5
                Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                Call WriteConsoleMsg(UserIndex, "Has ganado: " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El personaje " & .Name & " ha alcanzado el nivel máximo.", FontTypeNames.FONTTYPE_INFOGREEN))

                'Call RankUser_AddPoint(UserIndex, 100)
            End If
            
            .Stats.Exp = .Stats.Exp - .Stats.Elu
            .Stats.Elu = EluUser(.Stats.Elv)
            RecompensaPorNivel UserIndex
            
            .Stats.MinHp = .Stats.MaxHp
                
        Loop
    
        If pasoDeNivel Then

            'Send all gained skill points at once (if any)
            If Pts > 0 Then
                Call WriteLevelUp(UserIndex, .Stats.SkillPts + Pts)
                  
                .Stats.SkillPts = .Stats.SkillPts + Pts
                Call WriteConsoleMsg(UserIndex, "Has ganado: " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)

            End If
                
            If LastMap > 0 Then
                Call WriteConsoleMsg(UserIndex, "Vende los objetos que obtuviste en el dungeon y compra algunas pociones más abajo. Busca algun equipamiento básico y recorre el mundo. Ademas puedes verlo desde el botón de arriba.", FontTypeNames.FONTTYPE_USERGOLD)

            End If
                
            ' Comprueba si debe scar objetos
            Call QuitarLevelObj(UserIndex)
            Call WriteUpdateUserStats(UserIndex)
            Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr")
        Else
                
            Call WriteUpdateExp(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

CheckUserLevel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CheckUserLevel " & "at line " & Erl & " Valor Exp: " & UserList(UserIndex).Stats.Exp & " Level User: " & UserList(UserIndex).Stats.Elv

    '</EhFooter>
End Sub

Public Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PuedeAtravesarAgua_Err

    '</EhHeader>

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1 Or UserList(UserIndex).flags.Vuela = 1
    '<EhFooter>
    Exit Function

PuedeAtravesarAgua_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.PuedeAtravesarAgua " & "at line " & Erl
        
    '</EhFooter>
End Function

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: 13/07/2009
    'Moves the char, sending the message to everyone in range.
    '30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
    '28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
    '13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
    '13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
    '*************************************************
    '<EhHeader>
    On Error GoTo MoveUserChar_Err

    '</EhHeader>

    Dim nPos               As WorldPos

    Dim sailing            As Boolean

    Dim CasperIndex        As Integer

    Dim CasperHeading      As eHeading

    Dim isAdminInvi        As Boolean

    Dim isZonaOscura       As Boolean

    Dim isZonaOscuraNewPos As Boolean

    Dim UserMoved          As Boolean
 
    sailing = PuedeAtravesarAgua(UserIndex)
    nPos = UserList(UserIndex).Pos
    isZonaOscura = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)

    Call HeadtoPos(nHeading, nPos)

    isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)
    isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)

    If MoveToLegalPos(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y, sailing, Not sailing) Then
        UserMoved = True

        ' si no estoy solo en el mapa...
        If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
            CasperIndex = MapData(UserList(UserIndex).Pos.Map, nPos.X, nPos.Y).UserIndex

            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then

                ' Los admins invisibles no pueden patear caspers
                If Not isAdminInvi Then

                    With UserList(CasperIndex)
                    
                        If TriggerZonaPelea(UserIndex, CasperIndex) = TRIGGER6_PROHIBE Then
                            If .flags.SeguroResu = False Then
                                .flags.SeguroResu = True
                                Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)

                            End If
                             
                        End If
                            
                        '  If .LastHeading > 0 Then
                        '  CasperHeading = .LastHeading
                        ' .LastHeading = 0
                        'Else
                        CasperHeading = InvertHeading(nHeading)

                        'End If

                        '.LastHeading = .Char.Heading

                        Call HeadtoPos(CasperHeading, .Pos)

                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not (.flags.AdminInvisible = 1) Then
                            ' Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y))
                        
                            'Los valores de visible o invisible están invertidos porque estos flags son del UserIndex, por lo tanto si el UserIndex entra, el casper sale y viceversa :P
                            If isZonaOscura Then
                                If Not isZonaOscuraNewPos Then
                                    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                                End If

                            Else

                                If isZonaOscuraNewPos Then
                                    Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, CasperIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                                End If

                            End If

                        End If

                        Call WriteForceCharMove(CasperIndex, CasperHeading)

                        'Update map and char
                        .Char.Heading = CasperHeading
                        MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = CasperIndex
                        
                        'Actualizamos las áreas de ser necesario
                        Call ModAreas.UpdateEntity(CasperIndex, ENTITY_TYPE_PLAYER, .Pos)

                    End With

                End If

            End If

            ' Si es un admin invisible, no se avisa a los demas clientes
            'If Not isAdminInvi Then Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.X, nPos.Y))
            
        End If

        ' Los admins invisibles no pueden patear caspers
        If (Not isAdminInvi) Or (CasperIndex = 0) Then

            With UserList(UserIndex)

                ' Si no hay intercambio de pos con nadie
                If CasperIndex = 0 Then
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0

                End If

                .Pos = nPos
                .Char.Heading = nHeading
                MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                
                If Extra.IsAreaResu(UserIndex) Then
                    Call Extra.AutoCurar(UserIndex)

                End If

                If isZonaOscura Then
                    If Not isZonaOscuraNewPos Then
                        If (.flags.Invisible Or .flags.Oculto) = 0 Then
                            Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                        End If

                    End If

                Else

                    If isZonaOscuraNewPos Then
                        If (.flags.Invisible Or .flags.Oculto) = 0 Then
                            Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                        End If

                    End If

                End If

                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).Modality = Busqueda Then
                        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjEvent = 1 Then
                            Call EventosDS.Busqueda_GetObj(.flags.SlotEvent, .flags.SlotUserEvent)
                            MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjEvent = 0
                            EraseObj 10000, .Pos.Map, .Pos.X, .Pos.Y

                        End If

                    End If

                End If

                ' // NUEVO
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then
                    Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)

                End If
                
                'Actualizamos las áreas de ser necesario
                Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)

            End With

        Else
            Call WritePosUpdate(UserIndex)

        End If

    Else
        Call WritePosUpdate(UserIndex)

    End If

    If UserList(UserIndex).Counters.Trabajando Then UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
        
    MoveUserChar = UserMoved
    
    '<EhFooter>
    Exit Function

MoveUserChar_Err:
    Call LogError("Error " & Err.number & " (Linea: " & Erl & ") " & Err.description & " en User: " & UserList(UserIndex).Name & " con IP: " & UserList(UserIndex).IpAddress & " con pos " & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
    Call LogError("Error " & Err.number & " (Linea: " & Erl & ") " & Err.description & " en User: " & UserList(CasperIndex).Name & " con IP: " & UserList(CasperIndex).IpAddress & " con pos " & UserList(CasperIndex).Pos.Map & " X:" & UserList(CasperIndex).Pos.X & " Y:" & UserList(CasperIndex).Pos.Y)

    '</EhFooter>
End Function

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

    '<EhHeader>
    On Error GoTo InvertHeading_Err

    '</EhHeader>

    '*************************************************
    'Author: ZaMa
    'Last modified: 30/03/2009
    'Returns the heading opposite to the one passed by val.
    '*************************************************
    Select Case nHeading

        Case eHeading.EAST
            InvertHeading = WEST

        Case eHeading.WEST
            InvertHeading = EAST

        Case eHeading.SOUTH
            InvertHeading = NORTH

        Case eHeading.NORTH
            InvertHeading = SOUTH

    End Select

    '<EhFooter>
    Exit Function

InvertHeading_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.InvertHeading " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ChangeUserInv_Err

    '</EhHeader>

    UserList(UserIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
    '<EhFooter>
    Exit Sub

ChangeUserInv_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ChangeUserInv " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function NextOpenCharIndex() As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo NextOpenCharIndex_Err

    '</EhHeader>

    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS

        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then LastChar = LoopC
            
            Exit Function

        End If

    Next LoopC

    '<EhFooter>
    Exit Function

NextOpenCharIndex_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.NextOpenCharIndex " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub FreeSlot(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 01/10/2012
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo FreeSlot_Err

    '</EhHeader>

    If UserIndex = LastUser Then

        Do While (LastUser > 0)

            If UserList(LastUser).ConnIDValida Then Exit Do
            LastUser = LastUser - 1
        Loop

    End If

    '<EhFooter>
    Exit Sub

FreeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.FreeSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function PonerPuntos(numero As Long) As String
    
    On Error GoTo PonerPuntos_Err

    Dim i     As Integer

    Dim Cifra As String
 
    Cifra = Str(numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)

    For i = 0 To 4

        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

            End If

        Else

            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

            End If

            Exit For

        End If

    Next
 
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
    
    Exit Function

PonerPuntos_Err:
    'Call RegistrarError(err.Number, err.Description, "ModLadder.PonerPuntos", Erl)
    
End Function

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/05/2011 (Amraphen)
    '26/05/2011: Amraphen - Ahora envía la defensa adicional de la armadura de segunda jerarquía
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserStatsTxt_Err

    '</EhHeader>

    Dim GuildI             As Integer

    Dim ModificadorDefensa As Single 'Por las armaduras de segunda jerarquía.

    Dim Ups                As Single

    Dim UpsSTR             As String
    
    With UserList(UserIndex)
        Ups = .Stats.MaxHp - Mod_Balance.getVidaIdeal(.Stats.Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion))
        
        If Ups > 0 Then
            UpsSTR = "+" & Ups
        ElseIf Ups < 0 Then
            UpsSTR = Ups
        Else
            UpsSTR = "promedio"

        End If
        
        Call WriteConsoleMsg(sendIndex, "Personaje: " & .Name & ". " & ListaClases(.Clase) & " " & ListaRazas(.Raza) & " " & UpsSTR, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.Elv & "  EXP: " & PonerPuntos(CLng(.Stats.Exp)) & "/" & PonerPuntos(.Stats.Elu), FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase) & " " & ListaRazas(.Raza), FontTypeNames.FONTTYPE_INFO)
                
        If EsGmPriv(sendIndex) Then

            'Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.Gld & " Eldhires: " & .Stats.Eldhir & " Posición: " & .Pos.X & "," & .Pos.Y & " en mapa " & .Pos.Map, FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & " " & UpsSTR & " Maná: " & .Stats.MinMan & "/" & .Stats.MaxMan & "  Energía: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        End If
              
        Dim Faction As String
        
        If .Faction.Status <> r_None Then
            Faction = InfoFaction(.Faction.Status).Name & " <" & InfoFaction(.Faction.Status).Range(.Faction.Range).Text & ">"
            
            Call WriteConsoleMsg(sendIndex, "Facción: " & Faction, FontTypeNames.FONTTYPE_INFO)

        End If
        
        If .Faction.FragsCri > 0 Then Call WriteConsoleMsg(sendIndex, "Criminales Asesinados: " & .Faction.FragsCri, FontTypeNames.FONTTYPE_INFO)
        If .Faction.FragsCiu > 0 Then Call WriteConsoleMsg(sendIndex, "Ciudadanos Asesinados: " & .Faction.FragsCiu, FontTypeNames.FONTTYPE_INFO)
        If .flags.Traveling = 1 Then Call WriteConsoleMsg(sendIndex, "Tiempo restante para llegar a tu hogar: " & GetHomeArrivalTime(UserIndex) & " segundos.", FontTypeNames.FONTTYPE_INFO)

        If .Counters.TimeTelep > 0 Then Call WriteConsoleMsg(sendIndex, "Tiempo para irte del mapa: " & Int(.Counters.TimeTelep / 60) & " minuto/s", FontTypeNames.FONTTYPE_INFO)
        
        If .Counters.TimeBono > 0 Then Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto gema: " & Int(.Counters.TimeBono / 60) & " minuto/s", FontTypeNames.FONTTYPE_INFO)

        If .Counters.Pena > 0 Then
            Call WriteConsoleMsg(sendIndex, "Tiempo restante para salir en libertad: " & .Counters.Pena & " minuto" & IIf(.Counters.Pena = 1, vbNullString, "s"), FontTypeNames.FONTTYPE_INFOGREEN)

        End If
        
        If .Counters.TimeBonus > 0 Then
            If .Counters.TimeBonus < 60 Then
                Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: " & .Counters.TimeBonus & " segundos.", FontTypeNames.FONTTYPE_INFO)
            ElseIf .Counters.TimeBonus = 60 Then
                Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: 1 minuto.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "Tiempo restante del efecto: " & Int(.Counters.TimeBonus / 60) & " minutos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
            
        Call WriteConsoleMsg(sendIndex, "Puntos de Torneo: " & .Stats.Points, FontTypeNames.FONTTYPE_USERGOLD)
        Call WriteConsoleMsg(sendIndex, "Reputación: " & .Reputacion.promedio & "." & IIf(.Reputacion.promedio < 0, " Paga " & PonerPuntos(5 * Abs(.Reputacion.promedio) * 6) & " Monedas de Oro para ser Ciudadano.", vbNullString), FontTypeNames.FONTTYPE_ANGEL)
         
        If UserIndex = sendIndex Then
            
            If .Account.Premium > 0 Then
                Call WriteConsoleMsg(sendIndex, "Tiempo de Tier: " & .Account.Premium & " restante " & .Account.DatePremium & ".", FontTypeNames.FONTTYPE_USERGOLD)

            End If

        End If
        
    End With

    '<EhFooter>
    Exit Sub

SendUserStatsTxt_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserStatsTxt " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserInvTxt_Err

    '</EhHeader>

    Dim j As Long
    
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To .CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    End With

    '<EhFooter>
    Exit Sub

SendUserInvTxt_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserInvTxt " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserInvTxtFromChar_Err

    '</EhHeader>

    Dim j        As Long

    Dim Charfile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long
    
    Charfile = CharPath & charName & ".chr"
    
    If FileExist(Charfile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(Charfile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(Charfile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

SendUserInvTxtFromChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserInvTxtFromChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserSkillsTxt_Err

    '</EhHeader>

    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, InfoSkill(j).Name & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
    '<EhFooter>
    Exit Sub

SendUserSkillsTxt_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserSkillsTxt " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function EsMascotaCiudadano(ByVal NpcIndex As Integer, _
                                    ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EsMascotaCiudadano_Err

    '</EhHeader>

    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Escriminal(Npclist(NpcIndex).MaestroUser)

        If EsMascotaCiudadano Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(UserIndex).Name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

    '<EhFooter>
    Exit Function

EsMascotaCiudadano_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.EsMascotaCiudadano " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo NPCAtacado_Err

    '</EhHeader>

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
    If Npclist(NpcIndex).Movement <> Estatico And Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        Npclist(NpcIndex).Target = UserIndex
        Npclist(NpcIndex).Hostile = 1
        Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

    End If

    'Npc que estabas atacando.
    Dim LastNpcHit As Integer

    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).Name Then

        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

        End If

    End If
    
    If EsMascotaCiudadano(NpcIndex, UserIndex) Then
        Call VolverCriminal(UserIndex)
        Npclist(NpcIndex).Movement = TipoAI.NpcDefensa
        Npclist(NpcIndex).Hostile = 1
    Else
        EraCriminal = Escriminal(UserIndex)
        
        'Reputacion
        If Npclist(NpcIndex).flags.AIAlineacion = 0 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                Call VolverCriminal(UserIndex)

            End If
        
        ElseIf Npclist(NpcIndex).flags.AIAlineacion = 1 Then
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR / 2

            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP

        End If
        
        If Npclist(NpcIndex).MaestroUser <> UserIndex Then
            'hacemos que el npc se defienda
            Npclist(NpcIndex).Movement = TipoAI.NpcDefensa
            Npclist(NpcIndex).Hostile = 1

        End If
        
        If EraCriminal And Not Escriminal(UserIndex) Then
            Call VolverCiudadano(UserIndex)

        End If
        
        Call AllMascotasAtacanNPC(NpcIndex, UserIndex)

    End If
        
    If UserList(UserIndex).GuildIndex > 0 Then
        If Npclist(NpcIndex).CastleIndex > 0 Then
            Call Castle_Attack(Npclist(NpcIndex).CastleIndex, UserList(UserIndex).GuildIndex)

        End If

    End If
        
    '<EhFooter>
    Exit Sub

NPCAtacado_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.NPCAtacado " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PuedeApuñalar_Err

    '</EhHeader>
    
    Dim WeaponIndex As Integer
     
    With UserList(UserIndex)
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        If WeaponIndex > 0 Then
            If ObjData(WeaponIndex).Apuñala = 1 Then
                PuedeApuñalar = .Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR Or .Clase = eClass.Assasin

            End If

        End If
        
    End With
    
    '<EhFooter>
    Exit Function

PuedeApuñalar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.PuedeApuñalar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PuedeAcuchillar(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/01/2010 (ZaMa)
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PuedeAcuchillar_Err

    '</EhHeader>
    
    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)

        If .Clase = eClass.Thief Then
        
            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex > 0 Then
                PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)

            End If

        End If

    End With
    
    '<EhFooter>
    Exit Function

PuedeAcuchillar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.PuedeAcuchillar " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub SubirSkill(ByVal UserIndex As Integer, _
               ByVal Skill As Integer, _
               ByVal Acerto As Boolean)

    '*************************************************
    'Author: Unknown
    'Last modified: 11/19/2009
    '11/19/2009 Pato - Implement the new system to train the skills.
    '*************************************************
    '<EhHeader>
    On Error GoTo SubirSkill_Err

    '</EhHeader>

    Dim SubeSkill As Boolean
    
    With UserList(UserIndex)

        If .flags.Hambre = 0 And .flags.Sed = 0 Then

            With .Stats

                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                If .UserSkills(Skill) >= LevelSkill(.Elv).LevelValue Then Exit Sub
                      
                If Acerto Then
                    If RandomNumber(1, 100) <= 50 Then SubeSkill = True
                Else

                    If RandomNumber(1, 100) <= 20 Then SubeSkill = True

                End If
                
                If SubeSkill Then
                    .UserSkills(Skill) = .UserSkills(Skill) + 1
                    Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & InfoSkill(Skill).Name & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    
                    .Exp = .Exp + 50

                    If .Exp > MAXEXP Then .Exp = MAXEXP
                    
                    Call WriteConsoleMsg(UserIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    
                    Call WriteUpdateExp(UserIndex)
                    Call CheckUserLevel(UserIndex)

                    'Call CheckEluSkill(UserIndex, Skill, False)
                End If

            End With

        End If

    End With

    '<EhFooter>
    Exit Sub

SubirSkill_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SubirSkill " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Public Sub UserDie(ByVal UserIndex As Integer, _
                   Optional ByVal AttackerIndex As Integer = 0)

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
    On Error GoTo ErrorHandler

    Dim i           As Long

    Dim aN          As Integer
    
    Dim iSoundDeath As Integer
    
    Dim A           As Long
    
    Dim Time        As Long
    
    With UserList(UserIndex)
        
        ' # Masacre en mapas inseguros.
        If MapInfo(.Pos.Map).Pk And .flags.SlotEvent = 0 And .flags.SlotReto = 0 And .flags.SlotFast = 0 Then
            Time = GetTime
            
            If Time - MapInfo(.Pos.Map).DeadTime <= 30000 Then
                MapInfo(.Pos.Map).UsersDead = MapInfo(.Pos.Map).UsersDead + 1
                MapInfo(.Pos.Map).DeadTime = Time
                
                If MapInfo(.Pos.Map).UsersDead > 3 Then
                    WriteMessageDiscord CHANNEL_ONFIRE, "Masacre en **" & MapInfo(.Pos.Map).Name & "**. " & MapInfo(.Pos.Map).UsersDead & " víctimas caídas en menos de 30 segundos. Players: **" & MapInfo(.Pos.Map).NumUsers & "**"

                End If
           
            Else
                MapInfo(.Pos.Map).UsersDead = 0
                MapInfo(.Pos.Map).DeadTime = 0

            End If

        End If
        
        'Sonido
        If .Genero = eGenero.Mujer Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_MUJER

            End If

        Else

            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE

            End If

        End If
        
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, iSoundDeath)
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1
        
        .Counters.Trabajando = 0
        
        ' No se activa en arenas
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

        End If
        
        aN = .flags.AtacadoPorNpc

        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
            Npclist(aN).Target = 0

        End If
        
        aN = .flags.NPCAtacado

        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString

            End If

        End If

        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        Call PerdioNpc(UserIndex, False)
        
        '<<<< Atacable >>>>
        If .flags.AtacablePor > 0 Then
            .flags.AtacablePor = 0
            Call RefreshCharStatus(UserIndex)

        End If
        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
            
        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)

        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)

        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0, .Pos.X, .Pos.Y))

        End If
        
        '<<<< Invisible >>>>
        If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.Invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call SetInvisible(UserIndex, .Char.charindex, False)

        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 1

        End If
        
        If MapInfo(.Pos.Map).CaenItems > 0 Then
            If Not EsGm(UserIndex) Then
                If TieneObjetos(PENDIENTE_SACRIFICIO, 1, UserIndex) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 54, 2))
                    Call QuitarObjetos(PENDIENTE_SACRIFICIO, 1, UserIndex)
                
                Else

                    If MapInfo(.Pos.Map).CaenItems = 1 Then
                        Call TirarTodo(UserIndex)

                    End If

                End If

            End If

        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

        End If
        
        ' Desequipamos la montura
        If .Invent.MonturaObjIndex > 0 Then
            
            Call Desequipar(UserIndex, .Invent.MonturaSlot)
            
            If .flags.Montando Then

                .flags.Montando = False
                Call WriteMontateToggle(UserIndex)

            End If

        End If
        
        ' Desequipamos el pendiente de experiencia
        If .Invent.PendientePartyObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.PendientePartySlot)

        End If
        
        ' Desequipamos la reliquia
        If .Invent.ReliquiaObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ReliquiaSlot)

        End If
        
        ' Desequipamos el Objeto mágico (Laudes y Anillos mágicos)
        If .Invent.MagicObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)

        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

        End If
        
        'desequipar aura
        If .Invent.AuraEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.AuraEqpSlot)

        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

        End If
        
        'desequipar anillo magico/laud
        If .Invent.MagicSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)

        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

        End If

        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Or .flags.Transform = 1 Or .flags.TransformVIP = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
            Next A
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False
            .ShowName = True
            
        End If
        
        ' << Restauramos la transformación
        If .flags.Transform = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex
            Next A
            
            .Counters.TimeTransform = 0
            .flags.Transform = 0
            .flags.Mimetizado = 0

            ' Puede ser atacado por npcs (cuando resucite)
        End If
        
        ' << Restauramos la transformación VIP
        If .flags.TransformVIP = 1 Then
            .Char.Body = .CharMimetizado.Body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = .CharMimetizado.AuraIndex(A)
            Next A

            .flags.TransformVIP = 0
            .flags.Mimetizado = 0
            
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False

        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then

            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i

        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
            .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco

            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = NingunArma
            Next A

            Debug.Print .Char.Head
        Else
            .Char.Body = iFragataFantasmal

        End If

        If .MascotaIndex > 0 Then
            Call MuereNpc(.MascotaIndex, 0)

        End If
        
        ' Chequeos del Poder de las Medusas y lo saco al morir
        If Power.UserIndex = UserIndex Then
            If AttackerIndex > 0 Then
                'If Power.Active Then
                Call Power_Set(AttackerIndex, UserIndex)
                Call Power_Message
                'Else
                'Call Power_Set(0, 0)
                'End If
            Else
                Call Power_Set(0, UserIndex)

            End If

        End If
        
        '<< Actualizamos clientes >>
        Call RefreshCharStatus(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, .Char.AuraIndex)
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(UserIndex)
        
        ' Hay que teletransportar?
        Dim mapa As Integer

        mapa = .Pos.Map

        Dim MapaTelep As Integer

        MapaTelep = MapInfo(mapa).OnDeathGoTo.Map
        
        If MapaTelep <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(UserIndex, MapaTelep, MapInfo(mapa).OnDeathGoTo.X, MapInfo(mapa).OnDeathGoTo.Y, True, True)

        End If
        
        ' Retos
        If .flags.SlotReto Then
            Call mRetos.UserdieFight(UserIndex, AttackerIndex, False)

        End If
        
        ' Desafios
        If .flags.Desafiando > 0 Then
            Desafio_UserKill UserIndex

        End If
        
        ' Retos Rapidos
        If .flags.SlotFast > 0 Then
            RetoFast_UserDie UserIndex

        End If
        
        ' Eventos automáticos
        If .flags.SlotEvent > 0 Then
            Call Events_UserDie(UserIndex, AttackerIndex)

        End If

    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.number & " Descripción: " & Err.description)

End Sub

Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 13/07/2010
    '13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
    '***************************************************
    '<EhHeader>
    On Error GoTo ContarMuerte_Err

    '</EhHeader>

    If EsNewbie(Muerto) Then Exit Sub
        
    With UserList(Atacante)
        'Dim Value As Long
        'Value = CLng(.Stats.Elv - UserList(Muerto).Stats.Elv)
        ' If Abs(Value) > 12 Then Exit Sub
             
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        If AntiFrags_CheckUser(Atacante, Muerto, 1800) = False Then Exit Sub
        
        If Not MapInfo(.Pos.Map).FreeAttack Then
            If Escriminal(Muerto) Then
                If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                    .flags.LastCrimMatado = UserList(Muerto).Name
    
                    If .Faction.FragsCri < MAXUSERMATADOS Then .Faction.FragsCri = .Faction.FragsCri + 1

                End If
    
            Else
    
                If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                    .flags.LastCiudMatado = UserList(Muerto).Name
    
                    If .Faction.FragsCiu < MAXUSERMATADOS Then .Faction.FragsCiu = .Faction.FragsCiu + 1

                End If

            End If

        End If
        
        If .Faction.FragsOther < MAXUSERMATADOS Then .Faction.FragsOther = .Faction.FragsOther + 1
        
        'Call RankUser_AddPoint(Atacante, 1)
    End With

    '<EhFooter>
    Exit Sub

ContarMuerte_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ContarMuerte " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, _
              ByRef nPos As WorldPos, _
              ByRef Obj As Obj, _
              ByRef PuedeAgua As Boolean, _
              ByRef PuedeTierra As Boolean)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/09/2010
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
    '**************************************************************
    On Error GoTo ErrHandler

    Dim Found As Boolean

    Dim LoopC As Integer

    Dim tX    As Long

    Dim tY    As Long
    
    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, True) Then
        
        If Not HayObjeto(Pos.Map, nPos.X, nPos.Y, Obj.ObjIndex, Obj.Amount) Then
            Found = True

        End If
        
    End If
    
    ' Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then

        While (Not Found) And LoopC <= 16

            If RhombLegalTilePos(Pos, tX, tY, LoopC, Obj.ObjIndex, Obj.Amount, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True

            End If
        
            LoopC = LoopC + 1

        Wend
        
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0

    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.number & " - " & Err.description)

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 ByVal FX As Boolean, _
                 Optional ByVal Teletransported As Boolean)

    '<EhHeader>
    On Error GoTo WarpUserChar_Err

    '</EhHeader>

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 11/23/2010
    '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
    '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
    '16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
    '11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
    '**************************************************************
    Dim OldMap As Integer

    Dim OldX   As Integer

    Dim OldY   As Integer

    Dim nPos   As WorldPos
    
    If Map = 0 Or X = 0 Or Y = 0 Then
        Call LogError("Cuenta " & UserList(UserIndex).Account.Email & " NICK: " & UserList(UserIndex).Name & "  Map " & Map & " X: " & X & " Y: " & Y)
        Exit Sub

    End If
          
    With UserList(UserIndex)

        'Quitar el dialogo solo si no es GM.
        If .flags.AdminInvisible = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))

        End If
      
        OldMap = .Pos.Map
        OldX = .Pos.X
        OldY = .Pos.Y
              
        ' If OldMap <> Map Then

        If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)
                
            Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)

            Dim WasInvi      As Boolean

            'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
            If MapInfo(Map).InviSinEfecto > 0 And .flags.Invisible = 1 Then
                .flags.Invisible = 0
                .Counters.Invisibilidad = 0
                AhoraVisible = True
                WasInvi = True 'si era invi, para el string

            End If

            'Chequeo de flags de mapa por ocultar (C4b3z0n)
            If MapInfo(Map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                AhoraVisible = True
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0

            End If
                
            ' Chequeo de flags de gran poder
            If Power.UserIndex = UserIndex Then
                If Not MapInfo(Map).Poder = 1 Then
                    Call Power_Set(0, UserIndex)

                End If

            End If
                
            'Chequeo de Mimetismo de mapa
            If MapInfo(Map).MimetismoSinEfecto > 0 And .flags.Mimetizado = 1 Then
                Call Mimetismo_Reset(UserIndex)

            End If

            If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
                Call SetInvisible(UserIndex, .Char.charindex, False)

                If WasInvi Then 'era invi
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                Else 'estaba oculto
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
            
        Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)

        Call WritePlayMusic(UserIndex, val(ReadField(1, MapInfo(Map).Music, 45)))

        Call WriteChangeMap(UserIndex, Map)
        'Call WritePosUpdate(UserIndex)
                
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        MapInfo(OldMap).Players.Remove UserIndex
            
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        MapInfo(Map).Players.Add UserIndex

        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0

        End If
        
        'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
        Dim nextMap, previousMap As Boolean

        nextMap = IIf(distanceToCities(Map).distanceToCity(1) >= 0, True, False)
        previousMap = IIf(distanceToCities(.Pos.Map).distanceToCity(1) >= 0, True, False)

        If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
            'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
        ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
            .flags.LastMap = .Pos.Map
        ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el último mapa es 0 ya que no esta en un dungeon)
            .flags.LastMap = 0
        ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
            .flags.LastMap = .flags.LastMap

        End If

        If .flags.Privilegios = PlayerType.User Or .flags.Privilegios = PlayerType.RoyalCouncil Or .flags.Privilegios = PlayerType.ChaosCouncil Then
            Call WriteRemoveAllDialogs(UserIndex)

        End If
            
        '  Else
        '     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
        '   MapData(.Pos.Map, X, Y).UserIndex = UserIndex
        'End If

        .Pos.X = X
        .Pos.Y = Y
        .Pos.Map = Map
                
        'If OldMap <> Map Then

        Call MakeUserChar(True, Map, UserIndex, Map, X, Y)
        Call WriteUserCharIndexInServer(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterAttackMovement(UserList(UserIndex).Char.charindex), , True)
        'Actualizamos las áreas de ser necesario
        '   Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)
        '  End If
          
        ' // NUEVO
        If MapData(Map, X, Y).TileExit.Map > 0 Then
            Call DoTileEvents(UserIndex, Map, X, Y)

        End If

        'Seguis invisible al pasar de mapa
        If (.flags.Invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            
            ' No si estas navegando
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.charindex, True)

            End If

        End If

        If Teletransported Then
            If .flags.Traveling = 1 Then
                Call EndTravel(UserIndex, True)

            End If

        End If
        
        If FX And .flags.AdminInvisible = 0 And Not EsAdmin(.Name) Then  'FX
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, X, Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FXWARP, 0))

        End If
        
        If .MascotaIndex Then
            Call QuitarPet(UserIndex, .MascotaIndex)

            'Call WarpMascota_Map(UserIndex)
        End If

        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(UserIndex, True)
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(UserIndex, False)

        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User)) = 0 Then
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)

                End If

            Else

                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(UserIndex)

                End If

            End If

        End If

        ' Checking Event Teleports
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.Teleports Then
                If .Pos.Map = MapEvent.TeleportWin.Map And .Pos.X = MapEvent.TeleportWin.X And .Pos.Y = MapEvent.TeleportWin.Y Then
                    Events_Teleports_Finish UserIndex

                End If

            End If

        End If
        
    End With
    
    Exit Sub

WarpUserChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.WarpUserChar (Map: " & Map & " X: " & X & " Y: " & Y & ")" & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub WarpMascota_Map(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo WarpMascota_Map_Err

    '</EhHeader>

    '************************************************
    'Author: Uknown
    'Last Modified: 26/10/2010
    '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
    '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
    '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
    '26/10/2010: ZaMa - Ahora las mascotas rapswnean de forma aleatoria.
    '************************************************

    Dim PetTiempoDeVida As Integer

    Dim canWarp         As Boolean

    Dim Index           As Integer

    Dim iMinHP          As Integer
    
    Dim NpcNumber       As Integer

    With UserList(UserIndex)
        canWarp = (MapInfo(.Pos.Map).Pk = True)
        
        If .MascotaIndex And canWarp Then
            iMinHP = Npclist(.MascotaIndex).Stats.MinHp
            PetTiempoDeVida = Npclist(.MascotaIndex).Contadores.TiempoExistencia
            NpcNumber = Npclist(.MascotaIndex).numero
            
            Call QuitarNPC(.MascotaIndex)
            .MascotaIndex = 0
            
            Dim SpawnPos As WorldPos
        
            SpawnPos.Map = .Pos.Map
            SpawnPos.X = .Pos.X + RandomNumber(-3, 3)
            SpawnPos.Y = .Pos.Y + RandomNumber(-3, 3)
        
            'Index = SpawnNpc(NpcNumber, SpawnPos, False, False)
            Index = CrearNPC(NpcNumber, SpawnPos.Map, SpawnPos)
            
            If Index = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tu mascota no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                .MascotaIndex = Index

                ' Nos aseguramos de que conserve el hp, si estaba dañado
                Npclist(Index).Stats.MinHp = iMinHP
            
                Npclist(Index).MaestroUser = UserIndex
                Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(Index)

            End If
            
        End If

    End With
    
    If Not canWarp Then
        Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)

    End If
    
    '<EhFooter>
    Exit Sub

WarpMascota_Map_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.WarpMascota_Map " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Cerrar_Usuario_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/09/2010
    '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
    '***************************************************
    
    With UserList(UserIndex)

        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf(((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.Map).Pk), IntervaloCerrarConexion, 0)

            Call WriteConsoleMsg(UserIndex, "Cerrando...Se cerrará el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

Cerrar_Usuario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.Cerrar_Usuario " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo CancelExit_Err

    '</EhHeader>

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/08
    '
    '***************************************************
    If UserList(UserIndex).Counters.Saliendo Then

        ' Is the user still connected?
        If UserList(UserIndex).flags.UserLogged Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Pk, IntervaloCerrarConexion, 0)

        End If

    End If
        
    'If Teleports create, cancel
    Call Teleports_Cancel(UserIndex)
    '<EhFooter>
    Exit Sub

CancelExit_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CancelExit " & "at line " & Erl
        
    '</EhFooter>
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, _
                       ByVal UserIndexDestino As Integer, _
                       ByVal NuevoNick As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CambiarNick_Err

    '</EhHeader>

    Dim ViejoNick       As String

    Dim ViejoCharBackup As String
    
    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).Name
    
    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

    End If

    '<EhFooter>
    Exit Sub

CambiarNick_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CambiarNick " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserStatsTxtOFF_Err

    '</EhHeader>

    Dim Ups    As Single

    Dim Elv    As Long, MaxHp As Long, Clase As eClass, Raza As eRaza, Constitucion As Byte
    
    Dim Bronce As Byte, Plata As Byte, Oro As Byte, Premium As Byte
    
    If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadísticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        
        Elv = val(GetVar(CharPath & Nombre & ".chr", "STATS", "ELV"))
        MaxHp = val(GetVar(CharPath & Nombre & ".chr", "STATS", "MAXHP"))
        Clase = val(GetVar(CharPath & Nombre & ".chr", "INIT", "CLASE"))
        Raza = val(GetVar(CharPath & Nombre & ".chr", "INIT", "RAZA"))
        Constitucion = val(GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT" & eAtributos.Constitucion))
        Ups = MaxHp - Mod_Balance.getVidaIdeal(Elv, Clase, Constitucion)
         
        'Bronce = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "Bronce"))
        'Plata = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "PLATA"))
        'Oro = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "ORO"))
        'Premium = val(GetVar(CharPath & Nombre & ".chr", "FLAGS", "PREMIUM"))
        
        Call WriteConsoleMsg(sendIndex, "Nivel: " & Elv & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase-Raza: " & ListaClases(Clase) & " " & ListaRazas(Raza), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, IIf(Bronce > 0, "BRONCE: SI. ", "BRONCE: NO. ") & IIf(Plata > 0, "PLATA: SI. ", "PLATA: NO. ") & IIf(Premium > 0, "PREMIUM: SI. ", "PREMIUM: NO. ") & Oro, FontTypeNames.FONTTYPE_INFO)
        
        'Call WriteConsoleMsg(sendIndex, "Energía: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & MaxHp & " Ups: " & Ups & ",  Maná: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dsp: " & GetVar(CharPath & Nombre & ".chr", "stats", "ELDHIR"), FontTypeNames.FONTTYPE_INFO)
        
        #If ConUpTime Then

            Dim TempSecs As Long

            Dim TempSTR  As String

            TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
            TempSTR = (TempSecs \ 86400) & " Días, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempSTR, FontTypeNames.FONTTYPE_INFO)
        #End If
    
        'Call WriteConsoleMsg(sendIndex, "Dados: " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT1") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT2") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT3") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT4") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT5"), FontTypeNames.FONTTYPE_INFO)
    End If

    '<EhFooter>
    Exit Sub

SendUserStatsTxtOFF_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserStatsTxtOFF " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SendUserOROTxtFromChar_Err

    '</EhHeader>

    Dim Charfile As String

    Dim Account  As String

    Charfile = CharPath & charName & ".chr"
    
    If FileExist(Charfile, vbNormal) Then
        Account = GetVar(Charfile, "INIT", "ACCOUNTNAME")
        
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "GLD") & " Monedas de Oro en el banco.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(AccountPath & Account & ACCOUNT_FORMAT, "INIT", "ELDHIR") & " Monedas de Eldhir en el banco.", FontTypeNames.FONTTYPE_INFO)
    
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

SendUserOROTxtFromChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SendUserOROTxtFromChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo VolverCriminal_Err

    '</EhHeader>

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/02/2010
    'Nacho: Actualiza el tag al cliente
    '21/02/2010: ZaMa - Ahora deja de ser atacable si se hace criminal.
    '**************************************************************
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        If MapInfo(.Pos.Map).FreeAttack = True Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User) Then
            .Reputacion.BurguesRep = 0
            .Reputacion.NobleRep = 0
            .Reputacion.PlebeRep = 0
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + vlASALTO

            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            
            If .Faction.Status = r_Armada Then
                Call mFacciones.Faction_RemoveUser(UserIndex)
            Else
                Call Guilds_CheckAlineation(UserIndex, a_Neutral)

            End If
            
            If .flags.AtacablePor > 0 Then .flags.AtacablePor = 0

        End If

    End With
    
    Call RefreshCharStatus(UserIndex)
    '<EhFooter>
    Exit Sub

VolverCriminal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.VolverCriminal " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo VolverCiudadano_Err

    '</EhHeader>

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************
    With UserList(UserIndex)

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        .Reputacion.LadronesRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlASALTO

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        Call Guilds_CheckAlineation(UserIndex, a_Neutral)

    End With
    
    Call RefreshCharStatus(UserIndex)
    '<EhFooter>
    Exit Sub

VolverCiudadano_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.VolverCiudadano " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal Body As Integer) As Boolean

    '<EhHeader>
    On Error GoTo BodyIsBoat_Err

    '</EhHeader>

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 10/07/2008
    'Checks if a given body index is a boat
    '**************************************************************
    'TODO : This should be checked somehow else. This is nasty....
    If Body = iFragataReal Or Body = iFragataCaos Or Body = iBarcaPk Or Body = iGaleraPk Or Body = iGaleonPk Or Body = iBarcaCiuda Or Body = iGaleraCiuda Or Body = iGaleonCiuda Or Body = iFragataFantasmal Then
        BodyIsBoat = True

    End If

    '<EhFooter>
    Exit Function

BodyIsBoat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.BodyIsBoat " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub SetInvisible(ByVal UserIndex As Integer, _
                        ByVal userCharIndex As Integer, _
                        ByVal Invisible As Boolean, _
                        Optional ByVal Intermitencia As Boolean = False)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SetInvisible_Err

    '</EhHeader>

    Dim sndNick As String

    With UserList(UserIndex)
        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, Invisible, Intermitencia))
    
        sndNick = .Name
        
        If Invisible Then
            sndNick = sndNick & " " & TAG_USER_INVISIBLE
            
        Else
            
            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & GuildsInfo(.GuildIndex).Name & ">"

            End If
            
            Call WriteUpdateGlobalCounter(UserIndex, 1, 0)

        End If
    
        Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, UserIndex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))

    End With

    '<EhFooter>
    Exit Sub

SetInvisible_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SetInvisible " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub SetConsulatMode(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/06/10
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo SetConsulatMode_Err

    '</EhHeader>

    Dim sndNick As String

    With UserList(UserIndex)
        sndNick = .Name
    
        If EsGm(UserIndex) Then
            If UCase$(sndNick) <> "LION" Then
                sndNick = sndNick & " " & TAG_GAME_MASTER

            End If

        End If
                    
        If .flags.EnConsulta Then
            sndNick = sndNick & " " & TAG_CONSULT_MODE
        Else

            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & GuildsInfo(.GuildIndex).Name & ">"

            End If

        End If
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChangeNick(.Char.charindex, sndNick))

    End With

    '<EhFooter>
    Exit Sub

SetConsulatMode_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SetConsulatMode " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function IsArena(ByVal UserIndex As Integer) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 10/11/2009
    'Returns true if the user is in an Arena
    '**************************************************************
    '<EhHeader>
    On Error GoTo IsArena_Err

    '</EhHeader>
    IsArena = (TriggerZonaPelea(UserIndex, UserIndex) = TRIGGER6_PERMITE)
    '<EhFooter>
    Exit Function

IsArena_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.IsArena " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub PerdioNpc(ByVal UserIndex As Integer, _
                     Optional ByVal CheckPets As Boolean = True)

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 11/07/2010 (ZaMa)
    'The user loses his owned npc
    '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdió.
    '11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
    '13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
    '**************************************************************
    '<EhHeader>
    On Error GoTo PerdioNpc_Err

    '</EhHeader>

    Dim PetCounter As Long

    Dim PetIndex   As Integer

    Dim NpcIndex   As Integer
    
    With UserList(UserIndex)
        
        NpcIndex = .flags.OwnedNpc

        If NpcIndex > 0 Then
            
            If CheckPets Then
                If .MascotaIndex Then

                    ' Si esta atacando al npc deja de hacerlo
                    If Npclist(.MascotaIndex).TargetNPC = NpcIndex Then
                        Call FollowAmo(.MascotaIndex)

                    End If
                
                End If

            End If
            
            ' Reset flags
            Npclist(NpcIndex).Owner = 0
            .flags.OwnedNpc = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

PerdioNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.PerdioNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ApropioNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 27/07/2010 (zaMa)
    'The user owns a new npc
    '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
    '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
    '27/07/2010: ZaMa - El sistema no aplica a mapas seguros.
    '**************************************************************
    '<EhHeader>
    On Error GoTo ApropioNpc_Err

    '</EhHeader>

    With UserList(UserIndex)

        ' Los admins no se pueden apropiar de npcs
        If EsGm(UserIndex) Then Exit Sub
        
        Dim mapa As Integer

        mapa = .Pos.Map
        
        ' No aplica a triggers seguras
        If MapData(mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No se aplica a mapas seguros
        If MapInfo(mapa).Pk = False Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(mapa).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueño, lo perdio aca
        Npclist(NpcIndex).Owner = UserIndex
        .flags.OwnedNpc = NpcIndex

    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(UserIndex, True)
    '<EhFooter>
    Exit Sub

ApropioNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.ApropioNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function GetDireccion(ByVal UserIndex As Integer, _
                             ByVal OtherUserIndex As Integer) As String

    '<EhHeader>
    On Error GoTo GetDireccion_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve la direccion hacia donde esta el usuario
    '**************************************************************
    Dim X As Integer

    Dim Y As Integer
    
    X = UserList(UserIndex).Pos.X - UserList(OtherUserIndex).Pos.X
    Y = UserList(UserIndex).Pos.Y - UserList(OtherUserIndex).Pos.Y
    
    If X = 0 And Y > 0 Then
        GetDireccion = "Sur"
    ElseIf X = 0 And Y < 0 Then
        GetDireccion = "Norte"
    ElseIf X > 0 And Y = 0 Then
        GetDireccion = "Este"
    ElseIf X < 0 And Y = 0 Then
        GetDireccion = "Oeste"
    ElseIf X > 0 And Y < 0 Then
        GetDireccion = "NorEste"
    ElseIf X < 0 And Y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf X > 0 And Y > 0 Then
        GetDireccion = "SurEste"
    ElseIf X < 0 And Y > 0 Then
        GetDireccion = "SurOeste"

    End If

    '<EhFooter>
    Exit Function

GetDireccion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetDireccion " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function SameFaccion(ByVal UserIndex As Integer, _
                            ByVal OtherUserIndex As Integer) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve True si son de la misma faccion
    '**************************************************************
    '<EhHeader>
    On Error GoTo SameFaccion_Err

    '</EhHeader>
    SameFaccion = (esCaos(UserIndex) And esCaos(OtherUserIndex)) Or (esArmada(UserIndex) And esArmada(OtherUserIndex))
    '<EhFooter>
    Exit Function

SameFaccion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.SameFaccion " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal UserIndex As Integer, _
                         ByVal Skill As Byte, _
                         ByVal Allocation As Boolean)

    '*************************************************
    'Author: Torres Patricio (Pato)
    'Last modified: 11/20/2009
    '
    '*************************************************
    '<EhHeader>
    On Error GoTo CheckEluSkill_Err

    '</EhHeader>

    With UserList(UserIndex).Stats

        If .UserSkills(Skill) < MAXSKILLPOINTS Then
            If Allocation Then
                .ExpSkills(Skill) = 0
            Else
                .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)

            End If
        
            .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
        Else
            .ExpSkills(Skill) = 0
            .EluSkills(Skill) = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

CheckEluSkill_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CheckEluSkill " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function HasEnoughItems(ByVal UserIndex As Integer, _
                               ByVal ObjIndex As Integer, _
                               ByVal Amount As Long) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 25/11/2009
    'Cheks Wether the user has the required amount of items in the inventory or not
    '**************************************************************
    '<EhHeader>
    On Error GoTo HasEnoughItems_Err

    '</EhHeader>

    Dim Slot          As Long

    Dim ItemInvAmount As Long
    
    With UserList(UserIndex)

        For Slot = 1 To .CurrentInventorySlots

            ' Si es el item que busco
            If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                ' Lo sumo a la cantidad total
                ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).Amount

            End If

        Next Slot

    End With
    
    HasEnoughItems = Amount <= ItemInvAmount
    '<EhFooter>
    Exit Function

HasEnoughItems_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.HasEnoughItems " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, _
                                ByVal UserIndex As Integer) As Long

    '<EhHeader>
    On Error GoTo TotalOfferItems_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 25/11/2009
    'Cheks the amount of items the user has in offerSlots.
    '**************************************************************
    Dim Slot As Byte
    
    For Slot = 1 To MAX_OFFER_SLOTS

        ' Si es el item que busco
        If UserList(UserIndex).ComUsu.Objeto(Slot) = ObjIndex Then
            ' Lo sumo a la cantidad total
            TotalOfferItems = TotalOfferItems + UserList(UserIndex).ComUsu.cant(Slot)

        End If

    Next Slot

    '<EhFooter>
    Exit Function

TotalOfferItems_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.TotalOfferItems " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo getMaxInventorySlots_Err

    '</EhHeader>

    If UserList(UserIndex).Invent.MochilaEqpObjIndex > 0 Then
        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(UserList(UserIndex).Invent.MochilaEqpObjIndex).MochilaType * 5 '5=slots por fila, hacer constante
    Else
        getMaxInventorySlots = MAX_NORMAL_INVENTORY_SLOTS

    End If

    '<EhFooter>
    Exit Function

getMaxInventorySlots_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.getMaxInventorySlots " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub goHome(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo (lo saque de tPiquetec)
    '***************************************************
    '<EhHeader>
    On Error GoTo goHome_Err

    '</EhHeader>

    Dim Distance As Long

    Dim Tiempo   As Long
    
    With UserList(UserIndex)
        
        Select Case .Account.Premium

            Case 0
                Tiempo = 120

            Case 1
                Tiempo = 60

            Case 2
                Tiempo = 30

            Case 3
                Tiempo = 5

        End Select
        
        .Counters.goHomeSec = Tiempo
        Call IntervaloGoHome(UserIndex, Tiempo * 1000, True)
                
        ' If .flags.Navegando = 1 Then
        '  .Char.FX = AnimHogarNavegando(.Char.Heading)
        ' Else
        '  .Char.FX = AnimHogar(.Char.Heading)

        '   End If
                
        ' .Char.loops = INFINITE_LOOPS
        ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
                
        Call WriteMultiMessage(UserIndex, eMessages.Home, Distance, Tiempo, , MapInfo(Ciudades(.Hogar).Map).Name)
        
    End With
    
    '<EhFooter>
    Exit Sub

goHome_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.goHome " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub setHome(ByVal UserIndex As Integer, _
                   ByVal newHome As eCiudad, _
                   ByVal NpcIndex As Integer)

    '<EhHeader>
    On Error GoTo setHome_Err

    '</EhHeader>

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
    '01/06/2010: ZaMa - Ahora te avisa si ya tenes ese hogar.
    '***************************************************
    If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
    If newHome = eCiudad.cEsperanza And UserList(UserIndex).Stats.Elv >= 35 Then Exit Sub
          
    If UserList(UserIndex).Hogar <> newHome Then
        UserList(UserIndex).Hogar = newHome
    
        Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a nuestra comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.charindex, vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "¡¡¡Ya eres miembro de nuestra comunidad!!!", Npclist(NpcIndex).Char.charindex, vbWhite)

    End If

    '<EhFooter>
    Exit Sub

setHome_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.setHome " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function GetHomeArrivalTime(ByVal UserIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo GetHomeArrivalTime_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 01/06/2010
    'Calculates the time left to arrive home.
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTime
    
    With UserList(UserIndex)
        GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001

    End With

    '<EhFooter>
    Exit Function

GetHomeArrivalTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.GetHomeArrivalTime " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub HomeArrival(ByVal UserIndex As Integer)

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 01/06/2010
    'Teleports user to its home.
    '**************************************************************
    '<EhHeader>
    On Error GoTo HomeArrival_Err

    '</EhHeader>
    
    Dim tX   As Integer

    Dim tY   As Integer

    Dim tMap As Integer
        
    Dim A    As Long
        
    With UserList(UserIndex)

        'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
        If .flags.Navegando = 1 Then
            .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
            .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
    
            For A = 1 To MAX_AURAS
                .Char.AuraIndex(A) = NingunAura
            Next A
            
            .flags.Navegando = 0
            
            Call WriteNavigateToggle(UserIndex)

            'Le sacamos el navegando, pero no le mostramos a los demás porque va a ser sumoneado hasta ulla.
        End If
        
        tX = Ciudades(.Hogar).X
        tY = Ciudades(.Hogar).Y
        tMap = Ciudades(.Hogar).Map
        
        Call FindLegalPos(UserIndex, tMap, tX, tY)
        Call WarpUserChar(UserIndex, tMap, tX, tY, True)
        
        Call WriteMultiMessage(UserIndex, eMessages.FinishHome)
        
        Call EndTravel(UserIndex, False)
        
    End With
    
    '<EhFooter>
    Exit Sub

HomeArrival_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.HomeArrival " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub EndTravel(ByVal UserIndex As Integer, ByVal Cancelado As Boolean)

    '<EhHeader>
    On Error GoTo EndTravel_Err

    '</EhHeader>

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 11/06/2011
    'Ends travel.
    '**************************************************************
    With UserList(UserIndex)
        .Counters.goHome = 0
        .Counters.goHomeSec = 0
        .flags.Traveling = 0

        If Cancelado Then Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
        .Char.FX = 0
        .Char.loops = 0
        
        Call WriteUpdateGlobalCounter(UserIndex, 4, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))

    End With

    '<EhFooter>
        
    Exit Sub

EndTravel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.EndTravel " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function CharIsValid_Invisibilidad(ByVal UserIndex As Integer, _
                                          ByVal sndIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo CharIsValid_Invisibilidad_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        CharIsValid_Invisibilidad = (UserIndex = sndIndex)
        
        If .Faction.Status <> r_None Then
            CharIsValid_Invisibilidad = (.Faction.Status = UserList(sndIndex).Faction.Status)

        End If
        
        If .GuildIndex > 0 Then
            CharIsValid_Invisibilidad = (.GuildIndex = UserList(sndIndex).GuildIndex)

        End If
       
    End With
    
    '<EhFooter>
    Exit Function

CharIsValid_Invisibilidad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CharIsValid_Invisibilidad " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function CharIs_Admin(ByVal Name As String) As Boolean

    '<EhHeader>
    On Error GoTo CharIs_Admin_Err

    '</EhHeader>
    
    Select Case Name
    
        Case "LION": CharIs_Admin = True

        Case "MELKOR": CharIs_Admin = True
            
        Case "ARAGON": CharIs_Admin = True
            
        Case Else: CharIs_Admin = False
        
    End Select

    '<EhFooter>
    Exit Function

CharIs_Admin_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.UsUaRiOs.CharIs_Admin " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer, _
                                             ByVal ShiftRunner As Boolean) As Single

    On Error GoTo 0
    
    Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single
   
    velocidad = VelocidadNormal

    modificadorItem = 1
    modificadorHechizo = 1
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            
            'velocidad = VelocidadMuerto
            GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad

        End If
        
        ' El traje para nadar es considerado barco, de subtipo = 0
        'If (.flags.Navegando > 0) And (.Invent.BarcoObjIndex > 0) Then
        'modificadorItem = ObjData(.Invent.BarcoObjIndex).velocidad
        'End If
        
        ' If (.flags.Montado = 1) And (.Invent.MonturaObjIndex > 0) Then
        'modificadorItem = ObjData(.Invent.MonturaObjIndex).velocidad
        'End If
        
        ' Algun hechizo le afecto la velocidad
        'If .flags.VelocidadHechizada > 0 Then
        '  modificadorHechizo = .flags.VelocidadHechizada
        'End If
        
        velocidad = VelocidadNormal * modificadorItem * modificadorHechizo
UpdateSpeed:
        .Char.speeding = velocidad
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))
        Call WriteVelocidadToggle(UserIndex)
     
    End With

    Exit Function
    
ActualizarVelocidadDeUsuario_Err:

End Function

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As eHeading)

    On Error GoTo Handler

    With UserList(UserIndex)

        ' Elegimos un lado al azar
        Dim r As Integer

        r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

        ' Roto el heading original hacia ese lado
        Heading = Rotate_Heading(Heading, r)

        ' Intento moverlo para ese lado
        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub

        End If
        
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)

        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub

        End If
        
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As WorldPos

        Call ClosestLegalPos(.Pos, NuevaPos, .flags.Navegando = 1, .flags.Navegando = 0)
        Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, False)

    End With

    Exit Sub
    
Handler:

End Sub
