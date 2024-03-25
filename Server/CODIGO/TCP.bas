Attribute VB_Name = "TCP"
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

Function IBody_Generate(ByVal UserGenero As Byte, ByVal UserRaza As Byte) As Integer

    '<EhHeader>
    On Error GoTo DarCuerpoYCabeza_Err

    '</EhHeader>

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 14/03/2007
    'Elije una cabeza para el usuario y le da un body
    '*************************************************
    Dim NewBody As Integer

    Dim NewHead As Integer
    
    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    NewHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
                    NewBody = HUMANO_H_CUERPO_DESNUDO

                Case eRaza.Elfo
                    NewHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
                    NewBody = ELFO_H_CUERPO_DESNUDO

                Case eRaza.Drow
                    NewHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)
                    NewBody = DROW_H_CUERPO_DESNUDO

                Case eRaza.Enano
                    NewHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
                    NewBody = ENANO_H_CUERPO_DESNUDO

                Case eRaza.Gnomo
                    NewHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)
                    NewBody = GNOMO_H_CUERPO_DESNUDO

            End Select

        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    NewHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)
                    NewBody = HUMANO_M_CUERPO_DESNUDO

                Case eRaza.Elfo
                    NewHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)
                    NewBody = ELFO_M_CUERPO_DESNUDO

                Case eRaza.Drow
                    NewHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)
                    NewBody = DROW_M_CUERPO_DESNUDO

                Case eRaza.Enano
                    NewHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)
                    NewBody = ENANO_M_CUERPO_DESNUDO

                Case eRaza.Gnomo
                    NewHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)
                    NewBody = GNOMO_M_CUERPO_DESNUDO

            End Select

    End Select
    
    IBody_Generate = NewBody
    '<EhFooter>
    Exit Function

DarCuerpoYCabeza_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.DarCuerpoYCabeza " & "at line " & Erl
        
    '</EhFooter>
End Function

Function IHead_Generate(ByVal UserGenero As Byte, ByVal UserRaza As Byte) As Integer

    '<EhHeader>
    On Error GoTo DarCabezaRandom_Err

    '</EhHeader>
    Dim NewBody As Integer

    Dim NewHead As Integer
    
    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    NewHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    NewHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)

                Case eRaza.Drow
                    NewHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)

                Case eRaza.Enano
                    NewHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    NewHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)

            End Select

        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    NewHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    NewHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)

                Case eRaza.Drow
                    NewHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)

                Case eRaza.Enano
                    NewHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    NewHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)

            End Select

    End Select
    
    IHead_Generate = NewHead
    '<EhFooter>
    Exit Function

DarCabezaRandom_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.DarCabezaRandom " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ValidarCabeza(ByVal UserRaza As Byte, _
                       ByVal UserGenero As Byte, _
                       ByVal Head As Integer) As Boolean

    '<EhHeader>
    On Error GoTo ValidarCabeza_Err

    '</EhHeader>

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And Head <= HUMANO_H_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And Head <= ELFO_H_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And Head <= DROW_H_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And Head <= ENANO_H_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And Head <= GNOMO_H_ULTIMA_CABEZA)

            End Select
    
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And Head <= HUMANO_M_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And Head <= ELFO_M_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And Head <= DROW_M_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And Head <= ENANO_M_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And Head <= GNOMO_M_ULTIMA_CABEZA)

            End Select

    End Select
        
    '<EhFooter>
    Exit Function

ValidarCabeza_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ValidarCabeza " & "at line " & Erl
        
    '</EhFooter>
End Function

Function AsciiValidos(ByVal cad As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo AsciiValidos_Err

    '</EhHeader>

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
          
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False

            Exit Function

        End If
          
    Next i

    AsciiValidos = True

    '<EhFooter>
    Exit Function

AsciiValidos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.AsciiValidos " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ValidarNombre(Nombre As String) As Boolean

    '<EhHeader>
    On Error GoTo ValidarNombre_Err

    '</EhHeader>
    
    If Len(Nombre) < ACCOUNT_MIN_CHARACTER_CHAR Or Len(Nombre) > ACCOUNT_MAX_CHARACTER_CHAR Then Exit Function
    
    Dim Temp As String, CantidadEspacios As Byte

    Temp = UCase$(Nombre)
    
    Dim i As Long, Char As Integer, LastChar As Integer

    For i = 1 To Len(Temp)
        Char = Asc(mid$(Temp, i, 1))
        
        If (Char < 65 Or Char > 90) And Char <> 32 Then
            Exit Function
        
        ElseIf Char = 32 Then

            If LastChar = 32 Then
                Exit Function

            End If
                
            CantidadEspacios = CantidadEspacios + 1
                
            If CantidadEspacios > 1 Then
                Exit Function

            End If

        End If
        
        LastChar = Char
    Next

    If Asc(mid$(Temp, 1, 1)) = 32 Or Asc(mid$(Temp, Len(Temp), 1)) = 32 Then
        Exit Function

    End If
    
    ValidarNombre = True

    '<EhFooter>
    Exit Function

ValidarNombre_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ValidarNombre " & "at line " & Erl
        
    '</EhFooter>
End Function

Function Numeric(ByVal cad As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo Numeric_Err

    '</EhHeader>

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
    
        If (car < 48 Or car > 57) Then
            Numeric = False

            Exit Function

        End If
    
    Next i

    Numeric = True

    '<EhFooter>
    Exit Function

Numeric_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.Numeric " & "at line " & Erl
        
    '</EhFooter>
End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo NombrePermitido_Err

    '</EhHeader>

    Dim i As Integer

    For i = 1 To UBound(ForbidenNames)

        If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False

            Exit Function

        End If

    Next i

    NombrePermitido = True

    '<EhFooter>
    Exit Function

NombrePermitido_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.NombrePermitido " & "at line " & Erl
        
    '</EhFooter>
End Function

Function PalabraPermitida(ByVal Texto As String) As Boolean

    ' Realiza una comparación de palabras permitidas, previamente sacamos los espacios.
    ' Pasar textos siempre con lcase$()
    '<EhHeader>
    On Error GoTo PalabraPermitida_Err

    '</EhHeader>
    
    Dim i As Integer
    
    Texto = Replace(Texto, " ", "")

    For i = 1 To UBound(ForbidenText)

        If InStr(Texto, ForbidenText(i)) Then
            PalabraPermitida = False

            Exit Function

        End If

    Next i

    PalabraPermitida = True

    '<EhFooter>
    Exit Function

PalabraPermitida_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.PalabraPermitida " & "at line " & Erl
        
    '</EhFooter>
End Function

Function AsciiValidos_Chat(ByVal cad As String) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo AsciiValidos_Chat_Err

    '</EhHeader>

    Dim car As Byte

    Dim i   As Integer

    cad = LCase$(cad)

    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
          
        If (car = 126) Then
            AsciiValidos_Chat = False

            Exit Function

        End If
          
    Next i

    AsciiValidos_Chat = True

    '<EhFooter>
    Exit Function

AsciiValidos_Chat_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.AsciiValidos_Chat " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ValidateSkills_Err

    '</EhHeader>

    Dim LoopC As Integer

    ' For LoopC = 1 To NUMSKILLS

    '  If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then

    'Exit Function

    '   If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    '  End If

    '  Next LoopC

    ValidateSkills = True
    
    '<EhFooter>
    Exit Function

ValidateSkills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ValidateSkills " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ConnectNewUser(ByVal UserName As String, _
                        ByVal UserClase As eClass, _
                        ByVal UserRaza As eRaza, _
                        ByVal UserSexo As eGenero, _
                        ByVal UserHead As Integer, _
                        ByRef IUser As User) As User

    On Error GoTo ConnectNewUser_Err

    Dim i As Long
        
    Call ResetUserFlags(IUser)
         
    With IUser
        .LastHeading = 0
        .flags.Privilegios = 0
        .flags.TargetX = 0
        .flags.TargetY = 0
        .flags.TargetMap = 0
        
        .Stats.Elv = 1
        .Stats.Exp = 0

        .Stats.Elu = 300
              
        .Stats.Gld = 0
        .Stats.Eldhir = 0

        .Stats.BonusLast = 0
        ReDim .Stats.Bonus(0) As UserBonus
        '.Stats.Retos1Ganados = 0
        '.Stats.DesafiosGanados = 0
        '.Stats.Retos1Jugados = 0
        '.Stats.DesafiosJugados = 0
        '.Stats.TorneosGanados = 0
        '.Stats.TorneosJugados = 0
        .flags.SlotRetoUser = 255
        .flags.Muerto = 0
        .flags.Escondido = 0
    
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.LadronesRep = 0
        .Reputacion.NobleRep = 1000
        .Reputacion.PlebeRep = 30
    
        .Reputacion.promedio = 30 / 6
    
        .Name = UserName
        .Clase = UserClase
        .Raza = UserRaza
        .Genero = UserSexo
        .Hogar = cUllathorpe
        
        ' Dados 18
        .Stats.UserAtributos(eAtributos.Fuerza) = 18 + Balance.ModRaza(UserRaza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = 18 + Balance.ModRaza(UserRaza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = 18 + Balance.ModRaza(UserRaza).Inteligencia
        .Stats.UserAtributos(eAtributos.Carisma) = 18 + Balance.ModRaza(UserRaza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = 18 + Balance.ModRaza(UserRaza).Constitucion
    
        .Char.Heading = eHeading.SOUTH
        
        If .Account.Premium > 0 Then
            .Char.Head = UserHead
        Else
            .Char.Head = IHead_Generate(UserSexo, UserRaza)

        End If
                
        .Char.Body = IBody_Generate(UserSexo, UserRaza)
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        .Char.WeaponAnim = NingunArma
        .OrigChar = .Char
    
        #If ConUpTime Then
            .LogOnTime = Now
            .UpTime = 0
        #End If
        
    End With
        
    ConnectNewUser = IUser
        
    '<EhFooter>
    Exit Function

ConnectNewUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ConnectNewUser " & "at line " & Erl

    '</EhFooter>
End Function

Sub CloseSocket(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CloseSocket_Err

    '</EhHeader>

    Dim isNotVisible As Boolean

    Dim HiddenPirat  As Boolean
    
    With UserList(UserIndex)
        isNotVisible = (.flags.Oculto Or .flags.Invisible)

        If isNotVisible Then
            .flags.Invisible = 0
            .flags.Oculto = 0
                
            ' Para no repetir mensajes
            Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
            ' Si esta navegando ya esta visible
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.charindex, False)

            End If

        End If
            
        If .flags.Traveling = 1 Then
            Call EndTravel(UserIndex, True)

        End If
        
        'mato los comercios seguros
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(.ComUsu.DestUsu)

                End If

            End If

        End If
        
        ' Eventos automáticos
        Dim SlotEvent As Byte, SlotUser As Byte, TeamUser As Byte, MapFight As Byte

        SlotEvent = .flags.SlotEvent
        SlotUser = .flags.SlotUserEvent
        
        If SlotEvent > 0 Then
            TeamUser = Events(SlotEvent).Users(SlotUser).Team
            MapFight = Events(SlotEvent).Users(SlotUser).MapFight
            
            Call AbandonateEvent(UserIndex, , True)
            
            ' Si no empezo no tiene sentido comprobar esto. Es para buscar ganador
            If Events(SlotEvent).Run Then Call Events_CheckInscribed(UserIndex, SlotEvent, SlotUser, TeamUser, MapFight)

        End If
              
        ' Retos entre personajes
        If .flags.SlotReto > 0 Then
            Call mRetos.UserdieFight(UserIndex, 0, True)

        End If
        
        If .flags.Desafiando > 0 Then
            Desafio_UserKill UserIndex

        End If
        
        If .flags.SlotFast > 0 Then
            RetoFast_UserDie UserIndex, True

        End If
        
        If .flags.Transform Then
            Call Transform_User(UserIndex, 0)

        End If
        
        If .flags.TransformVIP Then
            Call TransformVIP_User(UserIndex, 0)

        End If
        
        If .flags.ClainObject = 1 Then
            Call mRetos.Retos_ReclameObj(UserIndex)

        End If
        
        If Power.UserIndex = UserIndex Then
            Call Power_Set(0, UserIndex)

        End If
            
        If .flags.BotList > 0 Then
            Call Streamer_SetBotList(UserIndex, .flags.BotList, True)

        End If
            
        Call Teleports_Cancel(UserIndex)
           
        If Not EsGm(UserIndex) Then
            Call WriteUpdateUserData(UserList(UserIndex))

        End If
            
        If .Pos.Map > 0 Then
                
            If .GuildIndex > 0 Then
                GuildsInfo(.GuildIndex).Members(.GuildSlot).UserIndex = 0
                .GuildSlot = 0
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageConsoleMsg("El personaje " & .Name & " se ha desconectado.", FontTypeNames.FONTTYPE_GUILDMSG))
            
            End If
            
            If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
                Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)

            End If
            
        End If

        If StreamerBot.Active = UserIndex Then
            Call Streamer_Initial(0, 0, 0, 0)

        End If
              
        'Call Streamer_CheckUser(UserIndex)
              
        If .flags.UserLogged Then
            Call CloseUser(UserIndex)

            '  #If Classic = 0 Then
            ' Battle_Arenas(.ServerSelected).Users = Battle_Arenas(.ServerSelected).Users - 1
            '   .ServerSelected = 0
            '   WriteLoggedAccountBattle UserIndex
            '  #End If
              
        End If
        
        Call ResetUserSlot(UserIndex)

    End With
   
    '<EhFooter>
    Exit Sub

CloseSocket_Err:
    Call ResetUserSlot(UserIndex)

    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.CloseSocket " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean

    '<EhHeader>
    On Error GoTo EstaPCarea_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim X As Integer, Y As Integer

    For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True

                Exit Function

            End If
        
        Next X
    Next Y

    EstaPCarea = False
    '<EhFooter>
    Exit Function

EstaPCarea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.EstaPCarea " & "at line " & Erl
        
    '</EhFooter>
End Function

Function HayPCarea(Pos As WorldPos) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo HayPCarea_Err

    '</EhHeader>

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True

                    Exit Function

                End If

            End If

        Next X
    Next Y

    HayPCarea = False
    '<EhFooter>
    Exit Function

HayPCarea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.HayPCarea " & "at line " & Erl
        
    '</EhFooter>
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo HayOBJarea_Err

    '</EhHeader>

    Dim X As Integer, Y As Integer

    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True

                Exit Function

            End If
        
        Next X
    Next Y

    HayOBJarea = False
    '<EhFooter>
    Exit Function

HayOBJarea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.HayOBJarea " & "at line " & Erl
        
    '</EhFooter>
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ValidateChr_Err

    '</EhHeader>

    ValidateChr = UserList(UserIndex).Char.Head <> 0 And UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

    '<EhFooter>
    Exit Function

ValidateChr_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ValidateChr " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function CheckPenas(ByVal UserIndex As Integer, ByVal Name As String) As Boolean

    '<EhHeader>
    On Error GoTo CheckPenas_Err

    '</EhHeader>
 
    Dim tStr As String
    
    If val(GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BAN")) > 0 Then
            
        tStr = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "DATEDAY")
            
        If tStr <> vbNullString Then
            If Format(Now, "dd/mm/yyyy") > tStr Then
                Call UnBan(UCase$(Name))

            End If
                
        Else

            Dim Razon As String

            Dim Pena  As String: Pena = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "CANT")

            Razon = GetVar(CharPath & UCase$(Name) & ".chr", "PENAS", "P" & Pena)
            Call WriteErrorMsg(UserIndex, "Tu personaje no tiene permitido ingresar al juego. RAZON: " & Razon)

            Exit Function

        End If
           
    End If
            
    CheckPenas = True
    '<EhFooter>
    Exit Function

CheckPenas_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.CheckPenas " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub UpdatePremium(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo UpdatePremium_Err

    '</EhHeader>

    Dim TimerPremium As String

    TimerPremium = UserList(UserIndex).Account.DatePremium

    If TimerPremium <> vbNullString Then
        If DateDiff("s", Now, TimerPremium) <= 0 Then
            ' If Format(Now, "dd/mm/aa hh:mm:ss") > TimerPremium Then
            UserList(UserIndex).Account.DatePremium = vbNullString
            UserList(UserIndex).Account.Premium = 0
            Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "INIT", "DATEPREMIUM", vbNullString)
            Call WriteVar(AccountPath & UserList(UserIndex).Account.Email & ".acc", "INIT", "PREMIUM", "0")
            
            Call WriteConsoleMsg(UserIndex, "¡El PREMIUM se ha ido de tu cuenta!", FontTypeNames.FONTTYPE_INFORED)
            
        Else
            Call WriteConsoleMsg(UserIndex, "Tu cuenta PREMIUM vence " & UserList(UserIndex).Account.DatePremium & ".", FontTypeNames.FONTTYPE_INFOGREEN)

        End If

    End If

    '<EhFooter>
    Exit Sub

UpdatePremium_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.UpdatePremium " & "at line " & Erl

    '</EhFooter>
End Sub

' # Chequea si ya pasó la hora y son llevados a donde empieza la versión
Public Sub User_Go_Initial_Version()

    If Not DateDiff("s", Now, DateAperture) <= 0 Then Exit Sub
    
    Dim A   As Long

    Dim Pos As WorldPos
    
    For A = 1 To LastUser

        With UserList(A)

            If .Pos.Map = CiudadFlotante.Map Then
                Pos.Map = Newbie.Map
                Pos.X = RandomNumber(Newbie.X - 5, Newbie.X + 5)
                Pos.Y = RandomNumber(Newbie.Y - 5, Newbie.Y + 5)
                
                Call EventWarpUser(A, Pos.Map, Pos.X, Pos.Y)
                Call SendData(SendTarget.ToOne, A, PrepareMessagePlayEffect(eSound.sVictory4, Pos.X, Pos.Y))
                 
                Call WriteConsoleMsg(A, "¡La espera terminó! A entrenar y disfrutar de una nueva versión", FontTypeNames.FONTTYPE_DESAFIOS)
                
            End If

        End With

    Next A
   
End Sub

' # Chequea el momento en el que logea si está en previa de apertura o ya comenzó la versión
Public Function User_Check_Login_Apertura(ByVal UserIndex As Integer) As WorldPos

    Dim Pos As WorldPos
    
    With UserList(UserIndex)
        
        If .Pos.Map = 0 Then
            ' # Ya comenzó la versión
            
            If DateDiff("s", Now, DateAperture) <= 0 Then
                Pos.Map = Newbie.Map
                Pos.Y = Newbie.Y
                Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)
            Else
                ' # Sum en ciudad flotante
                Pos.Map = CiudadFlotante.Map
                Pos.Y = CiudadFlotante.Y
                Pos.X = RandomNumber(CiudadFlotante.X - 3, CiudadFlotante.X + 3)

            End If
            
        Else

            ' El personaje deslogeo antes de tiempo y quedo en la flotante
            If .Pos.Map = CiudadFlotante.Map Then

                ' # Ya comenzó la versión. Lo llevamos al dungeon newbie
                If DateDiff("s", Now, DateAperture) <= 0 Then
                    Pos.Map = Newbie.Map
                    Pos.Y = Newbie.Y
                    Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)

                End If

            End If
            
        End If

    End With
    
    User_Check_Login_Apertura = Pos

End Function

Sub ConnectUser(ByVal UserIndex As Integer, _
                ByRef Name As String, _
                ByVal NewChar As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 24/07/2010 (ZaMa)
    '26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
    '12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
    '14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
    '11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
    '03/12/2009: Budi - Optimización del código
    '24/07/2010: ZaMa - La posicion de comienzo es namehuak, como se habia definido inicialmente.
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim N     As Integer
        
    Dim A     As Long

    Dim tStr  As String

    Dim Valid As Boolean

    With UserList(UserIndex)
    
        Dim i As Long
         
        Call ResetUserFlags(UserList(UserIndex))
        
        If Not CheckPenas(UserIndex, Name) Then Exit Sub

        '¿Ya esta conectado el personaje?
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
            Else
                Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")

            End If

            Exit Sub

        End If
            
        If val(GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BLOCKED")) > 0 Then

            Dim TempOfferTime As String

            TempOfferTime = GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "OFFERTIME")
                  
            If TempOfferTime <> vbNullString Then
                If Format(Now, "dd/mm/aa hh:mm:ss") > TempOfferTime Then
                    Call WriteVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "OFFERTIME", "")
                    Call WriteVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "BLOCKED", "0")
                Else
                    Call WriteErrorMsg(UserIndex, "Tu personaje ha sido ofrecido en MODO CANDADO. Esto significa que podrás entrar: " & TempOfferTime)

                End If
                         
            Else
                Call WriteErrorMsg(UserIndex, "El personaje está bloqueado ya que está en el mercado central. Deberás quitarlo de la misma para poder ingresar.")
                Exit Sub

            End If

        End If
            
        'Reseteamos los FLAGS
        .UserKey = 0
        .UserLastClick = 0
        .UserLastClick_Tolerance = 0
    
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Char.FX = 0
        .flags.MenuCliente = 255
        .flags.LastSlotClient = 255
            
        'Reseteamos los privilegios
        .flags.Privilegios = 0
    
        'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
        If EsAdmin(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        
            .flags.PrivEspecial = EsGmEspecial(Name)
        
            Call Logs_User(Name, eLog.eGm, eLogDescUser.eNone, "Se conecto con ip:" & .IpAddress)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True

        End If
        
        If ServerSoloGMs > 0 Then
            If Not Email_Is_Testing_Pro(.Account.Email) Then
                Call Protocol.Kick(UserIndex, "Servidor en mantenimiento. Consulta otros servidores para disfrutar y pasar el rato.")
        
                Exit Sub
        
            End If

        End If

        'Cargamos el personaje
        Dim Leer As clsIniManager

        Set Leer = New clsIniManager

        Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
        
        ' Cargamos la reputación antes para generar algunos cambios sobre los flags
        Call LoadUserReputacion(UserIndex, Leer)
        
        'Cargamos los datos del personaje
        Call LoadUserInit(UserIndex, Leer)

        Call LoadUserStats(UserIndex, Leer)

        Call LoadQuestStats(UserIndex, Leer)

        Call LoadUserAntiFrags(UserIndex, Leer)

        'Cargamos los mensajes privados del usuario.
        Call CargarMensajes(UserIndex, Leer)
    
        If Not ValidateChr(UserIndex) Then
            Call Protocol.Kick(UserIndex, "Error en el personaje.")

            Exit Sub

        End If
        
        Call LoadUserMeditations(UserIndex, Leer)

        Set Leer = Nothing
              
        If .Invent.ArmourEqpObjIndex > 0 Then .Char.AuraIndex(1) = ObjData(.Invent.ArmourEqpObjIndex).AuraIndex(1)
        If .Invent.WeaponEqpObjIndex > 0 Then .Char.AuraIndex(2) = ObjData(.Invent.WeaponEqpObjIndex).AuraIndex(2)
        If .Invent.CascoEqpObjIndex > 0 Then .Char.AuraIndex(3) = ObjData(.Invent.CascoEqpObjIndex).AuraIndex(3)
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.AuraIndex(4) = ObjData(.Invent.EscudoEqpObjIndex).AuraIndex(4)
        
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
              
        If .Invent.MochilaEqpSlot > 0 Then
            .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).ObjIndex).MochilaType * 5
        Else
            .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS

        End If

        .flags.DragBlocked = False

        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)

        If .flags.Paralizado Then
            Call WriteParalizeOK(UserIndex)

        End If

        Dim mapa          As Integer

        Dim MessageNewbie As String
        
        mapa = .Pos.Map
        
        Dim TempPos As WorldPos

        TempPos = User_Check_Login_Apertura(UserIndex)
        
        If TempPos.X <> 0 Then
            .Pos = TempPos

        End If
        
        'Posicion de comienzo
        If mapa = 0 Then
            mapa = .Pos.Map
            
        Else

            ' El personaje deslogeo antes de tiempo y quedo en la flotante
            If mapa = CiudadFlotante.Map Then

                ' # Ya comenzó la versión. Lo llevamos al dungeon newbie
                If DateDiff("s", Now, DateAperture) <= 0 Then
                    .Pos.Map = Newbie.Map
                    .Pos.Y = Newbie.Y
                    .Pos.X = RandomNumber(Newbie.X - 3, Newbie.X + 1)

                End If

            End If
            
            If Not MapaValido(mapa) Then
                Call Protocol.Kick(UserIndex, "El PJ se encuenta en un mapa inválido.")

                Exit Sub

            End If
        
            ' If map has different initial coords, update it
            Dim StartMap As Integer

            StartMap = MapInfo(mapa).StartPos.Map

            If StartMap <> 0 Then
                If MapaValido(StartMap) Then
                    .Pos = MapInfo(mapa).StartPos
                    mapa = StartMap

                End If

            End If
        
        End If
    
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
        If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

            Dim FoundPlace As Boolean

            Dim esAgua     As Boolean

            Dim tX         As Long

            Dim tY         As Long
        
            FoundPlace = False
            esAgua = HayAgua(mapa, .Pos.X, .Pos.Y)
        
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1

                    If esAgua Then

                        'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                        If LegalPos(mapa, tX, tY, True, False, True) Then
                            FoundPlace = True

                            Exit For

                        End If

                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                        If LegalPos(mapa, tX, tY, False, True, True) Then
                            FoundPlace = True

                            Exit For

                        End If

                    End If

                Next tX
            
                If FoundPlace Then Exit For
            Next tY
        
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.X = tX
                .Pos.Y = tY
            Else

                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                If MapData(mapa, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                    If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                        If UserList(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                            Call FlushBuffer(UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)

                        End If
                    
                        'Lo sacamos.
                        If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)

                        End If

                    End If
                
                    If UserList(MapData(mapa, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                        Call WriteErrorMsg(MapData(mapa, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)

                    End If
                
                    'Call CloseSocket(MapData(Mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call WriteDisconnect(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call FlushBuffer(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)
                    Call CloseSocket(MapData(mapa, .Pos.X, .Pos.Y).UserIndex)

                End If

            End If

        End If

        'Nombre de sistema
        .Name = Name
        .secName = .Name
    
        .ShowName = True 'Por default los nombres son visibles
    
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And (HayAgua(mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.Body)) Then

            .Char.Head = 0

            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(UserIndex)
            Else
                .Char.Body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = NingunAura
                Next A

            End If
        
            .flags.Navegando = 1

        End If
    
        'Info
        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
        Call WritePlayMusic(UserIndex, val(ReadField(1, MapInfo(.Pos.Map).Music, 45)))

        If .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(250, 250, 150)
        ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 0)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 255)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
            .flags.ChatColor = RGB(255, 128, 64)
        Else
            .flags.ChatColor = vbWhite

        End If

        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
    
        'Crea  el personaje del usuario
        If Not MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y) Then
            Exit Sub

        End If

        If (.flags.Privilegios And (PlayerType.User)) = 0 Then
            Call DoAdminInvisible(UserIndex)
            .flags.SendDenounces = True
        Else

            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

            End If

        End If

        Call WriteUserCharIndexInServer(UserIndex)
        Call ActualizarVelocidadDeUsuario(UserIndex, False)
        ''[/el oso]

        ' // NUEVO
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then
            Call DoTileEvents(UserIndex, .Pos.Map, .Pos.X, .Pos.Y)

        End If
        
        Call CheckUserLevel(UserIndex)
        Call WriteUpdateUserStats(UserIndex)
    
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateStrenghtAndDexterity(UserIndex)

        'Call SendMOTD(UserIndex)

        If haciendoBK Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)

        End If

        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER, eMessageType.Admin)

        End If

        If EnTesting Then
            Call WriteErrorMsg(UserIndex, "Servidor en Testeo. Espere unos momentos y consulte la página oficial. WWW.ARGENTUMGAME.COM")

            Exit Sub

        End If

        If TieneMensajesNuevos(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "¡Tienes mensajes privados sin leer!", FontTypeNames.FONTTYPE_FIGHT)

        End If

        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
        
        .flags.UserLogged = True

        'Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        MapInfo(.Pos.Map).Players.Add UserIndex
     
        If .Stats.SkillPts > 0 Then
            Call WriteLevelUp(UserIndex, .Stats.SkillPts)

        End If

        If NumUsers + UsersBot > RECORDusuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Seguimos sumando jugadores a nuestra comunidad!." & " Hay " & NumUsers + UsersBot & " usuarios conectados. Gracias por Jugar.", FontTypeNames.FONTTYPE_INFO))
            RECORDusuarios = NumUsers + UsersBot
            Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(RECORDusuarios))
        
            'Call EstadisticasWeb.Informar(RECORD_USUARIOS, RECORDusuarios)
        End If

        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(UserIndex)

        End If

        If .flags.Montando = 1 Then
            Call WriteMontateToggle(UserIndex)

        End If
        
        'If .flags.Muerto = 1 Then
        Call WriteUpdateUserDead(UserIndex, .flags.Muerto)
        
        'End If
        
        Call WriteConsoleMsg(UserIndex, "Desterium Online. Un servidor de Argentum Online.", FontTypeNames.FONTTYPE_CONSEJOVesA)
        ' Call WriteConsoleMsg(UserIndex, "Utiliza el comando /AYUDA. ¡Te dirá todo lo que necesitas saber para comenzar! Recuerda que desde la página principal podrás acceder a soporte 24/7", FontTypeNames.FONTTYPE_USERGOLD)
        
        If HappyHour Then
            Call WriteConsoleMsg(UserIndex, "¡HappyHour Activado! Exp x2 ¡Entrená tu personaje!", FontTypeNames.FONTTYPE_USERBRONCE)

        End If
            
        If PartyTime Then
            Call WriteConsoleMsg(UserIndex, "PartyTime» Los miembros de la party reciben 25% de experiencia extra.", FontTypeNames.FONTTYPE_INVASION)

        End If
            
        'Call WriteConsoleMsg(UserIndex, "MANUAL: WWW.ARGENTUMGAME.COM/wiki/", FontTypeNames.FONTTYPE_USERBRONCE)
        Call WriteConsoleMsg(UserIndex, "Cualquier acto considerado dañino para la comunidad y/o usuarios miembros de la misma retornará en un bloqueo de cuenta y personajes.", FontTypeNames.FONTTYPE_USERBRONCE)
        
        If MessageNewbie <> vbNullString Then
            Call WriteConsoleMsg(UserIndex, MessageNewbie, FontTypeNames.FONTTYPE_CONSEJOVesA)

        End If
        
        If .GuildIndex > 0 Then
            Call Guilds_Connect(UserIndex)
                  
        End If

        If (.flags.Muerto = 0) Then
            .flags.SeguroResu = False
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff)
        Else
            .flags.SeguroResu = True
            Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn)

        End If
        
        Call WriteMultiMessage(UserIndex, eMessages.DragSafeOff)
            
        If Escriminal(UserIndex) Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)

        End If
        
        If .Stats.Gld < 0 Then .Stats.Gld = 0

        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXIDs.FXWARP, 0))
    
        Call WriteLoggedMessage(UserIndex)
    
        ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
        Call IntervaloPermiteSerAtacado(UserIndex, True)
    
        Call MostrarNumUsers
        '
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
        
        If MapInfo(.Pos.Map).LvlMin > .Stats.Elv Then
            Call WarpUserChar(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, False)

        End If
        
        If MapInfo(.Pos.Map).OnLoginGoTo.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡No puedes circular por este mapa en estos momentos. Te llevare a un sitio seguro!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).OnLoginGoTo.Map, MapInfo(.Pos.Map).OnLoginGoTo.X, MapInfo(.Pos.Map).OnLoginGoTo.Y, True, True)

        End If
        
        If Not CheckMap_Onlines(UserIndex, .Pos) Then
        
            If MapInfo(.Pos.Map).GoToOns.Map > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡¡¡No puedes circular por este mapa en estos momentos. Te llevare a la entrada del mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
                Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).GoToOns.Map, MapInfo(.Pos.Map).GoToOns.X, MapInfo(.Pos.Map).GoToOns.Y, True, True)
            Else
                Call EventWarpUser(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)

            End If
            
        End If
        
        ' Chequea si está en un mapa por horario y lo regresa a la ciudad principal
        If Not CheckMap_HourDay(UserIndex, .Pos) Then
            Call EventWarpUser(UserIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y)

        End If
        
        If .flags.Envenenado = 1 Then
            Call WriteUpdateEffect(UserIndex)

        End If
        
        Call WriteSendIntervals(UserIndex)
              
        Call UpdatePremium(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Tipea /SHOP para ingresar a la Tienda Oficial de la comunidad, donde podrás comprar por DSP", FontTypeNames.FONTTYPE_INFOBOLD)
              
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageUpdateEvento(EsModoEvento))
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageUpdateMeditation(.MeditationUser, Meditation(.MeditationSelected)))

        If EventLast > 0 Then
            Call WriteConsoleMsg(UserIndex, "Eventos> Nuevos eventos en curso. Tipea /TORNEOS para saber más.", FontTypeNames.FONTTYPE_CRITICO)

        End If
        
        If Not NewChar Then
            If .Stats.Elv <= LimiteNewbie Then
                Call WriteQuestInfo(UserIndex, True, 0)
                Call WriteConsoleMsg(UserIndex, "Misiones> Accede al panel de misiones desde la tecla 'ESC' o bien escribiendo /MISIONES", FontTypeNames.FONTTYPE_CRITICO)

            End If

        End If
        
        ' # Comprueba la permanencia de skins especiales (Clanes)
        Call Skins_CheckGuild(UserIndex, True)
        
        FlushBuffer UserIndex

    End With
    
    Exit Sub
ErrHandler:

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetFacciones_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    '*************************************************
    With UserList(UserIndex).Faction
    
        .FragsCiu = 0
        .FragsCri = 0
        .FragsOther = 0
        .ExFaction = 0
        .Range = 0
        .Status = 0
        .StartDate = vbNullString
        .StartElv = 0

    End With

    '<EhFooter>
    Exit Sub

ResetFacciones_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetFacciones " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetContadores_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 10/07/2010
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '05/20/2007 Integer - Agregue todas las variables que faltaban.
    '10/07/2010: ZaMa - Agrego los counters que faltaban.
    '*************************************************
    With UserList(UserIndex).Counters
        .SpeedHackCounter = 0
        .LastStep = 0
        .TimeGMBOT = 0
        .controlHechizos.HechizosCasteados = 0
        .controlHechizos.HechizosTotales = 0
        .Incinerado = 0
        .LastSave = 0
        .TimeLastReset = 0
        .PacketCount = 0
        .RuidoPocion = 0
        .RuidoDopa = 0
        .SpamMessage = 0
        .MessageSend = 0
        .FightInvitation = 0
        .FightSend = 0
        .Drawers = 0
        .DrawersCount = 0
        .TimeInfoMao = 0
        .TimeDrop = 0
        .TimeEquipped = 0
        .TimerPuedeCastear = 0
        .TimerPuedeRecibirAtaqueCriature = 0
        .TimeInfoChar = 0
        .TimeCommerce = 0
        .TimeMessage = 0
        .TimeInfoMao = 0
        .TimePublicationMao = 0
        
        .TimeInactive = 0
        .TimeBono = 0
        .TimeTelep = 0
        .TimeApparience = 0
        .TimeFight = 0
        .TimeCreateChar = 0
        
        .AGUACounter = 0
        .AsignedSkills = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .failedUsageAttempts = 0
        .failedUsageAttempts_Clic = 0
        .failedUsageCastSpell = 0
        .Frio = 0
        .goHome = 0
        .goHomeSec = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Lava = 0
        .Mimetismo = 0
        .Ocultando = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .Saliendo = False
        .Salir = 0
        .STACounter = 0
        .TiempoOculto = 0
        .TimerEstadoAtacable = 0
        .TimerGolpeMagia = 0
        .TimerGolpeUsar = 0
        .TimerUsarClick = 0
        .TimerLanzarSpell = 0
        .BuffoAceleration = 0
        .TimerShiftear = 0
        .CaspeoTime = 0
        .TimerMagiaGolpe = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeSerAtacado = 0
        .TimerPuedeTrabajar = 0
        .TimerPuedeUsarArco = 0
        .TimerUsar = 0
        .Trabajando = 0
        .Veneno = 0

    End With

    '<EhFooter>
    Exit Sub

ResetContadores_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetContadores " & "at line " & Erl

    '</EhFooter>
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetCharInfo_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex).Char
        .Body = 0
        .CascoAnim = 0
        .charindex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .speeding = 0

    End With

    '<EhFooter>
    Exit Sub

ResetCharInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetCharInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetBasicUserInfo_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .secName = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .Clase = 0
        .Genero = 0
        .Hogar = 0
        .Raza = 0
        
        .GroupIndex = 0
        .GroupRequest = vbNullString
        .GroupRequestTime = 0
        .GroupSlotUser = 0
        
        With .Stats
            .Elv = 0
            .Elu = 0
            .Exp = 0

            .Armour = 0
            .ArmourMag = 0
            .Damage = 0
            .DamageMag = 0
            .RegHP = 0
            .RegMANA = 0
            .Cooldown = 0
            .Attack = 0
            .Movement = 0
                  
            .NPCsMuertos = 0
            .SkillPts = 0
            .Gld = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0

        End With
        
    End With

    '<EhFooter>
    Exit Sub

ResetBasicUserInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetBasicUserInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ResetReputacion_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .promedio = 0

    End With

    '<EhFooter>
    Exit Sub

ResetReputacion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetReputacion " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserMeditation(ByRef IUser As User)

    '<EhHeader>
    On Error GoTo ResetUserMeditation_Err

    '</EhHeader>

    Dim A As Long
    
    With IUser

        For A = 1 To MAX_MEDITATION
            .MeditationUser(A) = 0
        Next A
        
        .MeditationSelected = 0

    End With

    '<EhFooter>
    Exit Sub

ResetUserMeditation_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserMeditation " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserOld(ByRef IUser As User)

    '<EhHeader>
    On Error GoTo ResetUserOld_Err

    '</EhHeader>
    
    With IUser.OldInfo
        .Clase = 0
        .Raza = 0
        .GldBlue = 0
        .GldRed = 0
        .MaxHp = 0
        .MaxMan = 0
        .MaxSta = 0
        .Elv = 0
        .Exp = 0
        .Head = 0
              
        Dim A As Long

        For A = 1 To MAXUSERHECHIZOS
            .UserSpell(A) = 0
        Next A
    
    End With
    
    '<EhFooter>
    Exit Sub

ResetUserOld_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserOld " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserObjectClaim(ByRef IUser As User)

    '<EhHeader>
    On Error GoTo ResetUserObjectClaim_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To MAX_INVENTORY_SLOTS
        
        With IUser.ObjectClaim(A)
            .Amount = 0
            .ObjIndex = 0
            .Equipped = 0

        End With
        
    Next A

    '<EhFooter>
    Exit Sub

ResetUserObjectClaim_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserObjectClaim " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetCharacterStats(ByRef IUser As User)

    With IUser.CharacterStats
        .PassiveAccumulated = 0
        
    End With
    
End Sub

Sub ResetUserFlags(ByRef IUser As User)

    '*************************************************
    'Author: Unknown
    'Last modified: 06/28/2008
    'Resetea todos los valores generales y las stats
    '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
    '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
    '*************************************************
    '<EhHeader>
    On Error GoTo ResetUserFlags_Err

    '</EhHeader>
            
    Dim A As Long
        
    IUser.QuestLast = 0
    ReDim IUser.QuestStats(1 To MAXUSERQUESTS) As tUserQuest
          
    Call ResetUserMeditation(IUser)
    Call ResetUserOld(IUser)
    Call Reto_ResetUserTemp(IUser)
    Call ResetUserObjectClaim(IUser)
    Call AntiFrags_ResetInfo(IUser)
        
    Call ResetCharacterStats(IUser)
          
    ' # Reset quests
    For A = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(IUser, A)
    Next A
          
    Dim NullBot As tBotIntelligence
            
    For A = 1 To BOT_MAX_USER
            
        IUser.BotIntelligence(A) = NullBot
    Next A

    Dim i As Long
            
    With IUser
            
        .PosOculto.Map = 0
        .PosOculto.X = 0
        .PosOculto.Y = 0
            
        ReDim .Skins.ObjIndex(1 To MAX_INVENTORY_SKINS) As Integer
            
        .Skins.Last = 0
        .Skins.ArmourIndex = 0
        .Skins.HelmIndex = 0
        .Skins.ShieldIndex = 0
        .Skins.WeaponIndex = 0
        .Skins.WeaponArcoIndex = 0
        .Skins.WeaponDagaIndex = 0
            
        For i = 1 To MAX_INVENTORY_SKINS
            .Skins.ObjIndex(i) = 0
        Next i
            
        .GuildIndex = 0
        .GuildRange = 0
        .GuildSlot = 0
        .UseObj_Clic = 0
        .UseObj_Init_Clic = 0
        .UseObj_U = 0
        .UseObj_Init_U = 0
        .Next_UseItem = False
        
        .LastPotion = eLastPotion.eNullPotion
        .PotionBlue_Clic = 0
        .PotionBlue_Clic_Interval = 0
        .PotionBlue_U = 0
        .PotionBlue_U_Interval = 0
        .PotionRed_Clic = 0
        .PotionRed_Clic_Interval = 0
        .PotionRed_U = 0
        .PotionRed_U_Interval = 0
        
        For A = 0 To 1
            .interval(A).IAttack = 0
            .interval(A).IDrop = 0
            .interval(A).ISpell = 0
            .interval(A).IUse = 0
            .interval(A).ILeftClick = 0
        Next A
        
        .MascotaIndex = 0
        .DañoApu = 0
        .UserKey = 0
        .Power = False
        .UserLastClick = 0
        .UserLastClick_Tolerance = 0
        
        With .Stats
            .Points = 0

        End With

    End With
    
    With IUser.flags

        .RachasTemp = 0
        .Rachas = 0
        .RedLimit = 0
        .RedUsage = 0
        .RedValid = False
        .BotList = 0
        .TeleportInvoker = 0
        .LastInvoker = 0
        .TempAccount = vbNullString
        .TempPasswd = vbNullString
        .DeslogeandoCuenta = False
        .StreamUrl = vbNullString
        .ModoStream = False
        .ToleranceCheat = 0
        .DragBlocked = False
        .GmSeguidor = 0
        .LastSlotClient = 0
        .MenuCliente = 0
        .Montando = 0
        .DesafiosGanados = 0
        .Desafiando = 0
        .SelectedBono = 0
        .Premium = 0
        .Streamer = 0
        .Bronce = 0
        .Transform = 0
        .TransformVIP = 0
        .Plata = 0
        .Oro = 0
        .SlotReto = 0
        .SlotEvent = 0
        .SlotUserEvent = 0
        .SlotRetoUser = 255
        .SlotFast = 0
        .SlotFastUser = 0
        .SelectedEvent = 0
        .FightTeam = 0
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PrivEspecial = False
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .Silenciado = 0
        .AdminPerseguible = False
        .LastMap = 0
        .Traveling = 0
        .AtacablePor = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .ShareNpcWith = 0
        .EnConsulta = False
        .Ignorado = False
        .SendDenounces = False
        .ParalizedBy = vbNullString
        .ParalizedByIndex = 0
        .ParalizedByNpcIndex = 0
        
    End With

    '<EhFooter>
    Exit Sub

ResetUserFlags_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserFlags " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetUserSpells_Err

    '</EhHeader>

    Dim LoopC As Long

    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC

    '<EhFooter>
    Exit Sub

ResetUserSpells_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserSpells " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetUserBanco_Err

    '</EhHeader>

    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
        UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
    '<EhFooter>
    Exit Sub

ResetUserBanco_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserBanco " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo LimpiarComercioSeguro_Err

    '</EhHeader>

    With UserList(UserIndex).ComUsu

        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

LimpiarComercioSeguro_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.LimpiarComercioSeguro " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetUserSlot_Err

    '</EhHeader>

    Dim i As Long

    Call LimpiarComercioSeguro(UserIndex)
    Call ResetFacciones(UserIndex)
    Call ResetContadores(UserIndex)
    Call ResetPacketRateData(UserIndex)
    Call ResetCharInfo(UserIndex)
    Call ResetBasicUserInfo(UserIndex)
    Call ResetReputacion(UserIndex)
    Call ResetUserFlags(UserList(UserIndex))
    Call ResetKeyPackets(UserIndex)
    Call ResetPointer(UserIndex, Point_Inv)
    Call ResetPointer(UserIndex, Point_Spell)
    Call LimpiarInventario(UserIndex)
    Call ResetUserSpells(UserIndex)
    Call ResetUserBanco(UserIndex)
    Call LimpiarMensajes(UserIndex)

    With UserList(UserIndex).ComUsu
        .Acepto = False
    
        For i = 1 To MAX_OFFER_SLOTS
            .cant(i) = 0
            .Objeto(i) = 0
        Next i
        
        .EldhirAmount = 0
        .GoldAmount = 0
        .DestNick = vbNullString
        .DestUsu = 0

    End With
        
    If UserList(UserIndex).flags.OwnedNpc <> 0 Then
        Call PerdioNpc(UserIndex)

    End If
 
    '<EhFooter>
    Exit Sub

ResetUserSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.ResetUserSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub CloseUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CloseUser_Err

    '</EhHeader>

    Dim N    As Integer

    Dim Map  As Integer

    Dim Name As String

    Dim i    As Integer

    Dim aN   As Integer

    With UserList(UserIndex)
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
    
        Map = .Pos.Map
        Name = UCase$(.Name)
    
        .Char.FX = 0
        .Char.loops = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 0, 0))
    
        .flags.UserLogged = False
        .Counters.Saliendo = False
    
        'Le devolvemos el body y head originales
        If .flags.AdminInvisible = 1 Then
            .Char.Body = .flags.OldBody
            .Char.Head = .flags.OldHead
            .flags.AdminInvisible = 0

        End If
    
        'si esta en party le devolvemos la experiencia
        If .GroupIndex > 0 Then Call mGroup.AbandonateGroup(UserIndex)
    
        'Save statistics
        'Call Statistics.UserDisconnected(UserIndex)
    
        ' Grabamos el personaje del usuario
        Call SaveUser(UserList(UserIndex), CharPath & Name & ".chr")

        'Quitar el dialogo
        'If MapInfo(Map).NumUsers > 0 Then
        '    Call SendToUserArea(UserIndex, "QDL" & .Char.charindex)
        'End If
    
        If MapInfo(Map).NumUsers > 0 Then
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))

        End If
        
        'Update Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
        MapInfo(Map).Players.Remove UserIndex
    
        If MapInfo(Map).NumUsers < 0 Then
            MapInfo(Map).NumUsers = 0

        End If

        'End If
    
        'Borrar el personaje
        If .Char.charindex > 0 Then
            Call EraseUserChar(UserIndex, .flags.AdminInvisible = 1)

        End If
    
        'Borrar mascota
        If .MascotaIndex Then
            If Npclist(.MascotaIndex).flags.NPCActive Then Call QuitarNPC(.MascotaIndex)

        End If
            
        ' Remove Position
        Call Guilds_UpdatePosition(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

CloseUser_Err:
    LogError Err.description & vbCrLf & "in CloseUser " & "at line " & Erl

    '</EhFooter>
End Sub

Sub ReloadSokcet()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    'Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    'If NumUsers <= 0 Then
    'Call WSApiReiniciarSockets
    'Else
    '       Call apiclosesocket(SockListen)
    '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    'End If

    Exit Sub

ErrHandler:
    Call LogError("Error en CheckSocketState " & Err.number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EcharPjsNoPrivilegiados_Err

    '</EhHeader>

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser

        If UserList(LoopC).flags.UserLogged Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call Protocol.Kick(LoopC)
            
            End If

        End If

    Next LoopC

    '<EhFooter>
    Exit Sub

EcharPjsNoPrivilegiados_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.EcharPjsNoPrivilegiados " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub User_GenerateNewHead(ByVal UserIndex As Integer, ByVal Tipe As Byte)

    '<EhHeader>
    On Error GoTo User_GenerateNewHead_Err

    '</EhHeader>
    
    Dim NewHead    As Integer

    Dim UserRaza   As Byte

    Dim UserGenero As Byte
    
    UserGenero = UserList(UserIndex).Genero
    UserRaza = UserList(UserIndex).Raza

    Select Case UserGenero

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                        
                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(502, 546)
                    Else
                        NewHead = RandomNumber(1, 25)

                    End If
                    
                Case eRaza.Elfo

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(577, 608)
                    Else
                        NewHead = RandomNumber(102, 111)

                    End If
                        
                Case eRaza.Drow
                    '

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(639, 669)
                    Else
                        NewHead = RandomNumber(201, 205)

                    End If
                        
                Case eRaza.Enano

                    '
                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(700, 729)
                    Else
                        NewHead = RandomNumber(301, 305)

                    End If

                Case eRaza.Gnomo

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(760, 789)
                    Else
                        NewHead = RandomNumber(401, 405)

                    End If

            End Select

        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(547, 576)
                    Else
                        NewHead = RandomNumber(71, 75)

                    End If
                    
                Case eRaza.Elfo

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(609, 638)
                    Else
                        NewHead = RandomNumber(170, 176)

                    End If

                Case eRaza.Drow

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(670, 699)
                    Else
                        NewHead = RandomNumber(270, 276)

                    End If

                Case eRaza.Gnomo

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(790, 819)
                    Else
                        NewHead = RandomNumber(471, 475)

                    End If

                Case eRaza.Enano
                    '

                    If Tipe = eEffectObj.e_NewHead Then
                        NewHead = RandomNumber(730, 759)
                    Else
                        NewHead = RandomNumber(370, 371)

                    End If

            End Select

    End Select
    
    UserList(UserIndex).Char.Head = NewHead
    UserList(UserIndex).OrigChar.Head = NewHead
    
    '<EhFooter>
    Exit Sub

User_GenerateNewHead_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.TCP.User_GenerateNewHead " & "at line " & Erl

    '</EhFooter>
End Sub

Sub ResetPacketRateData(ByVal UserIndex As Integer)

    On Error GoTo ResetPacketRateData_Err

    Dim i As Long
        
    With UserList(UserIndex)
        
        For i = 1 To MAX_PACKET_COUNTERS
            .MacroIterations(i) = 0
            .PacketTimers(i) = 0
            .PacketCounters(i) = 0
        Next i
            
    End With
        
    Exit Sub
        
ResetPacketRateData_Err:

End Sub
