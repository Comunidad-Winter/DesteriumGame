Attribute VB_Name = "mCVC"
Option Explicit

Private Type tCVC
    Run As Boolean
    GuildOne() As Integer
    GuildTwo() As Integer
    
End Type

Public CVC As tCVC
Private Function CanPlayCVC(ByRef Users() As String) As Boolean
          Dim tUser As Integer
          Dim LoopC As Integer
          
10        For LoopC = LBound(Users()) To UBound(Users())
20            If Users(LoopC) = vbNullString Then
30                CanPlayCVC = False
40                Exit Function
50            End If
              
60            tUser = NameIndex(Users(LoopC))
              
70            If tUser <= 0 Then
                  ' Personaje offline
80                CanPlayCVC = False
90                Exit Function
100           End If
              
110           With UserList(tUser)
120               If .flags.Muerto Then Exit Function
130               If .flags.SlotReto > 0 Then Exit Function
140               If .flags.SlotEvent > 0 Then Exit Function
                  
150               If .Counters.Pena > 0 Then Exit Function
                  
160           End With
170       Next LoopC
          
          
180       CanPlayCVC = True
End Function
Private Function UserIsGuildOne(ByVal UserIndex As Integer) As Boolean
          Dim LoopC As Integer
          
10        With CVC
20            For LoopC = LBound(.GuildOne()) To UBound(.GuildTwo())
30                If .GuildOne(LoopC) = UserIndex Then
40                    UserIsGuildOne = True
50                    Exit For
60                End If
70            Next LoopC
80        End With
End Function
Private Function UserIsGuildTwo(ByVal UserIndex As Integer) As Boolean
          Dim LoopC As Integer
          
10        With CVC
20            For LoopC = LBound(.GuildTwo()) To UBound(.GuildTwo())
30                If .GuildTwo(LoopC) = UserIndex Then
40                    UserIsGuildTwo = True
50                    Exit For
60                End If
70            Next LoopC
80        End With
End Function
Private Function ContinueGuildOne() As Boolean
          Dim LoopC As Integer
          
10        For LoopC = LBound(CVC.GuildOne()) To UBound(CVC.GuildOne())
20            If CVC.GuildOne(LoopC) > 0 Then
30                If UserList(CVC.GuildOne(LoopC)).flags.Muerto = 0 Then
40                    ContinueGuildOne = True
50                    Exit For
60                End If
70            End If
80        Next LoopC
          
End Function
Private Function ContinueGuildTwo() As Boolean
          Dim LoopC As Integer
          
10        For LoopC = LBound(CVC.GuildTwo()) To UBound(CVC.GuildTwo())
20            If CVC.GuildOne(LoopC) > 0 Then
30                If UserList(CVC.GuildTwo(LoopC)).flags.Muerto = 0 Then
40                    ContinueGuildTwo = True
50                    Exit For
60                End If
70            End If
80        Next LoopC
          
End Function
Public Sub CloseUserCvc(ByVal UserIndex As Integer)

          Dim LoopC As Integer
          
10        With UserList(UserIndex)
20            .flags.InCVC = False
              
30            For LoopC = LBound(CVC.GuildOne()) To UBound(CVC.GuildOne())
40                If CVC.GuildOne(LoopC) = UserIndex Then
50                    CVC.GuildOne(LoopC) = 0
60                    Exit Sub
70                End If
80            Next LoopC
              
90            For LoopC = LBound(CVC.GuildTwo()) To UBound(CVC.GuildTwo())
100               If CVC.GuildTwo(LoopC) = UserIndex Then
110                   CVC.GuildTwo(LoopC) = 0
120                   Exit Sub
130               End If
140           Next LoopC
              
150       End With
End Sub
Private Function ValidNumberGuild(ByVal UserIndex As Integer, _
                                    ByVal tUser As Integer) As Byte
                                      
          Dim GuildOne As Integer
          Dim GuildTwo As Integer
          Dim CountOne As Byte
          Dim CountTwo As Byte
          Dim onlinelist As String
          Dim Users() As String
          
   On Error GoTo ValidNumberGuild_Error

10        ValidNumberGuild = 0
          
20        GuildOne = UserList(UserIndex).GuildIndex
30        GuildTwo = UserList(tUser).GuildIndex
          
40        onlinelist = modGuilds.m_ListaDeMiembrosOnline(UserIndex, GuildOne, True)
50        Users = Split(onlinelist, ",")
60        CountOne = UBound(Users)
          
70        onlinelist = modGuilds.m_ListaDeMiembrosOnline(tUser, GuildTwo, True)
80        Users = Split(onlinelist, ",")
90        CountTwo = UBound(Users)
          
100       If CountOne < 2 Then Exit Function
110       If CountTwo < 2 Then Exit Function
          
          
120       If CountOne < CountTwo Then
130           ValidNumberGuild = CountOne + 1
140       Else
150           ValidNumberGuild = CountTwo + 1
160       End If

   On Error GoTo 0
   Exit Function

ValidNumberGuild_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidNumberGuild of Módulo mCVC in line " & Erl
End Function
Public Sub SendFightGuild(ByVal UserIndex As Integer, _
                            ByVal tUser As Integer)
   On Error GoTo SendFightGuild_Error

10        With UserList(UserIndex)
20            If tUser <= 0 Then Exit Sub
30            If .GuildIndex = 0 Then Exit Sub
40            If UserList(tUser).GuildIndex = 0 Then Exit Sub
50            If UserIndex = tUser Then Exit Sub
              
60            If Not UCase$(modGuilds.GuildLeader(.GuildIndex)) <> UCase$(UserList(tUser).Name) Then
70                WriteConsoleMsg UserIndex, "El personaje al que acabas de enviar solicitud no es el dueño del clan.", FontTypeNames.FONTTYPE_INFO
80                Exit Sub
90            End If
              
100           ReDim .RetoTemp.Users(0 To 1) As String
              
110           .RetoTemp.Users(0) = UCase$(.Name)
120           .RetoTemp.Users(1) = UCase$(UserList(tUser).Name)
130           .RetoTemp.Tipo = FightClan
              
140           WriteConsoleMsg UserIndex, "Has enviado una solicitud de CVC al lider del clan. Espera pronta noticias de él.", FontTypeNames.FONTTYPE_INFO
150           WriteConsoleMsg tUser, "Has recibido una solicitud de CVC del lider " & UserList(UserIndex).Name & " <" & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & "> Tipea /CVC " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_VENENO
              
160       End With

   On Error GoTo 0
   Exit Sub

SendFightGuild_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SendFightGuild of Módulo mCVC in line " & Erl
End Sub

Private Sub SetUsersGuild(ByVal UserIndex As Integer, _
                            ByVal tUser As Integer, _
                            ByVal ValidNumber As Byte, _
                            ByRef UsersOne() As String, _
                            ByRef UsersTwo() As String)
          
   On Error GoTo SetUsersGuild_Error

10        Dim GuildOne As Integer: GuildOne = UserList(UserIndex).GuildIndex
20        Dim GuildTwo As Integer: GuildTwo = UserList(tUser).GuildIndex
          Dim onlinelist As String
          Dim One As String
          Dim Two As String
          Dim LoopC As Integer
          
          ' Dejamos los usuarios del primer clan
30        onlinelist = modGuilds.m_ListaDeMiembrosOnline(UserIndex, GuildOne, True)
40        UsersOne = Split(onlinelist, ",")

50        ReDim Preserve UsersOne(LBound(UsersOne) To ValidNumber - 1) As String
          
          ' Dejamos los usuarios del segundo clan
60        onlinelist = modGuilds.m_ListaDeMiembrosOnline(tUser, GuildTwo, True)
70        UsersTwo = Split(onlinelist, ",")

80        ReDim Preserve UsersTwo(LBound(UsersTwo) To ValidNumber - 1) As String
          
          

   On Error GoTo 0
   Exit Sub

SetUsersGuild_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SetUsersGuild of Módulo mCVC in line " & Erl
          
End Sub
Public Sub AcceptFightGuild(ByVal UserIndex As Integer, _
                            ByVal tUser As Integer)
                                  
   On Error GoTo AcceptFightGuild_Error

10        With UserList(UserIndex)
              Dim Users() As String
              Dim ValidNumber As Byte

              Dim ArrayNulo As Long
              Dim UsersOne() As String
              Dim UsersTwo() As String
              Dim LoopC As Integer
              
20            If .GuildIndex = 0 Then Exit Sub
30            If UserList(tUser).GuildIndex = 0 Then Exit Sub
              
40            GetSafeArrayPointer UserList(tUser).RetoTemp.Users, ArrayNulo
50            If ArrayNulo <= 0 Then Exit Sub
              
60            If UCase$(UserList(tUser).RetoTemp.Users(1)) <> UCase$(.Name) Then
70                WriteConsoleMsg UserIndex, "El personaje no te envió ninguna solicitud.", FontTypeNames.FONTTYPE_INFO
80                Exit Sub
90            End If
              
100           ValidNumber = ValidNumberGuild(UserIndex, tUser)

110           If ValidNumber > 0 Then
120               SetUsersGuild UserIndex, tUser, ValidNumber, UsersOne, UsersTwo
                  
130               If Not CanPlayCVC(UsersOne) Then Exit Sub
140               If Not CanPlayCVC(UsersTwo) Then Exit Sub
                  
150               With CVC
160                   ReDim .GuildOne(LBound(UsersOne) To UBound(UsersOne)) As Integer
170                   ReDim .GuildTwo(LBound(UsersTwo) To UBound(UsersTwo)) As Integer
                      
180                   For LoopC = LBound(UsersOne) To UBound(UsersTwo)
190                       .GuildOne(LoopC) = NameIndex(UsersOne(LoopC))
200                       .GuildTwo(LoopC) = NameIndex(UsersTwo(LoopC))
                          
210                       UserList(.GuildOne(LoopC)).flags.InCVC = True
220                       UserList(.GuildTwo(LoopC)).flags.InCVC = True
                          
230                       UserList(.GuildTwo(LoopC)).PosAnt.map = UserList(.GuildTwo(LoopC)).Pos.map
240                       UserList(.GuildTwo(LoopC)).PosAnt.X = UserList(.GuildTwo(LoopC)).Pos.X
250                       UserList(.GuildTwo(LoopC)).PosAnt.Y = UserList(.GuildTwo(LoopC)).Pos.Y
                          
260                       UserList(.GuildOne(LoopC)).PosAnt.map = UserList(.GuildOne(LoopC)).Pos.map
270                       UserList(.GuildOne(LoopC)).PosAnt.X = UserList(.GuildOne(LoopC)).Pos.X
280                       UserList(.GuildOne(LoopC)).PosAnt.Y = UserList(.GuildOne(LoopC)).Pos.Y
                          
290                       WarpUsersCVC True, .GuildOne
300                       WarpUsersCVC False, .GuildTwo
310                   Next LoopC
320               End With
330           End If
              
340       End With

   On Error GoTo 0
   Exit Sub

AcceptFightGuild_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AcceptFightGuild of Módulo mCVC in line " & Erl
End Sub
Private Sub WarpUsersCVC(ByVal Enemy As Boolean, ByRef GuildUsers() As Integer)
          Dim LoopC As Integer
          Dim tUser As Integer
          Dim Pos As WorldPos
          
   On Error GoTo WarpUsersCVC_Error

10        For LoopC = LBound(GuildUsers()) To UBound(GuildUsers())
20            tUser = GuildUsers(LoopC)
              
30            If tUser > 0 Then
40                Pos.map = 220
                  
50                If Enemy Then
60                    Pos.X = 41
70                    Pos.Y = 17
80                    UserList(tUser).flags.FightTeam = 1
90                Else
100                   Pos.X = 41
110                   Pos.Y = 56
120                   UserList(tUser).flags.FightTeam = 2
130               End If
                  
140               With UserList(tUser)
                      
170                   .Counters.TimeFight = 10
180                   WriteUserInEvent tUser
190                   RefreshCharStatus tUser

150                   ClosestStablePos Pos, Pos
160                   WarpUserChar tUser, Pos.map, Pos.X, Pos.Y, False
                      

200               End With
210           End If
220       Next LoopC

   On Error GoTo 0
   Exit Sub

WarpUsersCVC_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure WarpUsersCVC of Módulo mCVC in line " & Erl
End Sub

Private Sub FinishCVC(ByRef Users() As Integer)
          ' • Finalizamos el CVC
          Dim LoopC As Integer
          Dim GuildOne As String
          Dim GuildTwo As String
          
   On Error GoTo FinishCVC_Error

10        With CVC
20            Debug.Print
30            GuildOne = guilds(UserList(.GuildOne(0)).GuildIndex).GuildName
40            GuildTwo = guilds(UserList(.GuildTwo(0)).GuildIndex).GuildName
              
50            SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("RetosClan» " & GuildOne & " vs " & GuildTwo & ". El ganador fue el clan " & guilds(UserList(Users(0)).GuildIndex).GuildName, FontTypeNames.FONTTYPE_INFO)
              
60            For LoopC = LBound(.GuildOne()) To UBound(.GuildOne())
70                If .GuildOne(LoopC) > 0 Then
80                    UserList(.GuildOne(LoopC)).flags.InCVC = False
90                    WarpPosAnt .GuildOne(LoopC)
100               End If
                  
110               If .GuildTwo(LoopC) > 0 Then
120                   WarpPosAnt .GuildTwo(LoopC)
130               End If
                  
140           Next LoopC
              
150           .Run = False
         
160       End With

   On Error GoTo 0
   Exit Sub

FinishCVC_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure FinishCVC of Módulo mCVC in line " & Erl
End Sub


Public Sub UserdieCVC(ByVal UserIndex As Integer)
           ' • Un personaje en CVC es matado por otro.
          Dim LoopC As Integer
          Dim strTemp As String
          Dim SlotReto As Byte
          Dim TeamUser As Byte
          Dim Rounds As Byte
          
          ' Personaje perteneciente al GuildOne.
          ' Si no puede continuar su equipo terminamos el CVC.
   On Error GoTo UserdieCVC_Error

10        If UserIsGuildOne(UserIndex) Then
20            If Not ContinueGuildOne Then
30                FinishCVC CVC.GuildTwo
40            End If
              
50            Exit Sub
60        End If
          
          ' Personaje perteneciente al GuildTwo.
          ' Si no puede continuar su equipo terminamos el CVC.
70        If UserIsGuildTwo(UserIndex) Then
80            If Not ContinueGuildTwo Then
90                FinishCVC CVC.GuildOne
100           End If
              
110           Exit Sub
120       End If

   On Error GoTo 0
   Exit Sub

UserdieCVC_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserdieCVC of Módulo mCVC in line " & Erl
End Sub

