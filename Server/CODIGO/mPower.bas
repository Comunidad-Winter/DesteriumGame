Attribute VB_Name = "mPower"
' Programado para Desterium AO EXODO III
Option Explicit

Private Type tPower

    Desc As String
    UserIndex As Integer
    FindMap As Boolean
    PreviousUser As String
    Active As Boolean
    Time As Integer

End Type

Private NumPower As Byte

Public Power     As tPower

Public Sub Power_Search(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    Dim A As Long, B As Long, C As Long
    
    If (NumUsers + UsersBot) < 30 Then Exit Sub
    If Power.UserIndex > 0 Then Exit Sub
    
    With Power
        
        For A = 2 To NumMaps

            With MapInfo(A)

                If .Poder = 1 Then
                    If Power.UserIndex = 0 Then
                        If Not StrComp(UCase$(Power.PreviousUser), UCase$(UserList(UserIndex).Name)) = 0 Then
                            If UserList(UserIndex).flags.UserLogged And UserList(UserIndex).flags.Muerto = 0 And Not EsGm(UserIndex) And UserList(UserIndex).Clase <> eClass.Thief Then
                                    
                                Call Power_Set(UserIndex, Power.UserIndex)
                                Call Power_Message
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_WARP, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 52, 5))
                                'Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayEffect(128, NO_3D_SOUND, NO_3D_SOUND))
                                
                                Exit Sub
                                    
                            End If

                        End If

                    End If

                End If

            End With

        Next A
    
    End With

    Exit Sub
ErrHandler:

End Sub

Public Sub Power_Search_All()

    On Error GoTo ErrHandler
    
    'If Not Power.Active Then Exit Sub
    If (NumUsers + UsersBot) < 30 Then Exit Sub
    
    Dim A As Long
    
    For A = 1 To LastUser

        If UserList(A).flags.UserLogged Then
            Call Power_Search(A)

        End If

    Next A

    Exit Sub
ErrHandler:

End Sub

Public Sub Power_Set(ByVal UserIndex As Integer, ByVal PreviousUser As Integer)

    '<EhHeader>
    On Error GoTo Power_Set_Err

    '</EhHeader>
    
    With Power

        If Not PreviousUser = 0 Then
            .PreviousUser = UCase$(UserList(PreviousUser).Name)

        End If
            
        .UserIndex = UserIndex
        .Time = 900 '1800
        
        ' Si el UserIndex se resetea, buscamos al nuevo poder
        If .UserIndex = 0 Then
            
            Call Power_Search_All
        Else
            Call RefreshCharStatus(.UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

Power_Set_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mPower.Power_Set " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Power_Message()
    
    On Error GoTo ErrHandler
    
    With Power
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Desc & UserList(.UserIndex).Name, FontTypeNames.FONTTYPE_GUILD))

    End With
    
    Exit Sub
ErrHandler:
    
End Sub

Public Function Power_CheckTime() As Boolean

    '<EhHeader>
    On Error GoTo Power_CheckTime_Err

    '</EhHeader>
    Dim Time As String
    
    'Time = Format(Now, "hh:mm")
    
    If Power.Time > 0 Then
        Power.Time = Power.Time - 1
        
        If Power.Time = 0 Then
            Call WriteConsoleMsg(Power.UserIndex, "Has perdido el poder", FontTypeNames.FONTTYPE_INFORED)
            Call RefreshCharStatus(Power.UserIndex)
            Call Power_Set(0, Power.UserIndex)

        End If

    End If
 
    '<EhFooter>
    Exit Function

Power_CheckTime_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mPower.Power_CheckTime " & "at line " & Erl
        
    '</EhFooter>
End Function
