Attribute VB_Name = "mDias"
Option Explicit

Private Const MAX_BAN_DIAS As Byte = 200
Private Const MAX_DIAS_DIOS As Byte = 70

Private Type tBanDias
    UserName As String
    FechaUnBan As String ' FORMATO DE FECHA : "31-05-2018"
End Type

Private BanDias(1 To MAX_BAN_DIAS) As tBanDias

Private Type tDiosDias
    UserName As String
    FechaFinDios As String
End Type

Private DiosDias(1 To MAX_DIAS_DIOS) As tDiosDias

Public Sub CreateDataDias()
          Dim intFile As Integer
          Dim i As Integer
          
10        intFile = FreeFile

20        Open App.Path & "\DAT\DATADIAS.DAT" For Output As #intFile
30        Print #intFile, "[BAN]"

40        For i = 1 To MAX_BAN_DIAS
50            Print #intFile, i & "=|"
60        Next i
          
70        Print #intFile, vbNullString
80        Print #intFile, vbNullString
          
90        Print #intFile, "[DIOS]"
          
100       For i = 1 To MAX_DIAS_DIOS
110           Print #intFile, i & "=|"
120       Next i
          
130       Close #intFile
End Sub

Public Sub LoadDataDias()

          Dim LoopC As Integer
          Dim strTemp As String
          
10        If Not FileExist(App.Path & "\DAT\DATADIAS.DAT", vbNormal) Then
20            CreateDataDias
30        End If
          
40        For LoopC = 1 To MAX_BAN_DIAS
50            With BanDias(LoopC)
60                strTemp = GetVar(App.Path & "\DAT\DATADIAS.DAT", "BAN", LoopC)
                  
70                .UserName = ReadField(1, strTemp, Asc("|"))
80                .FechaUnBan = ReadField(2, strTemp, Asc("|"))
90            End With
100       Next LoopC
          
110       For LoopC = 1 To MAX_DIAS_DIOS
120           With DiosDias(LoopC)
130               strTemp = GetVar(App.Path & "\DAT\DATADIAS.DAT", "DIOS", LoopC)
                  
140               .UserName = ReadField(1, strTemp, Asc("|"))
150               .FechaFinDios = ReadField(2, strTemp, Asc("|"))
160           End With
170       Next LoopC
          
          
End Sub
Private Function FreeSlotDios() As Byte
          Dim LoopC As Integer
          
10        For LoopC = 1 To MAX_DIAS_DIOS
20            With DiosDias(LoopC)
30                If .UserName = vbNullString Then
40                    FreeSlotDios = LoopC
50                    Exit For
60                End If
70            End With
80        Next LoopC
End Function
Private Function FreeSlotBan() As Byte
          Dim LoopC As Integer
          
10        For LoopC = 1 To MAX_BAN_DIAS
20            With BanDias(LoopC)
30                If .UserName = vbNullString Then
40                    FreeSlotBan = LoopC
50                    Exit For
60                End If
70            End With
80        Next LoopC
End Function
Public Function IsEffectDios(ByVal UserName As String) As Boolean
          Dim LoopC As Integer
          
10        IsEffectDios = False
          
20        For LoopC = 1 To MAX_DIAS_DIOS
30            With DiosDias(LoopC)
40                If .UserName = UserName Then
50                    IsEffectDios = True
60                    Exit For
70                End If
80            End With
90        Next LoopC
          
End Function

Public Sub BanUserDias(ByVal UserIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal strDate As String)
                                       
          Dim SlotFree As Byte
          Dim tIndex As Integer
          Dim cantPenas As Byte
          
   On Error GoTo BanUserDias_Error

10        SlotFree = FreeSlotBan
                        
20        If SlotFree <> 0 Then
30            BanDias(SlotFree).UserName = UCase$(UserName)
40            BanDias(SlotFree).FechaUnBan = strDate
50            WriteVar App.Path & "\DAT\DATADIAS.DAT", "BAN", SlotFree, UCase$(UserName & "|" & strDate)
                  
60            WriteConsoleMsg UserIndex, "Has baneado al personaje " & UserName & " hasta la fecha " & strDate, FontTypeNames.FONTTYPE_INFO
                  
70            tIndex = NameIndex(UserName)
80            If tIndex > 0 Then
90                UserList(tIndex).flags.Ban = 1
100               Call FlushBuffer(tIndex)
110               Call CloseSocket(tIndex)
120           End If
              
130           Call Ban(UserName, "Ban por DIAS", "Baneado hasta la fecha " & strDate)
              
140           Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
150           cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
160           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
170           Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, "Ban por DIAS" & ": " & Date$ & " hasta " & strDate)
180       End If

   On Error GoTo 0
   Exit Sub

BanUserDias_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure BanUserDias of Módulo mDias in line " & Erl
End Sub
Public Sub TransformarUserDios(ByVal UserIndex As Integer, _
                                ByVal tIndex As Integer, _
                                ByVal strDate As String)

          Dim SlotNew As Byte
          
   On Error GoTo TransformarUserDios_Error

10        With UserList(tIndex)
20            SlotNew = FreeSlotDios
              
30            If SlotNew = 0 Then Exit Sub
              
40            If IsEffectDios(.Name) Then
50                WriteConsoleMsg UserIndex, "El personaje ya es DIOS.", FontTypeNames.FONTTYPE_INFO
60                Exit Sub
70            End If
              
80            DiosDias(SlotNew).UserName = UCase$(.Name)
90            DiosDias(SlotNew).FechaFinDios = strDate
100           WriteVar App.Path & "\DAT\DATADIAS.DAT", "DIOS", SlotNew, UCase$(.Name & "|" & strDate)
110           .flags.IsDios = True
              
120           .Stats.MaxHp = .Stats.MaxHp + 30
130           WriteUpdateUserStats tIndex
140           Call SaveUser(tIndex, CharPath & .Name & ".chr")
              
150           WriteConsoleMsg UserIndex, "Has hecho al personaje " & .Name & " el poder de DIOS hasta la fecha " & strDate, FontTypeNames.FONTTYPE_INFO
160           WriteConsoleMsg tIndex, "Te has convertido en DIOS. A partir de este momento hasta la fecha " & strDate & " serás inmune a la parálisis", FontTypeNames.FONTTYPE_GMMSG
170       End With

   On Error GoTo 0
   Exit Sub

TransformarUserDios_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure TransformarUserDios of Módulo mDias in line " & Erl
End Sub

Public Sub LoopDias()
          
          ' Chequeo cada 30 minutos diferencia de baneos y de usuarios con poder DIOS
          Dim LoopC As Long
          Dim tIndex As Integer
          Dim MaxHp As Integer
          
   On Error GoTo LoopDias_Error

10        For LoopC = 1 To MAX_BAN_DIAS
20            With BanDias(LoopC)
30                    If .UserName <> vbNullString Then
40                        If DateDiff("d", Date, .FechaUnBan) = 0 Then
50                            UnBan .UserName
60                            .UserName = vbNullString
70                            .FechaUnBan = vbNullString
80                            WriteVar App.Path & "\DAT\DATADIAS.DAT", "BAN", LoopC, "|"
90                        End If
100                   End If
110           End With
120       Next LoopC
          
130       For LoopC = 1 To MAX_DIAS_DIOS
140           With DiosDias(LoopC)
150               If .UserName <> vbNullString Then
160                   If DateDiff("d", Date, .FechaFinDios) = 0 Then

                          
170                       tIndex = NameIndex(UCase$(.UserName))
                          
180                       If tIndex > 0 Then
                              ' Personaje online
190                           WriteConsoleMsg tIndex, "El poder de Dios ha terminado. A partir de este momento volveras a poder ser paralizado con la misma frecuencia y tu vida volvió a ser la de antes.", FontTypeNames.FONTTYPE_CITIZEN
200                           UserList(tIndex).flags.IsDios = False
                              
210                           UserList(tIndex).Stats.MaxHp = UserList(tIndex).Stats.MaxHp - 30
220                           UserList(tIndex).Stats.MinHp = UserList(tIndex).Stats.MaxHp
230                           WriteUpdateUserStats tIndex
                              
240                           SaveUser tIndex, CharPath & UCase$(UserList(tIndex).Name) & ".chr"
                              
250                       Else
260                           MaxHp = val(GetVar(CharPath & .UserName & ".chr", "STATS", "MAXHP"))
270                           WriteVar CharPath & .UserName & ".chr", "STATS", "MAXHP", MaxHp - 30
280                       End If
                          
                          
290                       .FechaFinDios = vbNullString
300                       .UserName = vbNullString
                          
310                       WriteVar App.Path & "\DAT\DATADIAS.DAT", "DIOS", LoopC, "|"
320                   End If
330               End If
340           End With
350       Next LoopC

   On Error GoTo 0
   Exit Sub

LoopDias_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure LoopDias of Módulo mDias in line " & Erl
          
End Sub
