Attribute VB_Name = "mApuestas"
Option Explicit

Private Const PORCENTAJE_WIN As Byte = 30

' Datos de los apostadores.
Private Type tUserGamble
    ApuestaIndex As Byte
    Name As String
    Dsp As Long
    Gld As Long
End Type

Private Type tApuestas
    ' Ya hay una apuesta en curso
    Run As Boolean
    
    ' Una breve descripción hecha por el Game Master.
    desc As String
    
    ' ¿A quien le vamos a apostar?
    Apuestas() As String
    
    ' ¿Quienes son los apostadores y su información mínima?
    Users() As tUserGamble
    
    
    ' ¿Qué podemos apostar?
    GldAcumulado As Long
    DspAcumulado As Long
    
    
    ' Tiempo para que finalice la oportunidad de apostar.
    TimeFinish As Long
End Type

Public GambleSystem As tApuestas
Private Sub ResetGamble()
          Dim LoopC As Integer
          
          ' Reset the gamble
          
10        With GambleSystem
20            .Run = False
30            .desc = vbNullString
              
40            For LoopC = LBound(.Users()) To UBound(.Users())
50                .Users(LoopC).Dsp = 0
60                .Users(LoopC).Gld = 0
70                .Users(LoopC).Name = vbNullString
80            Next LoopC
              
90        End With
          
          
End Sub
Private Function strGamble(ByRef Apuestas() As String) As String
          Dim LoopC As Integer
          
   On Error GoTo strGamble_Error

10        For LoopC = LBound(Apuestas()) To UBound(Apuestas())
20            strGamble = strGamble & Apuestas(LoopC) & " , "
30        Next LoopC

40        strGamble = mid$(strGamble, 1, Len(strGamble) - 3)

   On Error GoTo 0
   Exit Function

strGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure strGamble of Módulo mApuestas in line " & Erl
End Function
Public Sub NewGamble(ByVal UserIndex As Integer, _
                        ByVal desc As String, _
                        ByVal TimeFinish As Long, _
                        ByVal AmountGamble As Byte, _
                        ByRef Apuestas() As String)
                              
   On Error GoTo NewGamble_Error

10        With GambleSystem
20            If Not EsGM(UserIndex) Then Exit Sub
              
30            If .Run Then
40                WriteConsoleMsg UserIndex, "Ya hay una apuesta en curso. Finaliza la que hay para realizar otra.", FontTypeNames.FONTTYPE_INFO
50                Exit Sub
60            End If
              
              
70            .Apuestas = Apuestas
80            .desc = desc
90            .Run = True
100           .TimeFinish = TimeFinish
              
110           ReDim .Users(0 To AmountGamble) As tUserGamble
              
120           SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Apuesta abierta» Las apuestas han sido abiertas. Podrás apostar por los siguientes usuarios/parejas: " & strGamble(.Apuestas) & "." & vbCrLf & "TIPEA /APUESTAS para saber más. ¡Tendrán " & .TimeFinish & " minutos para apostar!", FontTypeNames.FONTTYPE_GUILD)
          
130       End With

   On Error GoTo 0
   Exit Sub

NewGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure NewGamble of Módulo mApuestas in line " & Erl
                   
End Sub
Public Sub AddRewardGamble(ByVal UserName As String, ByVal Gld As Long, ByVal Dsp As Long)
          Dim UserIndex As Integer
          Dim DspObj As Obj
          Dim DspBov As Long
          Dim GldTemp As Long
          Dim LoopC As Integer
          Dim Entregado As Boolean
          
   On Error GoTo AddRewardGamble_Error

10        UserIndex = NameIndex(UserName)
          
20        If UserIndex > 0 Then
30            With UserList(UserIndex)
40                .Stats.Gld = .Stats.Gld + Gld
                  
50                WriteUpdateGold UserIndex
60                If Dsp > 0 Then
70                    DspObj.Amount = Dsp
80                    DspObj.ObjIndex = 880
                      
90                    If Not MeterItemEnInventario(UserIndex, DspObj) Then
                          
100                   End If
110               End If
                  
                  
120               WriteConsoleMsg UserIndex, "¡Has apostado bien! Has ganado!", FontTypeNames.FONTTYPE_INFO
                  
130           End With
140       Else
150           GldTemp = val(GetVar(CharPath & UCase$(UserName) & ".chr", "STATS", "GLD"))
160           WriteVar CharPath & UCase$(UserName) & ".chr", "STATS", "GLD", GldTemp + Gld
              
              
170           If Dsp > 0 Then
180               Entregado = False
190               For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
200                   If GetVar(CharPath & UCase$(UserName) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC) = "0-0" Then
210                       WriteVar CharPath & UCase$(UserName) & ".chr", "BANCOINVENTORY", "OBJ" & LoopC, "880-" & Dsp
220                       Entregado = True
230                       Exit For
240                   End If
250               Next LoopC
                  
260               If Entregado = False Then
                  
270               End If
280           End If
290       End If

   On Error GoTo 0
   Exit Sub

AddRewardGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AddRewardGamble of Módulo mApuestas in line " & Erl
End Sub
Public Sub CancelGamble(ByVal UserIndex As Integer)
          Dim LoopC As Integer
          Dim tUser As Integer
          Dim ObjDsp As Obj
          
          ' Al cancer una apuesta debemos devolver todo a los personajes
   On Error GoTo CancelGamble_Error

10        With GambleSystem
20            If Not EsGM(UserIndex) Then Exit Sub
              
30            For LoopC = LBound(.Users()) To UBound(.Users())
40                If .Users(LoopC).Name <> vbNullString Then
50                    AddRewardGamble .Users(LoopC).Name, .Users(LoopC).Gld, .Users(LoopC).Dsp
60                End If
                  
70            Next LoopC
              
80            ResetGamble
          
90        End With

   On Error GoTo 0
   Exit Sub

CancelGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CancelGamble of Módulo mApuestas in line " & Erl
End Sub

Private Function SlotUserGamble(ByVal UserName As String) As Integer
          Dim LoopC As Integer
          
   On Error GoTo SlotUserGamble_Error

10        SlotUserGamble = -1
          
20        For LoopC = LBound(GambleSystem.Users()) To UBound(GambleSystem.Users())
30            If StrComp(GambleSystem.Users(LoopC).Name, UserName) = 0 Then
40                SlotUserGamble = LoopC
50                Exit Function
60            End If
70        Next LoopC

   On Error GoTo 0
   Exit Function

SlotUserGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure SlotUserGamble of Módulo mApuestas in line " & Erl
          
End Function
Private Function NewSlot() As Integer
          Dim LoopC As Integer
          
   On Error GoTo NewSlot_Error

10        NewSlot = -1
          
20        For LoopC = LBound(GambleSystem.Users()) To UBound(GambleSystem.Users())
30            If GambleSystem.Users(LoopC).Name = vbNullString Then
40                NewSlot = LoopC
50                Exit Function
60            End If
70        Next LoopC

   On Error GoTo 0
   Exit Function

NewSlot_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure NewSlot of Módulo mApuestas in line " & Erl
End Function
Public Sub UserGamble(ByVal UserIndex As Integer, _
                        ByVal ApuestaIndex As Byte, _
                        ByVal Gld As Long, _
                        ByVal Dsp As Long)
          
          Dim Slot As Integer
          Dim SlotNew As Byte
          
   On Error GoTo UserGamble_Error

10        With UserList(UserIndex)
20            If Gld = 0 And Dsp = 0 Then Exit Sub
              
30            If GambleSystem.Run = False Then
40                WriteConsoleMsg UserIndex, "No hay apuestas para realizar.", FontTypeNames.FONTTYPE_INFO
50                Exit Sub
60            End If
              
70            If GambleSystem.TimeFinish = 0 Then
80                WriteConsoleMsg UserIndex, "Lo lamentamos, las apuestas han sido cerradas.", FontTypeNames.FONTTYPE_INFO
90                Exit Sub
100           End If
              
110           If .Stats.Gld < Gld Then
120               WriteConsoleMsg UserIndex, "No tienes suficiente oro para apostar", FontTypeNames.FONTTYPE_INFO
130               Exit Sub
140           End If
              
150           If Dsp > 0 Then
160               If Not TieneObjetos(880, Dsp, UserIndex) Then
170                   WriteConsoleMsg UserIndex, "No tienes los suficientes DSP que deseas apostar.", FontTypeNames.FONTTYPE_INFO
180                   Exit Sub
190               End If
200           End If
              
              ' ¿El personaje ya apostó alguna vez?
210           Slot = SlotUserGamble(UCase$(.Name))
              
              ' El personaje NO APOSTO , lo agregamos como NUEVO.
220           If Slot = -1 Then
230               SlotNew = NewSlot
                  
240               If SlotNew = -1 Then
250                   WriteConsoleMsg UserIndex, "Lo lamentamos. No hay más lugar en las apuestas. Mejor suerte para la próxima.", FontTypeNames.FONTTYPE_INFO
260                   Exit Sub
270               End If
                  
280               With GambleSystem.Users(SlotNew)
290                   .Dsp = Dsp
300                   .Gld = Gld
310                   .Name = UCase$(UserList(UserIndex).Name)
320                   .ApuestaIndex = ApuestaIndex
                      
330                   WriteConsoleMsg UserIndex, "Has entrado al sistema de apuestas. Recuerda que apostar compulsivamente es perjudicial para la salud.", FontTypeNames.FONTTYPE_INFO
340               End With
                  
350           Else
                  ' Actualizamos la apuesta del personaje.
360               With GambleSystem.Users(Slot)
370                   .Dsp = .Dsp + Dsp
380                   .Gld = .Gld + Gld
                      
390                   WriteConsoleMsg UserIndex, "Has actualizado tu apuesta. Dsp apostados: " & .Dsp & ". Monedas de oro apostadas: " & .Gld & ".", FontTypeNames.FONTTYPE_WARNING
400                   WriteConsoleMsg UserIndex, "Si es que ganas obtendrás un 25% de bonificación del total apostado. Recibirás " & (.Gld * 1.25) & " monedas de oro y " & (.Dsp * 1.25) & " monedas DSP.", FontTypeNames.FONTTYPE_WARNING
410               End With
420           End If
              
430           .Stats.Gld = .Stats.Gld - Gld
440           WriteUpdateGold UserIndex
              
450           If Dsp > 0 Then
460               Call QuitarObjetos(880, Dsp, UserIndex)
470           End If
              
480       End With

   On Error GoTo 0
   Exit Sub

UserGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserGamble of Módulo mApuestas in line " & Erl
End Sub
Private Function CheckExistGamble(ByVal UserName As String) As Integer
          Dim LoopC As Integer
          
   On Error GoTo CheckExistGamble_Error

10        CheckExistGamble = -1
20        For LoopC = LBound(GambleSystem.Apuestas()) To UBound(GambleSystem.Apuestas())
30            If StrComp(GambleSystem.Apuestas(LoopC), UserName) = 0 Then
40                CheckExistGamble = LoopC
50                Exit Function
60            End If
70        Next LoopC

   On Error GoTo 0
   Exit Function

CheckExistGamble_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckExistGamble of Módulo mApuestas in line " & Erl
End Function
Public Sub UserGambleWin(ByVal UserIndex As Integer, ByVal UserName As String)
          Dim LoopC As Integer
          Dim SlotGamble As Integer
          
   On Error GoTo UserGambleWin_Error

10        With GambleSystem
20            If Not EsGM(UserIndex) Then Exit Sub
              
30            SlotGamble = CheckExistGamble(UCase$(UserName))
              
40            If SlotGamble <> -1 Then
50                For LoopC = LBound(.Users()) To UBound(.Users())
60                    If .Users(LoopC).Name <> vbNullString Then
70                        If .Users(LoopC).ApuestaIndex = SlotGamble Then
80                            AddRewardGamble .Users(LoopC).Name, .Users(LoopC).Gld * 1.25, .Users(LoopC).Dsp * 1.25
90                        End If
100                   End If
110               Next LoopC
                  
120               ResetGamble
130           Else
140               WriteConsoleMsg UserIndex, "¿Vos sos retardado? Pusiste mal el usuario/pareja ganador/a", FontTypeNames.FONTTYPE_INFO
150           End If
              
              
160       End With

   On Error GoTo 0
   Exit Sub

UserGambleWin_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserGambleWin of Módulo mApuestas in line " & Erl
End Sub
