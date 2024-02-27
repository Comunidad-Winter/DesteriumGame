Attribute VB_Name = "Foto_Denuncias"
Option Explicit
 
'declare Constants and variables
 
'One second of interval.
'Changes to public of the use in call to capture picture.
Public Const FotoD_MAX_INTERVAL                As Long = 60000
'Here save last interval of photo report.
Private FotoD_LastIN                             As Long
'Number of last insult to the Array.
Private Const FOTOD_INSULTMAX                    As Byte = 36
'Container array of insult list.
Private FotoD_InsultList(1 To FOTOD_INSULTMAX)   As String
 
Sub FotoD_Initialize()
       
       
10    FotoD_InsultList(1) = "PT"
20    FotoD_InsultList(2) = "MANCO"
30    FotoD_InsultList(3) = "ASCO"
40    FotoD_InsultList(4) = "ASKO"
50    FotoD_InsultList(5) = "NW"
60    FotoD_InsultList(6) = "FRACA"
70    FotoD_InsultList(7) = "FRAKA"
80    FotoD_InsultList(8) = "PETE"
90    FotoD_InsultList(9) = "DAS PENA"
100   FotoD_InsultList(10) = "KB"
110   FotoD_InsultList(11) = "KABE"
120   FotoD_InsultList(12) = "CABE"
130   FotoD_InsultList(13) = "KBIO"
140   FotoD_InsultList(14) = "CABIO"
150   FotoD_InsultList(15) = "TAS EN LA RUINA"
160   FotoD_InsultList(16) = "PUTO"
170   FotoD_InsultList(17) = "PUTA"
180   FotoD_InsultList(18) = "PAJERO"
190   FotoD_InsultList(19) = "PAJERA"
200   FotoD_InsultList(20) = "CONCHA"
210   FotoD_InsultList(21) = "TU MADRE"
220   FotoD_InsultList(22) = "TU MAMA"
230   FotoD_InsultList(23) = "HIJO"
       
240   FotoD_InsultList(24) = _
          "LA PUTA QUE TE RE MIL PARIO PEDAZO DE FRACA HIJO DE PUTA DAS ASKO AJAJJAJAJAJA"
       
250   FotoD_InsultList(25) = "SORETE"
260   FotoD_InsultList(26) = "MIERDA"
270   FotoD_InsultList(27) = "PELOTUDO"
280   FotoD_InsultList(28) = "MOGOLICO"
290   FotoD_InsultList(29) = "RETRASADO"
300   FotoD_InsultList(30) = "ENFERMO"
310   FotoD_InsultList(31) = "DAWN"
320   FotoD_InsultList(32) = "SIMIO"
330   FotoD_InsultList(33) = "NO TENES VIDA"
340   FotoD_InsultList(34) = "CAGADA"
350   FotoD_InsultList(35) = "VIRGEN"
360   FotoD_InsultList(36) = "PENE"
       
370   FotoD_LastIN = 60001
       
End Sub
 
Sub FotoD_Capturar(refString As String)

      Dim loopX       As Long
      Dim sendString  As String
       
      'Whenever we initialize the variable is null.
10    sendString = vbNullString
       
20        For loopX = 1 To LastChar
             
30            With charlist(loopX)
             
                  'It's char in pc area?
40                If FotoD_CharInPCArea(loopX) Then
                      'Analize LastDialog
50                    If FotoD_DialogIsInsult(loopX) Then
                              'Save charDialogs and NickName here.
60                        sendString = sendString & "," & .Nombre & " : " & _
                              .LastDialog
70                    End If
80                End If
90            End With
             
100       Next loopX
       
110   refString = sendString
       
120   If refString <> vbNullString Then
130   FotoD_LastIN = GetTickCount
140   End If
       
End Sub
 
Sub FotoD_SaveLastDialog(ByVal CharIndex As Integer, ByVal DialoG As String)
       

       
10    With charlist(CharIndex)
20    If .Nombre = vbNullString Then Exit Sub
30    .LastDialog = DialoG
       
40    End With
       
End Sub
 
Sub FotoD_RemoveLastDialog(ByVal CharIndex As Integer)
       

10    If charlist(CharIndex).Nombre = vbNullString Then Exit Sub
20    charlist(CharIndex).LastDialog = vbNullString
       
End Sub
 
Function FotoD_DialogIsInsult(ByVal CharIndex As Integer) As Boolean
       

      Dim loopX      As Long
       
10        For loopX = 1 To UBound(FotoD_InsultList())
             
              'Analize charDialogs
             
20            If InStr(1, UCase$(charlist(CharIndex).LastDialog), _
                  FotoD_InsultList(loopX)) Then
                  'Insult are found? returns true and exit function!
30                FotoD_DialogIsInsult = True
40                Exit Function
50            End If
             
60        Next loopX
       
70    FotoD_DialogIsInsult = False
       
End Function
 
Function FotoD_CanSend() As Boolean

       
10    If FotoD_LastIN = 60001 Then FotoD_CanSend = True: Exit Function
       
20    FotoD_CanSend = (GetTickCount - FotoD_LastIN > FotoD_MAX_INTERVAL)
       
End Function
 
Function FotoD_CharInPCArea(ByVal CharIndex As Integer) As Boolean

       
10        With charlist(CharIndex)
         
20            FotoD_CharInPCArea = (.Pos.X > (UserPos.X - MinXBorder) And .Pos.X < _
                  (UserPos.X + MinXBorder) And .Pos.Y > (UserPos.Y - MinYBorder) And _
                  .Pos.Y < (UserPos.Y + MinYBorder))
             
30        End With
         
End Function
