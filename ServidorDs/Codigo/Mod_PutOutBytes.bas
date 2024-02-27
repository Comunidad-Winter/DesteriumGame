Attribute VB_Name = "Mod_PutOutBytes"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Code By Miqueas
' 05/10/14
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Option Explicit
 
' Server send, recieve bytes
 
Private prvStaticSendBytes As Long, prvStaticRecieveBytes As Long
             
Public Sub set_ByteRecieve(ByVal data As Long)
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
 
     Dim lBytePut As Long
 
     lBytePut = data
     prvStaticRecieveBytes = prvStaticRecieveBytes + lBytePut
 
End Sub
 
Public Sub set_ByteSend(ByVal data As String)
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
 
     Dim lByteOut As Long
 
     lByteOut = Len(data)
     prvStaticSendBytes = prvStaticSendBytes + lByteOut
 
End Sub
 
Private Function get_ByteRecieve() As Long
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
 
     get_ByteRecieve = prvStaticRecieveBytes
 
End Function
 
Private Function get_ByteSend() As Long
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
 
     get_ByteSend = prvStaticSendBytes
 
End Function
 
Public Sub PutInfoBytes()
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                         
     If (frmMain.Visible = True) Then
           frmMain.lblBytesSalida.Caption = "Bytes Salida: " & Round(Mod_PutOutBytes.get_ByteSend / 1024, 3) & "kb/s"
           frmMain.lblBytesEntrada.Caption = "Bytes Entrada: " & Round(Mod_PutOutBytes.get_ByteRecieve / 1024, 3) & "kb/s"
     End If
 
     Call ResetStatBytes
 
End Sub
 
Private Sub ResetStatBytes()
 
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      ' Code By Miqueas
     ' 05/10/14
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
     prvStaticRecieveBytes = 0
     prvStaticSendBytes = 0
 
End Sub


