Attribute VB_Name = "Mod_recordar_password"
Option Explicit
Public Type Recuperar 'type .
    Password As String
    Nick As String
End Type
Public Recu() As Recuperar 'array de nicks y passwords.
Public RecuPath As String 'El path para no escribir tanto xd
Public MaxRecu As Long 'maximo de nicks que cargamos
Public Const ENCRYPT As Byte = 1 'acciones
Public Const DECRYPT As Byte = 2 'acciones
Public Const MYKEY As String = "ClaveEncrypt1449"  'clave de encriptacion.
 
Public Sub LoadRecup()
10        RecuPath = App.path & "\Recursos\Datos.DS"
20        MaxRecu = Val(GetVar(RecuPath, "INIT", "PJs")) 'cargamos el maximo
30        If MaxRecu > 0 Then
40            ReDim Recu(1 To MaxRecu) ' redimencionamos array.
50        End If
       
          Dim loopX As Long
60        For loopX = 1 To MaxRecu 'Hacemos un bucle y cargamos cada una de las contraseñas y nicks.
70            Recu(loopX).Nick = EncryptString(MYKEY, GetVar(RecuPath, "INIT", "NICK" _
                  & loopX), DECRYPT)
80            Recu(loopX).Password = EncryptString(MYKEY, GetVar(RecuPath, "INIT", _
                  "PASS" & loopX), DECRYPT)
90        Next loopX
End Sub
 
Public Function StringIsRecup(ByVal Nombre As String) As Boolean ' no se usa, pueden borrarlo.
          Dim loopX As Long
10        For loopX = 1 To MaxRecu
20            If UCase$(Nombre) = UCase$(Recu(loopX).Nick) Then
30                StringIsRecup = True
40                Exit Function
50            End If
60        Next loopX
70        StringIsRecup = False
End Function
 
Public Function EncryptString(UserKey As String, Text As String, Action As _
    Single) As String
          Dim UserKeyX As String
          Dim temp     As Integer
          Dim Times    As Integer
          Dim i        As Integer
          Dim j        As Integer
          Dim n        As Integer
          Dim rtn      As String
         
       
10        n = Len(UserKey)
20        ReDim UserKeyASCIIS(1 To n)
30        For i = 1 To n
40            UserKeyASCIIS(i) = Asc(mid$(UserKey, i, 1))
50        Next
             
60        ReDim TextASCIIS(Len(Text)) As Integer
70        For i = 1 To Len(Text)
80            TextASCIIS(i) = Asc(mid$(Text, i, 1))
90        Next
         
100       If Action = ENCRYPT Then
110          For i = 1 To Len(Text)
120              j = IIf(j + 1 >= n, 1, j + 1)
130              temp = TextASCIIS(i) + UserKeyASCIIS(j)
140              If temp > 255 Then
150                 temp = temp - 255
160              End If
170              rtn = rtn + Chr$(temp)
180          Next
190       ElseIf Action = DECRYPT Then
200          For i = 1 To Len(Text)
210              j = IIf(j + 1 >= n, 1, j + 1)
220              temp = TextASCIIS(i) - UserKeyASCIIS(j)
230              If temp < 0 Then
240                 temp = temp + 255
250              End If
260              rtn = rtn + Chr$(temp)
270          Next
280       End If
         
290       EncryptString = rtn
End Function
 
 
Public Sub SaveRecu(ByVal Name As String, ByVal pass As String)
10        MaxRecu = MaxRecu + 1
20        Name = EncryptString(MYKEY, Name, ENCRYPT)
30        pass = EncryptString(MYKEY, pass, ENCRYPT)
40        Call WriteVar(RecuPath, "INIT", "PJs", MaxRecu)
50        Call WriteVar(RecuPath, "INIT", "NICK" & MaxRecu, Name)
60        Call WriteVar(RecuPath, "INIT", "PASS" & MaxRecu, pass)
       
70        ReDim Recu(1 To MaxRecu) ' redimencionamos el array.
80        Recu(MaxRecu).Nick = EncryptString(MYKEY, Name, DECRYPT) 'Lo desencriptamos y lo guardamos en memoria.
90        Recu(MaxRecu).Password = EncryptString(MYKEY, pass, DECRYPT) 'Idem password.
End Sub

