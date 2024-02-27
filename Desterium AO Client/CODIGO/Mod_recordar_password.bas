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
    RecuPath = App.path & "\Recursos\Datos.DS"
    MaxRecu = Val(GetVar(RecuPath, "INIT", "PJs")) 'cargamos el maximo
    If MaxRecu > 0 Then
        ReDim Recu(1 To MaxRecu) ' redimencionamos array.
    End If
 
    Dim loopX As Long
    For loopX = 1 To MaxRecu 'Hacemos un bucle y cargamos cada una de las contraseñas y nicks.
        Recu(loopX).Nick = EncryptString(MYKEY, GetVar(RecuPath, "INIT", "NICK" & loopX), DECRYPT)
        Recu(loopX).Password = EncryptString(MYKEY, GetVar(RecuPath, "INIT", "PASS" & loopX), DECRYPT)
    Next loopX
End Sub
 
Public Function StringIsRecup(ByVal Nombre As String) As Boolean ' no se usa, pueden borrarlo.
    Dim loopX As Long
    For loopX = 1 To MaxRecu
        If UCase$(Nombre) = UCase$(Recu(loopX).Nick) Then
            StringIsRecup = True
            Exit Function
        End If
    Next loopX
    StringIsRecup = False
End Function
 
Public Function EncryptString( _
    UserKey As String, Text As String, Action As Single _
    ) As String
    Dim UserKeyX As String
    Dim temp     As Integer
    Dim Times    As Integer
    Dim i        As Integer
    Dim j        As Integer
    Dim n        As Integer
    Dim rtn      As String
   
 
    n = Len(UserKey)
    ReDim UserKeyASCIIS(1 To n)
    For i = 1 To n
        UserKeyASCIIS(i) = Asc(mid$(UserKey, i, 1))
    Next
       
    ReDim TextASCIIS(Len(Text)) As Integer
    For i = 1 To Len(Text)
        TextASCIIS(i) = Asc(mid$(Text, i, 1))
    Next
   
    If Action = ENCRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           temp = TextASCIIS(i) + UserKeyASCIIS(j)
           If temp > 255 Then
              temp = temp - 255
           End If
           rtn = rtn + Chr$(temp)
       Next
    ElseIf Action = DECRYPT Then
       For i = 1 To Len(Text)
           j = IIf(j + 1 >= n, 1, j + 1)
           temp = TextASCIIS(i) - UserKeyASCIIS(j)
           If temp < 0 Then
              temp = temp + 255
           End If
           rtn = rtn + Chr$(temp)
       Next
    End If
   
    EncryptString = rtn
End Function
 
 
Public Sub SaveRecu(ByVal Name As String, ByVal pass As String)
    MaxRecu = MaxRecu + 1
    Name = EncryptString(MYKEY, Name, ENCRYPT)
    pass = EncryptString(MYKEY, pass, ENCRYPT)
    Call WriteVar(RecuPath, "INIT", "PJs", MaxRecu)
    Call WriteVar(RecuPath, "INIT", "NICK" & MaxRecu, Name)
    Call WriteVar(RecuPath, "INIT", "PASS" & MaxRecu, pass)
 
    ReDim Recu(1 To MaxRecu) ' redimencionamos el array.
    Recu(MaxRecu).Nick = EncryptString(MYKEY, Name, DECRYPT) 'Lo desencriptamos y lo guardamos en memoria.
    Recu(MaxRecu).Password = EncryptString(MYKEY, pass, DECRYPT) 'Idem password.
End Sub

