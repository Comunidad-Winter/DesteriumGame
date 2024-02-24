Attribute VB_Name = "mUpdate"
Option Explicit

Public Sub UpdateMd5File()
    On Error GoTo ErrHandler
    
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim i As Integer

    ' Especifica la ruta completa del archivo Md5.txt
    filePath = App.path & "\Md5Classic.txt"

    ' Verifica si el archivo existe
    If FileExists(filePath) Then
        ' Lee el contenido del archivo
        Open filePath For Input As #1
        fileContent = Input$(LOF(1), #1)
        Close #1

        ' Divide el contenido en l�neas
        lines = Split(fileContent, vbLf)

        ' Modifica las l�neas espec�ficas
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "DesteriumHD.exe") > 0 Or InStr(lines(i), "DesteriumClassic.exe") > 0 Then
                ' Encuentra las l�neas que contienen los nombres de los archivos y actualiza el hash
                lines(i) = UpdateHash(lines(i))
            End If
        Next i

        ' Une las l�neas modificadas
        fileContent = Join(lines, vbLf)

        ' Guarda los cambios en el archivo
        Open filePath For Output As #1
        Print #1, fileContent;
        Close #1
        
        MsgBox "El cliente se cerrar� por actualizaci�n obligatoria. Entra nuevamente y lograr�s entrar correctamente.", vbInformation
        prgRun = False
    End If
    
ErrHandler:
    Exit Sub
    
End Sub

Private Function UpdateHash(line As String) As String
    ' Actualiza el hash agregando la fecha en formato num�rico (ejemplo: "hash existente" -> "hash existente 17112023")
    Dim parts() As String
    parts = Split(line, "-")

    If UBound(parts) > 0 And InStr(line, "manifest") = 0 Then
        ' No modifica las l�neas que contienen "MANIFEST"
        
        ' Obt�n la fecha en formato num�rico
        Dim numericDate As String
        numericDate = Format(Now, "ddmmyyyy")

        ' Agrega la fecha al final del hash existente
        UpdateHash = parts(0) & "-" & parts(1) & numericDate
    Else
        ' Si no hay un hash existente o la l�nea contiene "MANIFEST", devuelve la l�nea sin cambios
        UpdateHash = line
    End If
End Function

Private Function FileExists(filePath As String) As Boolean
    ' Verifica si un archivo existe
    On Error Resume Next
    FileExists = (GetAttr(filePath) And vbDirectory) = 0
    On Error GoTo 0
End Function

Private Function MakeArray(ParamArray args() As Variant) As Variant
    ' Funci�n para crear un array a partir de una lista de argumentos
    MakeArray = args
End Function
