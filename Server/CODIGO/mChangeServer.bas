Attribute VB_Name = "mChangeServer"
Option Explicit

Private Declare Function GetModuleFileName _
                Lib "kernel32" _
                Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                            ByVal lpFileName As String, _
                                            ByVal nSize As Long) As Long

Public Function GetCurrentProcessName() As String

    Dim buffer As String * 260

    Dim length As Long
    
    ' Obtener la ruta del ejecutable actual
    length = GetModuleFileName(0, buffer, Len(buffer))
    
    ' Extraer solo el nombre del archivo del camino
    Dim fullPath As String

    fullPath = TrimNull(Left$(buffer, length))
    
    ' Obtener solo el nombre del archivo del camino
    GetCurrentProcessName = GetFileNameFromPath(fullPath)

End Function

Private Function GetFileNameFromPath(fullPath As String) As String

    ' Extraer el nombre del archivo de una ruta completa
    Dim Pos As Integer

    Pos = InStrRev(fullPath, "\")

    If Pos > 0 Then
        GetFileNameFromPath = mid(fullPath, Pos + 1)
    Else
        GetFileNameFromPath = fullPath

    End If

End Function

Private Function TrimNull(inputString As String) As String

    ' Eliminar caracteres nulos de la cadena
    Dim nullPos As Integer

    nullPos = InStr(inputString, vbNullChar)

    If nullPos > 0 Then
        TrimNull = Left$(inputString, nullPos - 1)
    Else
        TrimNull = inputString

    End If

End Function
