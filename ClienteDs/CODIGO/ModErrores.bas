Attribute VB_Name = "ModErrores"
Option Explicit


Public Sub LogError(Desc As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo ErrHandler

        ' LogError Desc
         
          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.path & "\Pasar_Archivo_Al_Soporte.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & Desc
50        Close #nfile
          
60        Exit Sub

ErrHandler:

End Sub

