Attribute VB_Name = "mReading"
Option Explicit


Sub Load_Main()
    IniPath = App.Path & "\RECURSOS\INIT\"
    
    
End Sub
Sub Read_Body()

    Dim N            As Integer

    Dim i            As Long

    N = FreeFile()
    Open IniPath & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim MisCuerpos(1 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
    Next i
    
    Close #N
    
    FrmMain.lblInfo.Caption = "Cuerpos Cargados..."
End Sub

Sub Write_Body()

    Dim N            As Integer

    Dim i            As Long

    N = FreeFile()
    Open IniPath & "PersonajesNew.ind" For Binary Access Write As #N
    
    'cabecera
    Put #N, , MiCabecera
    
    'num de cabezas
    Put #N, , NumCuerpos
    
    For i = 1 To NumCuerpos
        Put #N, , MisCuerpos(i)
    Next i
    
    Close #N
    
    FrmMain.lblInfo.Caption = "Personajes.ind generado..."
End Sub
