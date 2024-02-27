Attribute VB_Name = "mod_motd"
Sub LoadMotd()
'Author: SHAK
'Creado: 22/02/2014
'Última modificación: -
'**********************
 
    Dim i As Integer
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    Dim bold As Byte
    Dim Italy As Byte
   
    MaxLines = Val(GetVar(App.path & "\INIT\Motd.ini", "INIT", "NumLines"))
   
    ReDim MOTD(1 To MaxLines)
    For i = 1 To MaxLines
        MOTD(i).Texto = GetVar(App.path & "\INIT\Motd.ini", "Motd", "Line" & i)
       
        r = ReadField(2, MOTD(i).Texto, Asc("~"))
        g = ReadField(3, MOTD(i).Texto, Asc("~"))
        b = ReadField(4, MOTD(i).Texto, Asc("~"))
        bold = ReadField(5, MOTD(i).Texto, Asc("~"))
        Italy = ReadField(6, MOTD(i).Texto, Asc("~"))
       
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(MOTD(i).Texto, InStr(1, MOTD(i).Texto, "~") - 1), r, g, b, IIf(bold > 0, True, False), IIf(Italy > 0, True, False), True)
       
    Next i
   
End Sub
