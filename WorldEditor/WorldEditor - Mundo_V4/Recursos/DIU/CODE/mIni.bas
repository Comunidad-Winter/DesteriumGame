Attribute VB_Name = "mIni"
Option Explicit

Sub Ini_Load_Body()
    
    Dim N            As Integer

    Dim i            As Long
    Dim B As Long
    Dim Temp As String
    
    Dim Manager As IniManager
    Set Manager = New IniManager
    
    Call Manager.Initialize(IniPath & "Personajes.ini")
    
    NumCuerpos = Val(Manager.GetValue("INIT", "NumBodies"))
    
    ReDim MisCuerpos(1 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        With MisCuerpos(i)
            
            .HeadOffsetX = Val(Manager.GetValue("BODY" & i, "HeadOffSetX"))
            .HeadOffsetY = Val(Manager.GetValue("BODY" & i, "HeadOffSetY"))
             
            For B = 1 To 4
                .Body(B) = Val(Manager.GetValue("BODY" & i, "Walk" & B))
                .BodyOffsetX(B) = Val(Manager.GetValue("BODY" & i, "BodyOffSetX" & B))
                .BodyOffsetY(B) = Val(Manager.GetValue("BODY" & i, "BodyOffSetY" & B))
            Next B
        

        End With
    Next i
    
    Set Manager = Nothing
    FrmMain.lblInfo.Caption = "Personajes.ini cargado..."
End Sub
Sub Ini_Generate_Body()
    
    Dim N            As Integer

    Dim i            As Long
    Dim B As Long
    
    Dim Manager As IniManager
    
    Set Manager = New IniManager
    
    For i = 1 To NumCuerpos
        With MisCuerpos(i)
            Call Manager.ChangeValue("BODY" & i, "HeadOffSetX", CStr(.HeadOffsetX))
            Call Manager.ChangeValue("BODY" & i, "HeadOffSetY", CStr(.HeadOffsetY))
            
             '
            For B = 1 To 4
                Call Manager.ChangeValue("BODY" & i, "Walk" & B, CStr(.Body(B)) & "     ' " & Heading_To_String(B))
                Call Manager.ChangeValue("BODY" & i, "BodyOffSetX" & B, CStr(.BodyOffsetX(B)))
                Call Manager.ChangeValue("BODY" & i, "BodyOffSetY" & B, CStr(.BodyOffsetY(B)))
            Next B
        

        End With
    Next i
    
    Call Manager.ChangeValue("INIT", "NumBodies", NumCuerpos)
    Call Manager.DumpFile(IniPath & "Personajes.ini")
    
    Set Manager = Nothing
    FrmMain.lblInfo.Caption = "Personajes.ini generados..."
End Sub
