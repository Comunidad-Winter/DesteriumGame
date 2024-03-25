Attribute VB_Name = "mDrop"
Option Explicit

Public Type tDropData

    ObjIndex As Integer
    Amount(1) As Integer
    Prob As Byte

End Type

Public Type tDrop

    Last As Byte
    Data() As tDropData

End Type

Public DropLast   As Integer

Public DropData() As tDrop

Public Sub Drops_Load()

    '<EhHeader>
    On Error GoTo Drops_Load_Err

    '</EhHeader>
    Dim Manager As clsIniManager

    Dim A       As Long, B As Long

    Dim Temp    As String
    
    Set Manager = New clsIniManager
            
    Dim FilePath As String

    FilePath = Drops_FilePath
    Manager.Initialize (FilePath)
    
    DropLast = val(Manager.GetValue("INIT", "LAST"))
    
    ReDim DropData(1 To DropLast) As tDrop
    
    For A = 1 To DropLast

        With DropData(A)
            .Last = val(Manager.GetValue(A, "LAST"))
            
            ReDim .Data(1 To .Last) As tDropData
            
            For B = 1 To .Last
                Temp = Manager.GetValue(A, B)
                .Data(B).ObjIndex = val(ReadField(1, Temp, 45))
                .Data(B).Prob = val(ReadField(2, Temp, 45))
                .Data(B).Amount(0) = val(ReadField(3, Temp, 45))
                .Data(B).Amount(1) = val(ReadField(4, Temp, 45))
            Next B
        
        End With
            
    Next A

    Manager.DumpFile Drops_FilePath_Client
    Set Manager = Nothing
    
    '<EhFooter>
    Exit Sub

Drops_Load_Err:
    Set Manager = Nothing
    LogError Err.description & vbCrLf & "in ServidorArgentum.mDrop.Drops_Load " & "at line " & Erl
        
    '</EhFooter>
End Sub
