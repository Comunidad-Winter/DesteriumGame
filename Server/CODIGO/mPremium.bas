Attribute VB_Name = "mPremium"
Option Explicit

Public Type tPremium

    ObjIndex As Integer
    Amount As Integer
    RequiredObj As Long
    RequiredAmount As Long

End Type

Public Premiums()   As tPremium

Public Premium_Last As Integer

Public Sub Premiums_Load()

    '<EhHeader>
    On Error GoTo Premiums_Load_Err

    '</EhHeader>
    Dim Read As clsIniManager

    Dim A    As Long

    Dim Temp As String
    
    Set Read = New clsIniManager
    
    Read.Initialize (DatPath & "PREMIUM.DAT")
    
    Premium_Last = val(Read.GetValue("INIT", "LAST"))
    
    ReDim Premiums(1 To Premium_Last) As tPremium
    
    For A = 1 To Premium_Last

        With Premiums(A)
            Temp = Read.GetValue("LIST", A)
            .ObjIndex = val(ReadField(1, Temp, Asc("-")))
            .Amount = val(ReadField(2, Temp, Asc("-")))
            .RequiredAmount = val(ReadField(3, Temp, Asc("-")))
            .RequiredObj = 1466

        End With

    Next A
    
    Set Read = Nothing
    '<EhFooter>
    Exit Sub

Premiums_Load_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mPremium.Premiums_Load " & "at line " & Erl
        
    '</EhFooter>
End Sub

