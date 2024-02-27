Attribute VB_Name = "mCofres"
Option Explicit

Private Type tObjCofre
    ObjIndex As Integer
    Amount As Long
    Probability As Byte
End Type

Private Type tCofres
    ObjIndex As Integer
    NumObj As Integer
    Obj() As tObjCofre
    Probability As Byte
End Type

Private NumCofres As Byte
Private Cofres() As tCofres

Public Sub LoadCofres()
    Dim A As Long, B As Long
    Dim Read As clsIniManager
    Dim Temp As String
    
    
    Set Read = New clsIniManager
    
    Read.Initialize DatPath & "COFRES.DAT"
    
    NumCofres = val(Read.GetValue("INIT", "NUMCOFRES"))
    
    ReDim Cofres(1 To NumCofres) As tCofres
    
    For A = 1 To NumCofres
        With Cofres(A)
            .NumObj = val(Read.GetValue("COFRE" & A, "NumObj"))
            .Probability = val(Read.GetValue("COFRE" & A, "Probabilidad"))
            .ObjIndex = val(Read.GetValue("COFRE" & A, "ObjIndex"))
            
            ReDim .Obj(1 To .NumObj) As tObjCofre
            
            For B = 1 To .NumObj
                Temp = Read.GetValue("COFRE" & A, "OBJ" & B)
                
                
                .Obj(B).ObjIndex = val(ReadField(1, Temp, Asc("-")))
                .Obj(B).Amount = val(ReadField(2, Temp, Asc("-")))
                .Obj(B).Probability = val(ReadField(3, Temp, Asc("-")))
            Next B
        End With
    Next A
End Sub

Public Sub UsuarioPescaCofres(ByVal UserIndex As Integer)
    
    Dim Random As Integer
    With UserList(UserIndex)
        ' Requisitos?¿¡!
        If .Invent.WeaponEqpObjIndex <> CAÑA_COFRES Then Exit Sub
        If .Stats.UserSkills(eSkill.Pesca) < 100 Then Exit Sub
        If .Pos.map <> 34 Then Exit Sub
        
        Random = RandomNumber(1, 100)
        'WriteConsoleMsg UserIndex, "Random: " & Random, FontTypeNames.FONTTYPE_WARNING
        
        If Random <= 99 Then Exit Sub
        
        Dim A As Long, B As Long
        Dim Obj As Obj
        
        For A = 1 To NumCofres
        
            If RandomNumber(1, 100) <= Cofres(A).Probability Then
                
                Obj.ObjIndex = Cofres(A).ObjIndex
                Obj.Amount = 1
                
                If Not MeterItemEnInventario(UserIndex, Obj) Then
                    Call TirarItemAlPiso(.Pos, Obj)
                End If
                
                WriteConsoleMsg UserIndex, "¡Has pescado un cofre!", FontTypeNames.FONTTYPE_WARNING
                
                Exit For
            End If
        Next A
    End With
    
End Sub


Public Sub UsuarioUsaCofre(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    ' Programado para anti boludos que toquen el codigo despues de mi xD
    
    Dim A As Long
    Dim Value As Long
    Dim cant As Byte
    Dim Obj As Obj
    
    For A = 1 To NumCofres
        If Cofres(A).ObjIndex = ObjIndex Then
            Value = A
            Exit For
        End If
    Next A
    
    
    If Value > 0 Then
        With Cofres(Value)
            For A = 1 To .NumObj
                If cant = 3 Then Exit For
                
                Randomize
                
                If RandomNumber(1, 100) <= .Obj(A).Probability Then
                    
                    Obj.ObjIndex = .Obj(A).ObjIndex
                    Obj.Amount = .Obj(A).Amount
                    
                    If Not MeterItemEnInventario(UserIndex, Obj) Then
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj)
                    End If
                        
                    
                        
                    cant = cant + 1
                        
                End If
                
            Next A
            
            Call QuitarObjetos(ObjIndex, 1, UserIndex)
        End With
    
    End If
    
End Sub

'[Init]
'NumCofres = 1

'[COFRE1]
'NumObj = 2
'Obj1 = 1037 - 2 - 1
'Obj2 = 402 - 1 - 50
