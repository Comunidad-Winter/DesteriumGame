Attribute VB_Name = "mWork"
' Módulo de trabajo DesteriumAO Exodo III

Option Explicit

Enum eItemsConstruibles_Subtipo

    eArmadura = 1
    eCasco = 2
    eEscudo = 3
    eArmas = 4
    eMuniciones = 5
    eEmbarcaciones = 6
    eObjetoMagico = 7
    eInstrumento = 8

End Enum

Public Type eItemsConstruibles

    ObjIndex As Integer
    SubTipo As eItemsConstruibles_Subtipo

End Type

Public ObjBlacksmith()      As eItemsConstruibles

Public ObjBlacksmith_Amount As Integer

Public ObjCarpinter()       As Integer

Public ObjCarpinter_Amount  As Integer

Public Sub Crafting_Reset()
    ObjBlacksmith_Amount = 0
    ReDim ObjBlacksmith(0) As eItemsConstruibles

End Sub

' # Agregamos el objeto a la lista de herrería
Public Sub Crafting_BlackSmith_Add(ByVal ObjIndex As Integer)

    '<EhHeader>
    On Error GoTo Crafting_BlackSmith_Add_Err

    '</EhHeader>

    ObjBlacksmith_Amount = ObjBlacksmith_Amount + 1
    ReDim Preserve ObjBlacksmith(0 To ObjBlacksmith_Amount) As eItemsConstruibles
    
    ObjBlacksmith(ObjBlacksmith_Amount).ObjIndex = ObjIndex
    ObjBlacksmith(ObjBlacksmith_Amount).SubTipo = Set_Subtype_Object(ObjIndex)
    
    '<EhFooter>
    Exit Sub

Crafting_BlackSmith_Add_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mWork.Crafting_BlackSmith_Add " & "at line " & Erl
        
    '</EhFooter>
End Sub

' # Agregamos el SubTipo del objeto, para luego separarlos de forma mas rapida
Public Function Set_Subtype_Object(ByVal ObjIndex As Integer) As eItemsConstruibles_Subtipo

    '<EhHeader>
    On Error GoTo Set_Subtype_Object_Err

    '</EhHeader>
    
    With ObjData(ObjIndex)
        
        Select Case .OBJType
        
            Case eOBJType.otarmadura
                Set_Subtype_Object = eArmadura
                
            Case eOBJType.otAnillo, eOBJType.otMagic, eOBJType.oteffect
                Set_Subtype_Object = eObjetoMagico
                
            Case eOBJType.otescudo
                Set_Subtype_Object = eEscudo
                
            Case eOBJType.otcasco
                Set_Subtype_Object = eCasco
                
            Case eOBJType.otWeapon
                Set_Subtype_Object = eArmas
                
            Case eOBJType.otFlechas
                Set_Subtype_Object = eMuniciones
                
            Case eOBJType.otBarcos
                Set_Subtype_Object = eEmbarcaciones
                
            Case eOBJType.otInstrumentos
                Set_Subtype_Object = eInstrumento
                
            Case Else
                Set_Subtype_Object = eObjetoMagico

        End Select
        
    End With

    '<EhFooter>
    Exit Function

Set_Subtype_Object_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mWork.Set_Subtype_Object " & "at line " & Erl
        
    '</EhFooter>
End Function

' # Comprueba de tener los recursos necesarios que necesita el objeto para ser creado/mejorado
Public Function Crafting_Checking_Object(ByVal UserIndex As Integer, _
                                         ByVal QuestIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo Crafting_Checking_Object_Err

    '</EhHeader>
    Dim A    As Long

    Dim Temp As String
    
    Crafting_Checking_Object = True
    
    With QuestList(QuestIndex)

        For A = 1 To .RequiredOBJs

            If Not TieneObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex) Then
               
                Crafting_Checking_Object = False
                Exit Function

            End If

        Next A
        
    End With
    
    '<EhFooter>
    Exit Function

Crafting_Checking_Object_Err:
    LogError Err.description & vbCrLf & "in Crafting_Checking_Object " & "at line " & Erl

    '</EhFooter>
End Function

' # Quita los recursos necesarios para la construcción/mejora del objeto.
Public Sub Crafting_Remove_Object(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)

    '<EhHeader>
    On Error GoTo Crafting_Remove_Object_Err

    '</EhHeader>
    
    Dim A As Long
    
    With QuestList(QuestIndex)

        For A = 1 To .RequiredOBJs
            Call QuitarObjetos(.RequiredObj(A).ObjIndex, .RequiredObj(A).Amount, UserIndex)
        Next A

    End With
    
    '<EhFooter>
    Exit Sub

Crafting_Remove_Object_Err:
    LogError Err.description & vbCrLf & "in Crafting_Remove_Object " & "at line " & Erl

    '</EhFooter>
End Sub

' # Fundicion del Objeto
Private Sub Crafting_Fundition(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '<EhHeader>
    On Error GoTo Crafting_Fundition_Err

    '</EhHeader>
    
    Dim A    As Long

    Dim Temp As Long

    Dim Obj  As Obj
    
    If ConfigServer.ModoCrafting = 0 Then
        Call WriteConsoleMsg(UserIndex, "El servidor no admite la fundición de objetos.", FontTypeNames.FONTTYPE_INFORED)
        Exit Sub

    End If
    
    With UserList(UserIndex)
        
        For A = 1 To ObjData(ObjIndex).Upgrade.RequiredCant
            Temp = ObjData(ObjIndex).Upgrade.Required(A).Amount * 0.3
            
            If Temp > 0 Then
                Obj.Amount = Temp
                Obj.ObjIndex = ObjData(ObjIndex).Upgrade.Required(A).ObjIndex
                
                If Not MeterItemEnInventario(UserIndex, Obj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj)

                End If

            End If

        Next A
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(eSound.sConstruction, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
        UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta

        If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then UserList(UserIndex).Reputacion.PlebeRep = MAXREP

    End With

    '<EhFooter>
    Exit Sub

Crafting_Fundition_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mWork.Crafting_Fundition " & "at line " & Erl
        
    '</EhFooter>
End Sub

