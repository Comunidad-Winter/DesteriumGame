Attribute VB_Name = "mSecurity"
' Controles básicos

Option Explicit

' Constants Key Code Packets
Public Const MAX_KEY_PACKETS As Byte = 2
Public Const MAX_KEY_CHANGE As Byte = 30

' Constants Pointers
Public Const MAX_POINTERS As Byte = 2
Public Const LIMIT_POINTER As Byte = 9
Public Const LIMIT_FLOD_POINTER As Byte = 4

' Position of the pointer
Public Enum ePoint
    Point_Spell = 1
    Point_Inv = 2
End Enum

' Cursor X, Y
Public Type tPoint
    X(LIMIT_POINTER) As Long
    Y(LIMIT_POINTER) As Long
    Cant(LIMIT_POINTER) As Byte
End Type

Public Type tPackets
    Value As Byte
    Cant As Byte
End Type

' Key Code of the special Packets
Public Enum eKeyPackets
    Key_UseItem = 0
    Key_UseSpell = 1
    Key_UseWeapon = 2
End Enum




' Change Key Code
Public Sub UpdateKeyPacket(ByVal UserIndex As Integer, _
                            ByVal Packet As eKeyPackets, _
                            Optional ByVal UpdateAll As Boolean = False)
    
    Randomize
    
    With UserList(UserIndex)
        .KeyPackets(Packet).Value = RandomNumber(1, 255)
        .KeyPackets(Packet).Cant = 0
        
        WriteUpdateKey UserIndex, Packet, .KeyPackets(Packet).Value
    End With
End Sub

' Check Key Code
Public Function CheckKeyPacket(ByVal UserIndex As Integer, _
                            ByVal Packet As eKeyPackets, _
                            ByVal KeyPacket As Long) As Boolean
                            
                            
    With UserList(UserIndex)
    
        If .KeyPackets(Packet).Value = 0 Then
            UpdateKeyPacket UserIndex, Packet
            CheckKeyPacket = True
            Exit Function
        End If
        
        If .KeyPackets(Packet).Value <> KeyPacket Then Exit Function
        
        .KeyPackets(Packet).Cant = .KeyPackets(Packet).Cant + 1
        
        If .KeyPackets(Packet).Cant = MAX_KEY_CHANGE Then
            UpdateKeyPacket UserIndex, Packet
        End If
        
    End With
    
    CheckKeyPacket = True
End Function
                        
' Reset Key Code
Public Sub ResetKeyPackets(ByVal UserIndex As Integer)

    Dim A As Long
    
    With UserList(UserIndex)
    
        For A = 0 To MAX_KEY_PACKETS
            .KeyPackets(A).Value = 0
            .KeyPackets(A).Cant = 0
        Next A
        
    End With
End Sub

' Reset Pointer
Public Sub ResetPointer(ByVal UserIndex As Integer, _
                            ByVal Point As ePoint)
    Dim A As Long
    
    With UserList(UserIndex).Pointers(Point)
    
        For A = 0 To LIMIT_POINTER
            .Cant(A) = 0
            .X(A) = 0
            .Y(A) = 0
        Next A
        
    End With
End Sub

Public Sub UpdatePointer(ByVal UserIndex As Integer, _
                        ByVal Point As ePoint, _
                        ByVal X As Long, _
                        ByVal Y As Long)
                        
    Dim A As Integer
    Dim PointerCero As Byte
    
    With UserList(UserIndex).Pointers(Point)
        For A = 0 To LIMIT_POINTER
            
            ' Pointer cero
            If PointerCero = 0 And _
                (.X(A) = 0 And .Y(A) = 0) Then
                
                PointerCero = A
            End If
            
            ' Pointer repetido
            If .X(A) = X And .Y(A) = Y Then
                .Cant(A) = .Cant(A) + 1
                
                ' Cantidad = 4 (Máximo permitido en el Point de igualdad)
                If .Cant(A) = LIMIT_FLOD_POINTER Then
                    LogAntiCheat "El personaje " & UserList(UserIndex).Name & " tuvo una precisión del Point n° " & Point
                    ResetPointer UserIndex, Point
                    Exit Sub
                End If
                
                Exit Sub
            End If
        Next A
        
        If PointerCero = 0 Then
            ResetPointer UserIndex, Point
            .X(0) = X
            .Y(0) = Y
            .Cant(0) = 1
        Else
            .X(PointerCero) = X
            .Y(PointerCero) = Y
            .Cant(PointerCero) = 1
        End If
        
    End With
End Sub
