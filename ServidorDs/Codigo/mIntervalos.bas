Attribute VB_Name = "mIntervalos"
' Lautaro Leonel Marino (Shakeño)

Option Explicit

Public Enum eInterval
    iUseItem = 0
    iUseItemClick = 1
    iUseSpell = 2
End Enum

Public Type tInterval
    Default As Long
    Modify As Long
    ModifyTimer As Long
    UseInvalid As Byte
End Type

' Maximo de Intervalos
Public Const MAX_INTERVAL As Byte = 2

' Configuración Inicial de los Intervalos
Public DefaultIntervalos(0 To MAX_INTERVAL) As Long

' Calcular Tiempos
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' Intervalos Default
Public Sub SetInterval()
    
    DefaultIntervalos(eInterval.iUseItem) = 2
    DefaultIntervalos(eInterval.iUseItemClick) = 10
    DefaultIntervalos(eInterval.iUseSpell) = 1000
    
End Sub

' Restamos tiempo de los intervalos
Public Sub LoopInterval(ByVal Userindex As Integer)

    Dim A As Long
    
    With UserList(Userindex)
    
        For A = 0 To MAX_INTERVAL
            If .Intervalos(A).Modify > 0 Then .Intervalos(A).Modify = .Intervalos(A).Modify - 1
        Next A
        
    End With
    
End Sub

' Chequeamos si un intervalo sigue descontando
Public Function CheckInterval(ByVal Userindex As Integer, _
                            ByVal iType As eInterval) As Boolean

        
    ' Primer chequeo ¿El intervalo sigue descontandose?
    If UserList(Userindex).Intervalos(iType).Modify > 0 Then
        Exit Function
    End If
    
        ' Segundo chequeo ?
    If (timeGetTime - UserList(Userindex).Intervalos(iType).Modify) <= 200 Then
        Exit Function
    End If
        

    SetInterval
    CheckInterval = True
    
End Function

' Asignamos al intervalo el tiempo para descontarlo
Public Sub AssignedInterval(ByVal Userindex As Integer, _
                                ByVal iType As eInterval)
                                
    UserList(Userindex).Intervalos(iType).Modify = DefaultIntervalos(iType)
    UserList(Userindex).Intervalos(iType).ModifyTimer = timeGetTime
End Sub

