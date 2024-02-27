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
    UseInvalid As Byte
    ModifyTime As Long
End Type

Public Const MAX_INTERVAL As Byte = 2
Public Intervalos(0 To MAX_INTERVAL) As tInterval
Public DefaultIntervalos(0 To MAX_INTERVAL) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
  
' Intervalos Default
Public Sub SetIntervalos()
    
    DefaultIntervalos(eInterval.iUseItem) = 2
    DefaultIntervalos(eInterval.iUseItemClick) = 5
    DefaultIntervalos(eInterval.iUseSpell) = 1000
    
End Sub

' Restamos tiempo de los intervalos
Public Sub LoopInterval()

    Dim A As Long
    
    For A = 0 To MAX_INTERVAL
        If Intervalos(A).Modify > 0 Then Intervalos(A).Modify = Intervalos(A).Modify - 1
    Next A
    
End Sub

' Chequeamos si un intervalo sigue descontando
Public Function CheckInterval(ByVal iType As eInterval) As Boolean
    
    Dim Time As Long
    
    Time = timeGetTime - Intervalos(iUseItemClick).ModifyTime
    
    If (Time) <= 60 Then
        Exit Function
    End If
    
    'If Intervalos(iType).Modify > 0 Then Exit Function
    
    CheckInterval = True
    SetIntervalos
    
End Function

' Asignamos al intervalo el tiempo para descontarlo
Public Sub AssignedInterval(ByVal iType As eInterval)
                                
    Intervalos(iType).Modify = DefaultIntervalos(iType)
    Intervalos(iType).ModifyTime = timeGetTime
    
    
End Sub



