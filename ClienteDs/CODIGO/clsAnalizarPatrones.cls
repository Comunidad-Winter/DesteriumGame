VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnalizarPatrones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const Limite_Medir As Byte = 4
Private Type tClicks
    LCLista As Long
    intClick(Limite_Medir) As Long
    aInt As Byte
    Pend As Boolean
End Type

Private Type tClickPositions
    X(Limite_Medir) As Single
    Y(Limite_Medir) As Single
    aInt As Byte
End Type



Private GameClick As tClicks '(1 To 3) As tClicks
Private ClickPositions As tClickPositions '(1 To 4) As tClickPositions

Private ModoIntervalo As Boolean


Private miIndex As Byte
Public LIMITE_INTERVALO As Byte  'Es el limite de diferencia entre intervalos para guardar en logs en milisegundos

Public Sub Inicializar(ByVal modIntervalo As Boolean, ByVal Index As Byte, Optional ByVal LimiteInt As Byte = 10)
    'LimiteINt seria el limite de comparacion de intervalos en milisegundos para despertar una "advertencia"
    ModoIntervalo = modIntervalo
    LIMITE_INTERVALO = LimiteInt
    miIndex = Index
    
End Sub

Public Function puedeVerPos() As Boolean
    If GameClick.Pend Then
        
    End If
End Function
' Se almazenan los intervalos entre el click en la lista de hechizos y el boton lanzar
' y luego de 5 intervalos, se comparan

Public Sub PrimerEvento() 'cambio al menu inventario
    GameClick.LCLista = timeGetTime
    GameClick.Pend = True
End Sub

Public Sub SegundoEvento() 'Click en algun objeto del inventario
    With GameClick
        If .Pend = True Then
            
            .intClick(.aInt) = timeGetTime - .LCLista
            .Pend = False
            If .aInt = (Limite_Medir) Then
                .aInt = 0
                Call CompararInt '(2)
            End If
            .aInt = .aInt + 1
        End If
    End With
End Sub


Public Sub ClickEnObjetoPos(ByVal X As Single, ByVal Y As Single)
    With ClickPositions
        .aInt = .aInt + 1
        .X(.aInt) = X
        .Y(.aInt) = Y
        
        If .aInt = Limite_Medir Then
            .aInt = 0
            Call CompararInt '(3 + TIPO)
        End If
    End With
End Sub



Private Sub CompararInt()
    Dim advs As Byte
    If ModoIntervalo = True Then
        With GameClick
        
            advs = CompararIntervalos(.intClick)
            
            If advs >= 8 Then _
                Call WriteEnviarAviso(miIndex) 'SendData("SACSAC1" & miIndex)
            
        End With
    Else
        With ClickPositions
            
            advs = CompararPosiciones(.X)
                    
            advs = advs + CompararPosiciones(.Y)

            If advs >= 16 Then _
                Call WriteEnviarAviso(miIndex) ' SendData("SACSAC1" & miIndex)
                
        End With
    End If
End Sub
Private Function CompararPosiciones(ByRef ints() As Single) As Byte 'Devuelve la cantidad de posiciones iguales de click encontrados
    Dim xx As Long, yy As Long
    Dim dif() As Single
    Dim advs As Byte
    Dim nDif As Byte
    Dim Advertencias As Byte
    ReDim Preserve dif(0 To 1) As Single
    For xx = 0 To Limite_Medir
        For yy = 0 To Limite_Medir
            If yy <> xx Then
                dif(nDif) = ints(xx) - ints(yy)
                If dif(nDif) < 1 And dif(nDif) > -1 Then
                    Advertencias = Advertencias + 1
                End If
                nDif = nDif + 1
                ReDim Preserve dif(0 To nDif) As Single
            End If
        Next yy
    Next xx
    CompararPosiciones = Advertencias
End Function

Private Function CompararIntervalos(ByRef ints() As Long) As Byte 'Devuelve la cantidad de intervalos parecidos encontrados
    Dim xx As Long, yy As Long
    Dim dif() As Long
    Dim advs As Byte
    Dim nDif As Byte
    Dim Advertencias As Byte
    
    
    ReDim Preserve dif(0 To 1) As Long
    
    For xx = 0 To Limite_Medir
        For yy = 0 To Limite_Medir
        
            If yy <> xx Then
                dif(nDif) = Abs(ints(xx) - ints(yy))
                If dif(nDif) <= LIMITE_INTERVALO Then
                    Advertencias = Advertencias + 1
                End If
                nDif = nDif + 1
                ReDim Preserve dif(0 To nDif) As Long
            End If
            
        Next yy
    Next xx
    CompararIntervalos = Advertencias
End Function



