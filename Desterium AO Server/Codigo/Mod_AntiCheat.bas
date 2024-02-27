Attribute VB_Name = "Mod_AntiCheat"
'***************************************************
'// Autor:  Miqueas
'// Creado : 22/02/2010
'// Sistema de seguridad Basico, contra pete sirve. _
 Al ser controlado desde el servidor los que "editan memoria" _
 No pueden hacer nada para poder sacar ventaja de variables por parte del cliente
'***************************************************
Option Explicit

Public Type Intervalos

    Poteo As Long
    Golpe As Integer
    Casteo As Integer

End Type

Private Declare Function GetTickCount Lib "kernel32" () As Long

'// Las declaramos aca para evitar una nueva declaracion cadaves que se llame al sub
Private IntervaloCasteo As Integer
Private IntervaloPego   As Integer
 
Public Sub RestoTiempo(ByVal UserIndex As Integer)

    '// Miqueas150
    '// Vamos restando tiempo a os intervalos para poder ejecutarlos :v
    With UserList(UserIndex).Counters

        If .Seguimiento.Golpe > 0 Then '// Restamos al intervalo "Golpe" para poder pegar

            .Seguimiento.Golpe = .Seguimiento.Casteo - 1

        End If

        If .Seguimiento.Casteo > 0 Then '// Restamos al intervalo "Casteo" para poder pegar

            .Seguimiento.Casteo = .Seguimiento.Casteo - 1

        End If
    End With
End Sub
 
Public Sub SetIntervalos(ByVal UserIndex As Integer)

    '// Miqueas150
    '// Seteamos las Variables a 0
    With UserList(UserIndex).Counters

        .Seguimiento.Casteo = 0
        .Seguimiento.Golpe = 0

    End With

    '// Wa
    '// We
    '// Wi
    '// Wo
    '// Wu
    '// No quiero intervalos con , puta GoDKeR
    IntervaloCasteo = Int(IntervaloLanzaHechizo / 40)
    '// No quiero intervalos con , puta GoDKeR
    IntervaloPego = Int(IntervaloUserPuedeAtacar / 40)
    '// Si preguntan el porque /40 ? Es porque el timer principal de AO usa 40 ms _
     y bueno loco son las reglas no jodan ...

End Sub
 
Public Function PuedoCasteoHechizo(ByVal UserIndex As Integer) As Boolean

    '// Miqueas
    '// Controlamos que pueda Tirar Hechizos
    With UserList(UserIndex).Counters

        If .Seguimiento.Casteo > 0 Then

            PuedoCasteoHechizo = False
            Exit Function

        End If

        PuedoCasteoHechizo = True
        '// ....
        .Seguimiento.Casteo = IntervaloCasteo

    End With
End Function
 
Public Function PuedoPegar(ByVal UserIndex As Integer) As Boolean

    '// Miqueas
    '// Controlamos que pueda Pegar
    With UserList(UserIndex).Counters

        If .Seguimiento.Golpe > 0 Then

            PuedoPegar = False
            Exit Function

        End If

        PuedoPegar = True
        '// ....
        .Seguimiento.Golpe = IntervaloPego

    End With
End Function
 
Public Function PuedoUsar(ByVal UserIndex As Integer, ByVal tipo As Byte) As Boolean

    '// Miqueas
    '// Controlamos que pueda usar cosas e.e (?)
    With UserList(UserIndex).Counters

        If .Seguimiento.Poteo > 0 Then
            If Not PuedeChupar(UserIndex, tipo) Then Exit Function

            PuedoUsar = True
        Else
            .Seguimiento.Poteo = GetTickCount
            PuedoUsar = False

        End If
    End With
End Function
 
Private Function PuedeChupar(ByVal UserIndex As Integer, ByVal tipo As Byte) As Boolean

    '// Miqueas : Funcion Creada por el puto amo MaTih.-
    Dim IntervaloUsar As Integer

    If (tipo <> 0) Then

        IntervaloUsar = IntervaloUserPuedeUsar '// Intervalo seteado en server.ini
    Else
        IntervaloUsar = IntervaloUserPuedeUsar * 0.5 '// Al intervalo para u + click lo ponemos mas rapido

    End If

    With UserList(UserIndex).Counters

        '// GoDKeR sos Puto si lee esto
        If GetTickCount - .Seguimiento.Poteo < IntervaloUsar Then

            PuedeChupar = False
        Else
            PuedeChupar = True
            .Seguimiento.Poteo = 0

        End If
    End With
End Function
 
Private Sub BanAntiCheat(ByVal UserIndex As Integer)

    '***************************************************
    '// Autor: Miqueas
    '// 23/11/13
    '// No implementado
    '// ¿Hace falta una explicacion de lo que hace ?
    '// Bueno si, Banea al usuario, Bane codigo original funcion de baneo x ip
    '***************************************************
    Dim tUser     As Integer
    Dim cantPenas As Byte

    Const Reason  As String = "Uso de programas externos"

    tUser = UserIndex

    With UserList(tUser)

        '// Msj para escracharlo
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Sistema de AntiCheat> " & " ha baneado a " & .Name & ": BAN POR " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
        '// Ponemos el flag de ban a 1
        .flags.Ban = 1
        '// Ponemos el flag de ban a 1
        Call WriteVar(CharPath & .Name & ".chr", "FLAGS", "Ban", "1")
        '// Ponemos la pena
        cantPenas = val(GetVar(CharPath & .Name & ".chr", "PENAS", "Cant"))
        '// Sumamos la pena
        Call WriteVar(CharPath & .Name & ".chr", "PENAS", "Cant", cantPenas + 1)
        '// Aplicamos por que se lo Baneo
        Call WriteVar(CharPath & .Name & ".chr", "PENAS", "P" & cantPenas + 1, "By - Anti Cheat" & ": BAN POR " & LCase$(Reason) & " " & Date$ & " " & time$)
        Call CloseSocket(tUser)

    End With
End Sub


