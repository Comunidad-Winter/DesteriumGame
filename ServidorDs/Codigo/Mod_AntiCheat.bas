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
 
Public Sub RestoTiempo(ByVal Userindex As Integer)

          '// Miqueas150
          '// Vamos restando tiempo a os intervalos para poder ejecutarlos :v
10        With UserList(Userindex).Counters

20            If .Seguimiento.Golpe > 0 Then '// Restamos al intervalo "Golpe" para poder pegar

30                .Seguimiento.Golpe = .Seguimiento.Casteo - 1

40            End If

50            If .Seguimiento.Casteo > 0 Then '// Restamos al intervalo "Casteo" para poder pegar

60                .Seguimiento.Casteo = .Seguimiento.Casteo - 1

70            End If
80        End With
End Sub
 
Public Sub SetIntervalos(ByVal Userindex As Integer)

          '// Miqueas150
          '// Seteamos las Variables a 0
10        With UserList(Userindex).Counters

20            .Seguimiento.Casteo = 0
30            .Seguimiento.Golpe = 0

40        End With

          '// Wa
          '// We
          '// Wi
          '// Wo
          '// Wu
          '// No quiero intervalos con , puta GoDKeR
50        IntervaloCasteo = Int(IntervaloLanzaHechizo / 40)
          '// No quiero intervalos con , puta GoDKeR
60        IntervaloPego = Int(IntervaloUserPuedeAtacar / 40)
          '// Si preguntan el porque /40 ? Es porque el timer principal de AO usa 40 ms _
           y bueno loco son las reglas no jodan ...

End Sub
 
Public Function PuedoCasteoHechizo(ByVal Userindex As Integer) As Boolean

          '// Miqueas
          '// Controlamos que pueda Tirar Hechizos
10        With UserList(Userindex).Counters

20            If .Seguimiento.Casteo > 0 Then

30                PuedoCasteoHechizo = False
40                Exit Function

50            End If

60            PuedoCasteoHechizo = True
              '// ....
70            .Seguimiento.Casteo = IntervaloCasteo

80        End With
End Function
 
Public Function PuedoPegar(ByVal Userindex As Integer) As Boolean

          '// Miqueas
          '// Controlamos que pueda Pegar
10        With UserList(Userindex).Counters

20            If .Seguimiento.Golpe > 0 Then

30                PuedoPegar = False
40                Exit Function

50            End If

60            PuedoPegar = True
              '// ....
70            .Seguimiento.Golpe = IntervaloPego

80        End With
End Function
 
Public Function PuedoUsar(ByVal Userindex As Integer, ByVal tipo As Byte) As Boolean

          '// Miqueas
          '// Controlamos que pueda usar cosas e.e (?)
10        With UserList(Userindex).Counters

20            If .Seguimiento.Poteo > 0 Then
30                If Not PuedeChupar(Userindex, tipo) Then Exit Function

40                PuedoUsar = True
50            Else
60                .Seguimiento.Poteo = GetTickCount
70                PuedoUsar = False

80            End If
90        End With
End Function
 
Private Function PuedeChupar(ByVal Userindex As Integer, ByVal tipo As Byte) As Boolean

          '// Miqueas : Funcion Creada por el puto amo MaTih.-
          Dim IntervaloUsar As Integer

10        If (tipo <> 0) Then

20            IntervaloUsar = IntervaloUserPuedeUsar '// Intervalo seteado en server.ini
30        Else
40            IntervaloUsar = IntervaloUserPuedeUsar * 0.5 '// Al intervalo para u + click lo ponemos mas rapido

50        End If

60        With UserList(Userindex).Counters

              '// GoDKeR sos Puto si lee esto
70            If GetTickCount - .Seguimiento.Poteo < IntervaloUsar Then

80                PuedeChupar = False
90            Else
100               PuedeChupar = True
110               .Seguimiento.Poteo = 0

120           End If
130       End With
End Function
 
Private Sub BanAntiCheat(ByVal Userindex As Integer)

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

10        tUser = Userindex

20        With UserList(tUser)

              '// Msj para escracharlo
30            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Sistema de AntiCheat> " & " ha baneado a " & .Name & ": BAN POR " & LCase$(Reason) & ".", FontTypeNames.FONTTYPE_SERVER))
              '// Ponemos el flag de ban a 1
40            .flags.Ban = 1
              '// Ponemos el flag de ban a 1
50            Call WriteVar(CharPath & .Name & ".chr", "FLAGS", "Ban", "1")
              '// Ponemos la pena
60            cantPenas = val(GetVar(CharPath & .Name & ".chr", "PENAS", "Cant"))
              '// Sumamos la pena
70            Call WriteVar(CharPath & .Name & ".chr", "PENAS", "Cant", cantPenas + 1)
              '// Aplicamos por que se lo Baneo
80            Call WriteVar(CharPath & .Name & ".chr", "PENAS", "P" & cantPenas + 1, "By - Anti Cheat" & ": BAN POR " & LCase$(Reason) & " " & Date$ & " " & time$)
90            Call CloseSocket(tUser)

100       End With
End Sub


