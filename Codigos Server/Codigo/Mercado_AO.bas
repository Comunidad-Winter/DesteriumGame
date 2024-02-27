Attribute VB_Name = "Mercado_AO"
Option Explicit
 
Public Type tPj
  Nombre As String
  Oro As Long
  MinimeLvl As Byte
  NamePjRecibidor As String
End Type
 
Public Type PjComer
  Pjs() As tPj
  LastPj As Byte
End Type

Public ComercioPJ As PjComer
'~147~250~69~1~1~ COLOR VERDE
 
Sub CPJ_LoadVentas()
 
Dim salamechar As String
Dim i As Long
salamechar = App.Path & "\Ventas.ini"
 
With ComercioPJ
 
.LastPj = val(GetVar(salamechar, "CANTIDAD", "INIT"))
 
For i = 1 To .LastPj
 
.Pjs(i).MinimeLvl = val(GetVar(salamechar, "PJ" & i, "NivelMinimo"))
.Pjs(i).Nombre = GetVar(salamechar, "PJ" & i, "Nombre")
.Pjs(i).NamePjRecibidor = GetVar(salamechar, "PJ" & i, "PjRecibe")
.Pjs(i).Oro = val(GetVar(salamechar, "PJ" & i, "OroPaComprar"))
 
Next i
 
End With
 
End Sub
Sub CPJ_GrabarVentas(ByVal iUserIndex As Integer)
 
Dim salamechar As String
Dim i As Long
salamechar = App.Path & "\Ventas.ini"
 
With ComercioPJ
 
.LastPj = val(GetVar(salamechar, "CANTIDAD", "INIT"))
 
For i = 1 To 2
 
WriteVar salamechar, "PJ" & i, "NivelMinimo", ComercioPJ.Pjs(iUserIndex).MinimeLvl
WriteVar salamechar, "PJ" & i, "Nombre", ComercioPJ.Pjs(iUserIndex).Nombre
WriteVar salamechar, "PJ" & i, "PjRecibe", ComercioPJ.Pjs(iUserIndex).NamePjRecibidor
WriteVar salamechar, "PJ" & i, "OroPaComprar", ComercioPJ.Pjs(iUserIndex).Oro
 
Next i
 
End With
 
End Sub
 
Public Sub CPJ_AddPersonaje(ByVal iUserIndex As Integer, ByVal Minimo As Long, ByVal Oro As Long, ByVal PJR As String)
UserList(iUserIndex).flags.EstaEnMercado = False
If UserList(iUserIndex).flags.EstaEnMercado = True Then

WriteConsoleMsg iUserIndex, "Mercado> Este personaje ya está a la venta! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_GUILD
 Else
ComercioPJ.LastPj = ComercioPJ.LastPj + 1
 
ReDim Preserve ComercioPJ.Pjs(1 To ComercioPJ.LastPj) As tPj
 
ComercioPJ.Pjs(ComercioPJ.LastPj).Nombre = UserList(iUserIndex).Name
ComercioPJ.Pjs(ComercioPJ.LastPj).MinimeLvl = Minimo
ComercioPJ.Pjs(ComercioPJ.LastPj).Oro = Oro
ComercioPJ.Pjs(ComercioPJ.LastPj).NamePjRecibidor = PJR

WriteVar CharPath & UserList(iUserIndex).Name & ".chr", "VENTA", "iVenta", ComercioPJ.LastPj
 
WriteVar CharPath & UserList(iUserIndex).Name & ".chr", "VENTA", "EnVenta", "1"
WriteUpdateUserStats iUserIndex
WriteConsoleMsg iUserIndex, "¡Has añadido a la venta a " & UserList(iUserIndex).Name & " satisfactoriamente!", FontTypeNames.fonttype_dios
UserList(iUserIndex).flags.EstaEnMercado = True

WriteVar App.Path & "\Ventas.ini", "CANTIDAD", "INIT", ComercioPJ.LastPj
WriteVar App.Path & "\Ventas.ini", "PJ" & ComercioPJ.LastPj, "NivelMinimo", ComercioPJ.Pjs(ComercioPJ.LastPj).MinimeLvl
WriteVar App.Path & "\Ventas.ini", "PJ" & ComercioPJ.LastPj, "Nombre", ComercioPJ.Pjs(ComercioPJ.LastPj).Nombre
WriteVar App.Path & "\Ventas.ini", "PJ" & ComercioPJ.LastPj, "PjRecibe", ComercioPJ.Pjs(ComercioPJ.LastPj).NamePjRecibidor
WriteVar App.Path & "\Ventas.ini", "PJ" & ComercioPJ.LastPj, "OroPaComprar", ComercioPJ.Pjs(ComercioPJ.LastPj).Oro
CPJ_LoadVentas
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Un personaje ha sido añadido al mercado! Escribe /MERCADO para saber cúal fue.", FontTypeNames.fonttype_dios))
End If
End Sub
 
Public Sub CPJ_ComprarPersonaje(ByVal Comprador As Integer, ByVal NamePJ As String)
 
Dim EnVenta As Byte
Dim IndexVenta As Byte
Dim MiPass As String
Dim CharFilePath As String
CharFilePath = CharPath & NamePJ & ".chr"
 
MiPass = GetVar(CharPath & UserList(Comprador).Name & ".chr", "INIT", "Password")
 
EnVenta = CByte(val(GetVar(CharFilePath, "VENTA", "EnVenta")))
 
If EnVenta = 0 Then Exit Sub
 
IndexVenta = val(GetVar(CharFilePath, "VENTA", "iVenta"))
 
 
With UserList(Comprador)
 
If ComercioPJ.Pjs(IndexVenta).Oro > 0 Then
 
If .Stats.Gld < ComercioPJ.Pjs(IndexVenta).Oro Then
WriteConsoleMsg Comprador, "No tienes suficientes monedas de oro para comprar a " & NamePJ & " (" & ComercioPJ.Pjs(IndexVenta).Oro & ")", FontTypeNames.fonttype_dios
Exit Sub
End If
 
.Stats.Gld = .Stats.Gld - ComercioPJ.Pjs(IndexVenta).Oro
 
WriteUpdateGold Comprador
 
WriteVar CharPath & .Name & ".chr", "STATS", "GLD", .Stats.Gld
 
WriteVar CharFilePath, "INIT", "Password", MiPass
 
WriteVar CharPath & ComercioPJ.Pjs(IndexVenta).NamePjRecibidor & ".chr", "STATS", "GLD", UserList(NameIndex(NamePJ)).Stats.Gld + CStr(ComercioPJ.Pjs(IndexVenta).Oro)
 
WriteVar CharFilePath, "VENTA", "iVenta", "0"
WriteVar CharFilePath, "VENTA", "EnVenta", "0"
 
If NameIndex(NamePJ) > 0 Then
WriteErrorMsg NameIndex(NamePJ), .Name & " a comprado este personaje, el oro que se requeria para comprar el mismo ha sido depositado en " & ComercioPJ.Pjs(IndexVenta).NamePjRecibidor
CloseSocket NameIndex(NamePJ)
End If
 
WriteConsoleMsg Comprador, "Has comprado a " & NamePJ & " su contraseña es igual a la de " & .Name, FontTypeNames.FONTTYPE_GUILD
 
ComercioPJ.Pjs(IndexVenta).Nombre = "Vendido"
ComercioPJ.Pjs(IndexVenta).MinimeLvl = 0
ComercioPJ.Pjs(IndexVenta).NamePjRecibidor = "Vendido"
ComercioPJ.Pjs(IndexVenta).Oro = 0
 UserList(NameIndex(NamePJ)).flags.EstaEnMercado = False
End If
 
 
End With
 
 
End Sub
 
Public Sub CPJ_EnviarSolicitud(ByVal Solicitador As Integer, ByVal OtherPj As String)
 
With UserList(Solicitador)
 
Dim EnVenta As Byte
Dim IndexVenta As Byte
Dim Ofertadores As Byte
Dim CharFilePath As String
CharFilePath = CharPath & OtherPj & ".chr"
 
EnVenta = CByte(val(GetVar(CharFilePath, "VENTA", "EnVenta")))
 
If EnVenta = 0 Then
WriteConsoleMsg Solicitador, "Mercado> El personaje al cual deseas enviar una solicitud de cambio de personaje no se encuentra a la venta! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_GUILD
Else
 
IndexVenta = val(GetVar(CharFilePath, "VENTA", "iVenta"))
Ofertadores = val(GetVar(CharFilePath, "VENTA", "LastOfertador"))
 
If .Stats.ELV < ComercioPJ.Pjs(IndexVenta).MinimeLvl Then
WriteConsoleMsg Solicitador, "El dueño de " & OtherPj & " quiere personajes + " & ComercioPJ.Pjs(IndexVenta).MinimeLvl, FontTypeNames.fonttype_dios
Exit Sub
End If
 WriteVar CharFilePath, "VENTA", "LastOfertador", Ofertadores + 1
 
 
WriteVar CharFilePath, "VENTA", "Ofertador" & Ofertadores + 1, UCase$(.Name)
 
If NameIndex(OtherPj) > 0 Then
  WriteConsoleMsg NameIndex(OtherPj), .Name & " Ha enviado una solicitud de cambio de personaje, para más información escribe /INFOPJ.", FontTypeNames.FONTTYPE_CITIZEN
  WriteConsoleMsg Solicitador, "Has enviado una solicitud de cambio de personaje a " & OtherPj & ".", FontTypeNames.fonttype_dios
  ElseIf NameIndex(OtherPj) <= 0 Then
  WriteConsoleMsg Solicitador, "Mercado> El personaje al que deseas enviar la solicitud de cambio no se encuentra online! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_GUILD
End If
 End If
End With
 
End Sub
 
Public Sub CPJ_CancelarSolicitud(ByVal Cancelador As Integer, ByVal NameP As String)
 
Dim OfertadorI As Byte
Dim CharP As String
Dim i As Long
 
CharP = CharPath & NameP & ".chr"
 
OfertadorI = val(GetVar(CharP, "VENTA", "LastOfertador"))
 UserList(Cancelador).flags.UltimoMensaje = 0
For i = 1 To OfertadorI
 
If GetVar(CharP, "VENTA", "Ofertador" & i) = UCase$(UserList(Cancelador).Name) Then
 
WriteVar CharP, "VENTA", "Ofertador" & i, "Cancelado"
 
WriteConsoleMsg Cancelador, "Mercado> Se ha cancelado la solicitud de cambio. ~250~147~250~1~1~", FontTypeNames.FONTTYPE_CITIZEN
 
Else
 If UserList(Cancelador).flags.UltimoMensaje = 0 Then
WriteConsoleMsg Cancelador, "Mercado> No le has ofrecido cambiar tu personaje. ~250~147~250~1~1~", FontTypeNames.FONTTYPE_CITIZEN
 End If
 UserList(Cancelador).flags.UltimoMensaje = 1
End If
 
Next i
 
End Sub
 
Public Sub CPJ_DenegarSolicitud(ByVal Cancelador As Integer, ByVal Cancelado As String)
 
Dim i As Long
Dim ok As String
Dim cPath As String
Dim Ofertadores As Byte
 Dim SAPE As String
cPath = CharPath & UserList(Cancelador).Name & ".chr"
Ofertadores = val(GetVar(cPath, "VENTA", "LastOfertador"))

If Ofertadores > 0 Then
 
For i = 1 To Ofertadores
 SAPE = GetVar(cPath, "VENTA", "Ofertador" & i)
If NameIndex(SAPE) = Cancelador Then
 
WriteVar cPath, "VENTA", "Ofertador" & i, "Cancelado"
 
ok = "Mercado> Has Cancelado a userlist(nameindex(cancelado)).name. ~250~147~250~1~1~"
 
Else
 
ok = "Mercado> " & UserList(NameIndex(Cancelado)).Name & " no te ha ofrecido cambiar. ~250~147~250~1~1~"
 
End If
 
Next i
 
WriteConsoleMsg Cancelador, ok, FontTypeNames.FONTTYPE_CITIZEN
 
 
End If
 
End Sub
 
Public Sub CPJ_QuitarPersonaje(ByVal Quitador As Integer)
 
Dim CharFile As String
 
With UserList(Quitador)
If .flags.EstaEnMercado = False Then
WriteConsoleMsg Quitador, "Mercado> Este personaje que deseas quitar no está a la venta! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_GUILD
Else
ReDim Preserve ComercioPJ.Pjs(1 To ComercioPJ.LastPj) As tPj
 
 
CharFile = CharPath & .Name & ".chr"
 
ComercioPJ.Pjs(val(GetVar(CharFile, "VENTA", "iVenta"))).Nombre = "Cancelado"
ComercioPJ.Pjs(val(GetVar(CharFile, "VENTA", "iVenta"))).NamePjRecibidor = "Cancelado"
ComercioPJ.Pjs(val(GetVar(CharFile, "VENTA", "iVenta"))).MinimeLvl = 0
ComercioPJ.Pjs(val(GetVar(CharFile, "VENTA", "iVenta"))).Oro = 0
 
WriteVar CharFile, "VENTA", "EnVenta", "0"
WriteVar CharFile, "VENTA", "iVenta", "0"

WriteConsoleMsg Quitador, "Mercado> Has quitado de la venta a " & .Name & ".", FontTypeNames.FONTTYPE_INFO


.flags.EstaEnMercado = False

Dim i As Long
Dim pt As String
pt = CByte(val(GetVar(CharFile, "VENTA", "LastOfertador")))
For i = 1 To pt
 
WriteVar CharFile, "VENTA", "Ofertador" & i, "Cancelado"
 
WriteVar CharFile, "VENTA", "LastOfertador", "0"
 
Next i
 End If
End With
 
End Sub

Public Sub CPJ_CambioPJ(ByVal Comprador As Integer, ByVal NamePJ As String)
 
Dim EnVenta As Byte
Dim IndexVenta As Byte
Dim MiPass As String
Dim SuPass As String
Dim MiPin As String
Dim SuPin As String
Dim MiEmail As String
Dim SuEmail As String
Dim Ofertadores As String
Dim CharFilePath As String
Dim pt As String 'usado para el puto compradoR
CharFilePath = CharPath & NamePJ & ".chr"
pt = CharPath & UserList(Comprador).Name & ".chr"
 
MiPass = GetVar(CharPath & UserList(Comprador).Name & ".chr", "INIT", "Password")
SuPass = GetVar(CharPath & NamePJ & ".chr", "INIT", "Password")
MiPin = GetVar(CharPath & UserList(Comprador).Name & ".chr", "INIT", "PIN")
SuPin = GetVar(CharPath & NamePJ & ".chr", "INIT", "PIN")
MiEmail = GetVar(CharPath & UserList(Comprador).Name & ".chr", "CONTACTO", "Email")
SuEmail = GetVar(CharPath & NamePJ & ".chr", "CONTACTO", "Email")
 
EnVenta = CByte(val(GetVar(CharFilePath, "VENTA", "EnVenta")))
 Dim PETON As String
 PETON = val(GetVar(CharPath & NamePJ & ".chr", "VENTA", "EnVenta"))
'If (PETON = "1") Then
'WriteConsoleMsg Comprador, "Mercado> El personaje con el que deseas cambiar se encuentra en el mercado! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_GUILD
'Exit Sub
'End If
Dim QuiereComp As String
QuiereComp = val(GetVar(pt, "VENTA", "EnVenta"))
If QuiereComp = "0" Then
WriteConsoleMsg Comprador, "Mercado> Solo el personaje que está en mercado puede usar este comando!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If
' UserList(NameIndex(NamePJ)).flags.EstaEnMercado = False
IndexVenta = val(GetVar(CharFilePath, "VENTA", "iVenta"))
 Ofertadores = val(GetVar(CharFilePath, "VENTA", "LastOfertador"))
 
With UserList(Comprador)
 
If ComercioPJ.Pjs(IndexVenta).Oro > 0 Then
 Dim i As Long
 If Ofertadores > 0 Then
 For i = 1 To Ofertadores
 Dim SAPO As String
 SAPO = GetVar(CharFilePath, "VENTA", "Ofertador" & i)
If NameIndex(SAPO) = Ofertadores Then
WriteConsoleMsg Comprador, "Mercado> No solicitaste el cambio de personaje con ese usuario! ~147~250~69~1~1~", FontTypeNames.FONTTYPE_INFO
Else
 UserList(NameIndex(NamePJ)).flags.EstaEnMercado = False
.flags.EstaEnMercado = False
WriteVar CharFilePath, "INIT", "Password", MiPass
WriteVar pt, "INIT", "Password", SuPass
WriteVar CharFilePath, "INIT", "PIN", MiPin
WriteVar pt, "INIT", "PIN", SuPin
WriteVar CharFilePath, "CONTACTO", "Email", MiEmail
WriteVar pt, "CONTACTO", "Email", SuEmail
 
WriteVar CharFilePath, "VENTA", "iVenta", "0"
WriteVar CharFilePath, "VENTA", "EnVenta", "0"

 If NameIndex(NamePJ) > 0 And Comprador > 0 Then
 WriteErrorMsg Comprador, "¡Transferencia exitosa! La contraseña/pin/email de " & UserList(NameIndex(NamePJ)).Name & " es igual a la de " & UserList(Comprador).Name & ""
WriteErrorMsg NameIndex(NamePJ), "¡Transferencia exitosa! La contraseña/pin/email de " & UserList(Comprador).Name & " es igual a la de " & UserList(NameIndex(NamePJ)).Name & ""
FlushBuffer NameIndex(NamePJ)
FlushBuffer Comprador
CloseSocket NameIndex(NamePJ)
CloseSocket Comprador
Else
If NameIndex(NamePJ) > 0 Then
WriteErrorMsg NameIndex(NamePJ), "¡Transferencia exitosa! La contraseña/pin/email es igual a la de " & UserList(NameIndex(NamePJ)).Name & ""
FlushBuffer NameIndex(NamePJ)
CloseSocket NameIndex(NamePJ)
End If
If Comprador > 0 Then
WriteErrorMsg Comprador, "¡Transferencia exitosa! La contraseña/pin/email de " & UserList(NameIndex(NamePJ)).Name & " es igual a la de " & UserList(Comprador).Name & ""
FlushBuffer Comprador
CloseSocket Comprador
End If
End If

ComercioPJ.Pjs(IndexVenta).Nombre = "Vendido"
ComercioPJ.Pjs(IndexVenta).MinimeLvl = 0
ComercioPJ.Pjs(IndexVenta).NamePjRecibidor = "Vendido"
ComercioPJ.Pjs(IndexVenta).Oro = 0

 
End If
 Next i
 End If
 End If
End With
 
 
End Sub


