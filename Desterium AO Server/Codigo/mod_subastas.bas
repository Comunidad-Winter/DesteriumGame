Attribute VB_Name = "mod_subastas"

Option Explicit
Private Const SUBASTAGLD_MAX As Long = 50000000
Private Const GLD_MIN As Long = 1000
Private Const MAX_SUBASTAS As Byte = 10
Public LastSubasta As Byte

Public ObjSubasta As tSubastas
Private UltimateSlotLibre As Byte
 
Type tSubastas
NombreObj(1 To MAX_SUBASTAS) As String
Vendedor(1 To MAX_SUBASTAS) As String
precio(1 To MAX_SUBASTAS) As Long
ItemIndex(1 To MAX_SUBASTAS) As Integer
IndexVendedor(1 To MAX_SUBASTAS) As Integer
End Type
 
 
 
Public Sub HandleSubastarObjeto(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal PrecioObj As Long)
If ObjIndex <= 0 Then
Call WriteConsoleMsg(UserIndex, "El objeto es inválido.", FontTypeNames.FONTTYPE_INFO)
ElseIf EsNewbie(UserIndex) Then
Call WriteConsoleMsg(UserIndex, "Eres newbie!.", FontTypeNames.FONTTYPE_INFO)
ElseIf PrecioObj > SUBASTAGLD_MAX Then
Call WriteConsoleMsg(UserIndex, "El valor máximo de oro para subastar es de " & SUBASTAGLD_MAX & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
ElseIf LastSubasta >= MAX_SUBASTAS Then
Call WriteConsoleMsg(UserIndex, "Ya se estan subastando demaciados objetos!.", FontTypeNames.FONTTYPE_INFO)
ElseIf UserList(UserIndex).flags.Subastando > 0 Then
Call WriteConsoleMsg(UserIndex, "¡Ya te encuentras en una subasta!, para hacer otra deberás salir y entrar en el juego.", FontTypeNames.FONTTYPE_INFO)
ElseIf PrecioObj < GLD_MIN Then
Call WriteConsoleMsg(UserIndex, "El precio mínimo es de 1.000 monedas de oro.", FontTypeNames.FONTTYPE_INFO)
Else
Dim LastetSubasta As Byte
If UltimateSlotLibre > 0 Then
LastetSubasta = UltimateSlotLibre
Else
LastSubasta = LastSubasta + 1
LastetSubasta = LastSubasta
End If
ObjSubasta.NombreObj(LastetSubasta) = UCase$(ObjData(ObjIndex).name)
ObjSubasta.precio(LastetSubasta) = PrecioObj
ObjSubasta.Vendedor(LastetSubasta) = UCase$(UserList(UserIndex).name)
ObjSubasta.ItemIndex(LastetSubasta) = ObjIndex
ObjSubasta.IndexVendedor(LastetSubasta) = UserIndex
UserList(UserIndex).flags.Subastando = 1
UserList(UserIndex).flags.Index_Subasta = LastetSubasta
Call QuitarObjetos(ObjIndex, 1, UserIndex)
PrecioObj = 0
ObjIndex = 0
UltimateSlotLibre = 0
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subastas] Un nuevo objeto se agregó al sistema de compra & venta de items. Dirígete al NPC de subastas para obtener más información. Precio del objeto " & ObjSubasta.precio(LastetSubasta) & ".", FontTypeNames.FONTTYPE_GUILD))
End If
End Sub
 
Public Sub UsuarioDesconectaEnSubasta(ByVal UserIndex As Integer)
If UserList(UserIndex).flags.Subastando < 1 Then Exit Sub 'JAO ;-)
Dim i As Integer
Dim MiObj As Obj
For i = 1 To LastSubasta
If UserList(UserIndex).flags.Index_Subasta = i Then
'JAO ; CON ESTO DEPOSITAMOS EL OBJETO SI NO TIENE LUGAR EN EL INV
MiObj.Amount = 1
MiObj.ObjIndex = ObjSubasta.ItemIndex(UserList(UserIndex).flags.Index_Subasta)
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj) 'si no tiene lugar, que se joda :-)
End If
'JAO ; CON ESTO DEPOSITAMOS EL OBJETO SI NO TIENE LUGAR EN EL INV
UltimateSlotLibre = UserList(UserIndex).flags.Index_Subasta
ObjSubasta.NombreObj(UserList(UserIndex).flags.Index_Subasta) = "(VACÍO)"
ObjSubasta.precio(UserList(UserIndex).flags.Index_Subasta) = 0
ObjSubasta.Vendedor(UserList(UserIndex).flags.Index_Subasta) = ""
ObjSubasta.ItemIndex(UserList(UserIndex).flags.Index_Subasta) = 0
ObjSubasta.IndexVendedor(UserList(UserIndex).flags.Index_Subasta) = 0
UserList(UserIndex).flags.Subastando = 0
UserList(UserIndex).flags.Index_Subasta = 0
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subastas] La subasta de " & UserList(UserIndex).name & " se suspende por la desconección del mismo.", FontTypeNames.FONTTYPE_GUILD))
End If
Next i
End Sub
 
Public Sub EnviarSubastaUser(ByVal UserIndex As Integer)
Dim i As Byte
 
'JAO ;-)
If UserList(UserIndex).flags.TargetNPC = 0 Then
Call WriteConsoleMsg(UserIndex, "¡Clickea el npc subastador!.", FontTypeNames.FONTTYPE_INFO)
ElseIf Not Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Subastador Then Call WriteConsoleMsg(UserIndex, "¡Clickea al npc subastador!.", FontTypeNames.FONTTYPE_INFO)
Call WriteConsoleMsg(UserIndex, "¡Clickea al npc subastador!.", FontTypeNames.FONTTYPE_INFO)
Else
'JAO ;-)
 
For i = 1 To MAX_SUBASTAS
If ObjSubasta.NombreObj(i) = "" Then
ObjSubasta.NombreObj(i) = "(VACÍO)"
ObjSubasta.precio(i) = 0
ObjSubasta.Vendedor(i) = "> "
End If
Call WriteSendSubastas(UserIndex, ObjSubasta.NombreObj(i), ObjSubasta.precio(i), ObjSubasta.Vendedor(i))
Next i
End If
End Sub
 
Public Sub ComprarObjetoSubasta(ByVal UserIndex As Integer, ByVal ObjLista As Byte)
If ObjLista <= 0 Then
Call WriteConsoleMsg(UserIndex, "Objeto inválido, selecciona uno.", FontTypeNames.FONTTYPE_INFO)
ElseIf ObjSubasta.ItemIndex(ObjLista) <= 0 Then
Call WriteConsoleMsg(UserIndex, "Objeto inválido, selecciona uno.", FontTypeNames.FONTTYPE_INFO)
'ElseIf UserList(userIndex).flags.Subastando > 0 Or ObjSubasta.Vendedor(UserList(userIndex).flags.Index_Subasta) = UCase$(UserList(userIndex).name) Then
'Call WriteConsoleMsg(userIndex, "No podes comprar un objeto que vos mismo subastaste. Sal y entra del juego para cancelar la subasta.", FontTypeNames.FONTTYPE_INFO)
ElseIf UserList(UserIndex).Stats.Gld < ObjSubasta.precio(ObjLista) Then
Call WriteConsoleMsg(UserIndex, "Para comprar éste objeto necesitas " & ObjSubasta.precio(ObjLista) & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
ElseIf EsNewbie(UserIndex) Then
Call WriteConsoleMsg(UserIndex, "¡Los newbies no pueden acceder a las subastas!", FontTypeNames.FONTTYPE_INFO)
ElseIf ObjSubasta.IndexVendedor(ObjLista) <= 0 Then
Call WriteConsoleMsg(UserIndex, "El vendedor posee un index inválido, reporte el error a los administradores.", FontTypeNames.FONTTYPE_INFO)
Else
Dim MiObj As Obj
MiObj.Amount = 1
MiObj.ObjIndex = ObjSubasta.ItemIndex(ObjLista)
If Not MeterItemEnInventario(UserIndex, MiObj) Then
Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
Call WriteConsoleMsg(ObjSubasta.IndexVendedor(ObjLista), "" & UserList(UserIndex).name & " compró tu objeto ha " & ObjSubasta.precio(ObjLista) & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
UserList(ObjSubasta.IndexVendedor(ObjLista)).Stats.Gld = UserList(ObjSubasta.IndexVendedor(ObjLista)).Stats.Gld + ObjSubasta.precio(ObjLista)
Call WriteUpdateUserStats(ObjSubasta.IndexVendedor(ObjLista))
UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld - ObjSubasta.precio(ObjLista)
Call WriteUpdateUserStats(UserIndex)
UltimateSlotLibre = ObjLista
ObjSubasta.IndexVendedor(ObjLista) = 0
ObjSubasta.ItemIndex(ObjLista) = 0
ObjSubasta.NombreObj(ObjLista) = ""
ObjSubasta.precio(ObjLista) = 0
ObjSubasta.Vendedor(ObjLista) = ""
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subastas] El slot " & val(ObjLista) & " se encuentra disponible.", FontTypeNames.FONTTYPE_GUILD))
ObjLista = 0
End If
End Sub
