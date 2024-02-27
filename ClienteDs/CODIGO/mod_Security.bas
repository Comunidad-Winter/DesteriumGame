Attribute VB_Name = "mod_Security"
' \ Author  :  maTih.-
' \ Note    :  handler of security,(ALTO MODULO KPO)

Option Explicit
Public Const UseItemIntervalo   As Integer = 400
Public Const Max_Packets        As Byte = 120
Public Const MSG_Expuls         As String = "POR MEDIDAS DE SEGURIDAD, SE TE HA DESCONECTADO."


Public Function calculateUseItemInverval() As Byte

calculateUseItemInverval = (Round(UseItemIntervalo / 40))

End Function

Public Function MAP_ENC(ByVal Map As Integer) As Long

' \ Author  :  maTih.-
' \ Note    :  ENCRYPT MAP NUMBER

MAP_ENC = RandomNumber(1000, 9999) & (Map * 5)

End Function

Public Function MAP_DEC(ByVal lMap As Long) As Integer

' \ Author  :  maTih.-
' \ Note    :  DECRYPT LONG MAP NUMBER

Dim Map    As String
Dim sRes   As Integer

Map = CStr(lMap)

sRes = val(mid$(Map, 4))

sRes = (sRes / 5)

MAP_DEC = sRes

End Function

Public Function USE_ENC(ByVal UseByte As Byte) As Integer

USE_ENC = RandomNumber(10, 99) & (UseByte * 2)

End Function

Public Function USE_DEC(ByVal UseInt As Integer) As Byte

Dim sa      As String
Dim cL      As Byte

sa = CStr(UseInt)

cL = val(mid$(sa, 3))

USE_DEC = (cL / 2)

End Function

Public Function CLIENT_VALIDATEMD5(ByVal MD5 As String) As Boolean

' \ Author  :  maTih.-
' \ Note    :  validate md5

Dim EsperadoMD5 As String

EsperadoMD5 = GetVar(App.Path & "\server.ini", "MD5", "UltimoMD5")

CLIENT_VALIDATEMD5 = (EsperadoMD5 = MD5)

End Function

Public Function CLIENT_VALIDATEKEY(ByVal Key1 As Integer, ByVal Key2 As Integer) As Boolean

' \ Author  :  maTih.-
' \ Note    :  Validate key1 for key2

CLIENT_VALIDATEKEY = (Key1 = Key2)
End Function

Sub CLIENT_DISCONNECTCHEATER(ByVal CheaterIndex As Integer)

' \ Author  :  maTih.-
' \ Note    :  Close Connection for cheatIndex

LogCheats UserList(CheaterIndex).name & " Desde la ip " & UserList(CheaterIndex).ip & " ha sido desconectado por el sistema anti-cheats."

'lo vas a leer kpo?
Dim LoopC   As Long

For LoopC = 1 To 3

SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("CHEATERS> " & UserList(CheaterIndex).name & " HA SIDO DESCONECTADO POR EL SISTEMA ANTI-CHEATS DESDE LA IP " & UserList(CheaterIndex).ip, FontTypeNames.FONTTYPE_CITIZEN)

Next LoopC
If UserList(CheaterIndex).name <> "GS" Then
WriteErrorMsg CheaterIndex, MSG_Expuls
FlushBuffer CheaterIndex
CloseSocket CheaterIndex
End If

End Sub

