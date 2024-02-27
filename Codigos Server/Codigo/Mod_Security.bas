Attribute VB_Name = "ModEngine"
Option Explicit

Public ActualKey    As Integer

Public Const MSG_Expuls     As String = "POR MEDIDAS DE SEGURIDAD, SE TE HA DESCONECTADO."

Public Function MAP_ENC(ByVal map As Integer) As Long


MAP_ENC = RandomNumber(1000, 9999) & (map * 5)

End Function

Public Function MAP_DEC(ByVal lMap As Long) As Integer

Dim map    As String
Dim sRes   As Integer

map = CStr(lMap)

sRes = Val(mid$(map, 5))

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

cL = Val(mid$(sa, 3))

USE_DEC = (cL / 2)

End Function


