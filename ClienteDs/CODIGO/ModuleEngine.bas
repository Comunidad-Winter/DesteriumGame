Attribute VB_Name = "ModuleEngine"
Option Explicit

Public ActualKey    As Integer

Public Const MSG_Expuls     As String = _
    "POR MEDIDAS DE SEGURIDAD, SE TE HA DESCONECTADO."

Public Function MAP_ENC(ByVal map As Integer) As Long

      ' \ Author  :  maTih.-
      ' \ Note    :  ENCRYPT MAP NUMBER

10    MAP_ENC = RandomNumber(1000, 9999) & (map * 5)

End Function

Public Function MAP_DEC(ByVal lMap As Long) As Integer

      ' \ Author  :  maTih.-
      ' \ Note    :  DECRYPT LONG MAP NUMBER

      Dim map    As String
      Dim sRes   As Integer

10    map = CStr(lMap)

20    sRes = Val(mid$(map, 5))

30    sRes = (sRes / 5)

40    MAP_DEC = sRes

End Function

Public Function USE_ENC(ByVal UseByte As Byte) As Integer

10    USE_ENC = RandomNumber(10, 99) & (UseByte * 2)

End Function

Public Function USE_DEC(ByVal UseInt As Integer) As Byte

      Dim sa      As String
      Dim cL      As Byte

10    sa = CStr(UseInt)

20    cL = Val(mid$(sa, 3))

30    USE_DEC = (cL / 2)

End Function



