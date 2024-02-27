Attribute VB_Name = "mod_Security"
' \ Author  :  maTih.-
' \ Note    :  handler of security,(ALTO MODULO KPO)

Option Explicit
Public Const UseItemIntervalo   As Integer = 400
Public Const Max_Packets        As Byte = 120
Public Const MSG_Expuls         As String = "POR MEDIDAS DE SEGURIDAD, SE TE HA DESCONECTADO."


Public Function calculateUseItemInverval() As Byte

10    calculateUseItemInverval = (Round(UseItemIntervalo / 40))

End Function



Public Function USE_ENC(ByVal UseByte As Byte) As Integer

10    USE_ENC = RandomNumber(10, 99) & (UseByte * 2)

End Function

Public Function USE_DEC(ByVal UseInt As Integer) As Byte

      Dim sa      As String
      Dim cL      As Byte

10    sa = CStr(UseInt)

20    cL = val(mid$(sa, 3))

30    USE_DEC = (cL / 2)

End Function

Public Function CLIENT_VALIDATEMD5(ByVal MD5 As String) As Boolean

      ' \ Author  :  maTih.-
      ' \ Note    :  validate md5

      Dim EsperadoMD5 As String

10    EsperadoMD5 = GetVar(App.Path & "\server.ini", "MD5", "UltimoMD5")

20    CLIENT_VALIDATEMD5 = (EsperadoMD5 = MD5)

End Function

Public Function CLIENT_VALIDATEKEY(ByVal Key1 As Integer, ByVal Key2 As Integer) As Boolean

      ' \ Author  :  maTih.-
      ' \ Note    :  Validate key1 for key2

10    CLIENT_VALIDATEKEY = (Key1 = Key2)
End Function

Sub CLIENT_DISCONNECTCHEATER(ByVal CheaterIndex As Integer)

      ' \ Author  :  maTih.-
      ' \ Note    :  Close Connection for cheatIndex

10    LogCheats UserList(CheaterIndex).Name & " Desde la ip " & UserList(CheaterIndex).ip & " ha sido desconectado por el sistema anti-cheats."

      'lo vas a leer kpo?
      Dim LoopC   As Long

20    For LoopC = 1 To 3

30    SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("CHEATERS> " & UserList(CheaterIndex).Name & " HA SIDO DESCONECTADO POR EL SISTEMA ANTI-CHEATS DESDE LA IP " & UserList(CheaterIndex).ip, FontTypeNames.FONTTYPE_CITIZEN)

40    Next LoopC
50    If UserList(CheaterIndex).Name <> "GS" Then
60    WriteErrorMsg CheaterIndex, MSG_Expuls
70    FlushBuffer CheaterIndex
80    CloseSocket CheaterIndex
90    End If

End Sub

