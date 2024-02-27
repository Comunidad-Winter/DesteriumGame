Attribute VB_Name = "mGranPoder"
' Módulo de gran poder. Idea basada en el viejo código de TPAO pero con el toque de Lautaro
' Si lees esto te violaron de chiquito

Option Explicit

Private Type GreatPower
    LastUser As String
    CurrentUser As String
    CurrentMap As Integer
End Type

Public GreatPower As GreatPower

Public Function UserIndex_GreatPower() As Boolean
          ' • Chequeo el usuario que va a recibir el poder
          ' • Otorgamos el Gran Poder a un usuario Random
          ' • Comentarios por si otro programador toca esto(?)
          
On Error GoTo UserIndex_GreatPower_Error

          Dim LoopC As Integer
          Dim UserIndex As Integer
10        Dim Exist As Boolean: Exist = False
50        UserIndex = RandomNumber(1, LastUser)
               
60        With UserList(UserIndex)
70                If (.flags.UserLogged = True) And (.flags.Muerto = 0) And (.flags.Privilegios = User) And (.Pos.map <> 176 And .Pos.map <> 191) Then
80                    If (StrComp(GreatPower.LastUser, UCase$(.Name)) <> 0) And (StrComp(GreatPower.CurrentUser, UCase$(.Name)) <> 0) And _
                          (MapInfo(.Pos.map).Pk = True) Then
                          
90                        GreatPower.LastUser = UCase$(GreatPower.CurrentUser)
100                       GreatPower.CurrentUser = UCase$(.Name)
110                       GreatPower.CurrentMap = .Pos.map
120                       UserIndex_GreatPower = True
130                       Exist = True
150                   End If
160               End If
170       End With
          
190       If UserIndex_GreatPower Then
200           With UserList(UserIndex)
210               SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                      "Gran Poder de los Dioses» Los dioses le han otorgado el poder al personaje " & GreatPower.CurrentUser & _
                      " ubicado en el mapa " & GreatPower.CurrentMap & "(" & MapInfo(GreatPower.CurrentMap).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
                  
220               RefreshCharStatus UserIndex
230           End With
240       End If


   On Error GoTo 0
   Exit Function

UserIndex_GreatPower_Error:

    ReportError "GRANPODER", "Error " & Err.Number & " (" & Err.Description & ") in procedure UserIndex_GreatPower of Módulo mGranPoder in line " & Erl
End Function

Public Function Check_GreatPower(ByVal UserIndex As Integer, _
                                Optional ByVal AttackerIndex As Integer = 0) As Boolean
10        Check_GreatPower = True
          
On Error GoTo Check_GreatPower_Error

20        With UserList(UserIndex)

              
30            Exit Function
              
              ' ¿Se fue a zona segura?
40            If Not MapInfo(.Pos.map).Pk Then Check_GreatPower = False
              
              ' ¿Deslogea?
50            If Not .flags.UserLogged Then Check_GreatPower = False
              
              ' ¿Muerto?
60            If .flags.Muerto Then Check_GreatPower = False

              ' USUARIO SIGUE CON GRAN PODER PERO CAMBIO DE MAPA
70            If Check_GreatPower Then
80                If .Pos.map <> GreatPower.CurrentMap Then
                      
90                    GreatPower.CurrentMap = .Pos.map
                      
100                   If RandomNumber(1, 10) <= 4 Then
110                       SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                          "Gran Poder de los Dioses» " & GreatPower.CurrentUser & _
                          " ubicado en el mapa " & GreatPower.CurrentMap & "(" & MapInfo(GreatPower.CurrentMap).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
120                   End If
130               End If
140           End If
              
              ' Se busca nuevo usuario si se pierde por causa "natural"
150           If (Check_GreatPower = False) And (AttackerIndex = 0) Then
160               GreatPower.LastUser = UCase$(.Name)
170               GreatPower.CurrentUser = vbNullString
180               GreatPower.CurrentMap = 0

                  RefreshCharStatus UserIndex
190               UserIndex_GreatPower
                  
200
              ' Muere por un user mas polenta
210           ElseIf (Check_GreatPower = False) And (AttackerIndex > 0) Then
220               GreatPower.LastUser = UCase$(.Name)
230               GreatPower.CurrentUser = UCase$(UserList(AttackerIndex).Name)
240               GreatPower.CurrentMap = UserList(AttackerIndex).Pos.map
250               RefreshCharStatus UserIndex
260               RefreshCharStatus AttackerIndex
                  
270               SendData SendTarget.ToAll, 0, PrepareMessageConsoleMsg( _
                      "Gran Poder de los Dioses» El poder ha pasado a manos de " & UserList(AttackerIndex).Name & _
                      " ubicado en el mapa " & UserList(AttackerIndex).Pos.map & "(" & MapInfo(UserList(AttackerIndex).Pos.map).Name & ")", FontTypeNames.FONTTYPE_PREMIUM)
280           End If
290       End With

   On Error GoTo 0
   Exit Function

Check_GreatPower_Error:

    ReportError "GRANPODER", "Error " & Err.Number & " (" & Err.Description & ") in procedure Check_GreatPower of Módulo mGranPoder in line " & Erl
End Function
