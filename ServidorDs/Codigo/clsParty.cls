VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsParty.cls
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

Private p_members(1 To PARTY_MAXMEMBERS) As tPartyMember
'miembros

Private p_expTotal As Long
'Estadistica :D

Private p_Fundador As Integer
'el creador

Private p_CantMiembros As Integer
'cantidad de miembros

Private p_SumaNivelesElevados As Single
'suma de todos los niveles elevados a la ExponenteNivelParty > Esta variable se usa para calcular la experiencia repartida en la Party.

'datos en los pjs: | indexParty(indice en p_members), partyLeader(userindex del lider) |

'Constructor de clase
Public Sub Class_Initialize()
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/08
      'Last Modification By: Marco Vanotti (MarKoxX)
      ' - 09/29/07 p_SumaNiveles added (Tavo)
      ' - 07/04/08 p_SumaNiveles changed to p_SumaNivelesElevados (MarKoxX)
      '***************************************************
10        p_expTotal = 0
20        p_CantMiembros = 0
30        p_SumaNivelesElevados = 0
          
          ' // maTih.- //
40        Call initPorcentajes
          
End Sub

'Destructor de clase
Public Sub Class_Terminate()
End Sub

''
' Sets the new p_sumaniveleselevados to the party.
'
' @param lvl Specifies reference to user level
' @remarks When a user level up and he is in a party, we update p_sumaNivelesElavados so the formula still works.
Public Sub UpdateSumaNivelesElevados(ByVal Lvl As Integer)
      '*************************************************
      'Author: Marco Vanotti (MarKoxX)
      'Last modified: 11/24/09
      '11/24/09: Pato - Change the exponent to a variable with the exponent
      '*************************************************
10        p_SumaNivelesElevados = p_SumaNivelesElevados - ((Lvl - 1) ^ ExponenteNivelParty) + Lvl ^ ExponenteNivelParty
End Sub

Public Function MiExperiencia(ByVal Userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: 11/27/09
      'Last Modification By: Budi
      ' - 09/29/07 Experience is round to the biggest number less than that number
      ' - 09/29/07 Now experience is a real-number
      ' - 11/27/09 Arreglé el Out of Range.
      '***************************************************
      'Me dice cuanta experiencia tengo colectada ya en la party
      Dim i As Integer
10    i = 1

20    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
30        i = i + 1
40    Wend

50    If i <= PARTY_MAXMEMBERS Then
60        MiExperiencia = Fix(p_members(i).Experiencia)
70    Else        'esto no deberia pasar :p
80        MiExperiencia = -1
90    End If

End Function

Public Sub ObtenerExito(ByVal ExpGanada As Long, ByVal Mapa As Integer, X As Integer, Y As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/08
      'Last Modification By: Marco Vanotti (MarKoxX)
      ' - 09/29/07 New formula for calculating the experience point of each user
      ' - 09/29/07 Experience is round to the biggest number less than that number
      ' - 09/29/07 Now experience is a real-number
      ' - 04/04/08 Ahora antes de calcular la experiencia a X usuario se fija si ese usuario existe (MarKoxX)
      ' - 07/04/08 New formula to calculate Experience for each user. (MarKoxX)
      '***************************************************
      'Se produjo un evento que da experiencia en la wp referenciada
      Dim i As Integer
      Dim UI As Integer
      Dim expThisUser As Double

10    If p_expTotal >= 2000000000 Then Exit Sub
20    p_expTotal = p_expTotal + ExpGanada


30    For i = 1 To PARTY_MAXMEMBERS
40        UI = p_members(i).Userindex
50        If UI > 0 Then
              ' Formula: Exp* (Nivel ^ ExponenteNivelParty) / sumadeNivelesElevados
              ' expThisUser = CDbl(ExpGanada * (UserList(p_members(i).UserIndex).Stats.ELV ^ ExponenteNivelParty) / p_SumaNivelesElevados)
              
60            expThisUser = Porcentaje(ExpGanada, p_members(i).bPorcentaje)
70            Debug.Print p_expTotal
80            Debug.Print expThisUser
90            If Mapa = UserList(UI).Pos.map And UserList(UI).flags.Muerto = 0 Then
100               If Distance(UserList(UI).Pos.X, UserList(UI).Pos.Y, X, Y) <= PARTY_MAXDISTANCIA Then
110                   p_members(i).Experiencia = p_members(i).Experiencia + expThisUser
120                   If p_members(i).Experiencia < 0 Then
130                       p_members(i).Experiencia = 0
140                   End If
150                   If PARTY_EXPERIENCIAPORGOLPE Then
160                       UserList(UI).Stats.Exp = UserList(UI).Stats.Exp + Fix(expThisUser)
170                       If UserList(UI).Stats.Exp > MAXEXP Then _
                              UserList(UI).Stats.Exp = MAXEXP
180                       Call CheckUserLevel(UI)
190                       Call WriteUpdateUserStats(UI)
200                   End If
210               End If
220           End If
230       End If
240   Next i

End Sub

Public Sub initPorcentajes()

          '
          ' @ maTih.-
          
          Dim pNumUsers   As Integer
          
          Dim j As Long
          
10        For j = 1 To PARTY_MAXMEMBERS
20            If p_members(j).Userindex > 0 Then pNumUsers = (pNumUsers + 1)
30        Next j
          
40        If (pNumUsers > 0) Then
50            For j = 1 To PARTY_MAXMEMBERS
60                If p_members(j).Userindex > 0 Then
70                   p_members(j).bPorcentaje = getPorcentaje(pNumUsers)
                     
80                   Call Protocol.WriteConsoleMsg(p_members(j).Userindex, "Party> Porcentajes actualizados." & vbNewLine & "Porcentaje de experiencia por usuario:" & CStr(p_members(j).bPorcentaje) & ".", FontTypeNames.FONTTYPE_PARTY)
90                End If
100           Next j
110       End If
          
End Sub

Public Function preparePorcentajeString() As String

          '
          ' @ maTih.-
          
          Dim Str As String
          Dim j   As Long
          Dim tmp As Integer
          
10        For j = 1 To PARTY_MAXMEMBERS
20            tmp = p_members(j).Userindex
              
30            If (tmp > 0) Then
40                If (Str = vbNullString) Then
50                    Str = UserList(tmp).Name & "*" & CStr(p_members(j).Experiencia) & "*" & CStr(p_members(j).bPorcentaje)
60                Else
70                    Str = Str & "," & UserList(tmp).Name & "*" & CStr(p_members(j).Experiencia) & "*" & (p_members(j).bPorcentaje)
80                End If
90            End If
100       Next j
          
110       preparePorcentajeString = Str

End Function

Public Function stringToArray(ByRef strCompare As String) As Byte()

          '
          ' @ maTih.- Chupa la pija, aca manda Fran loco
          
          Dim tempArray() As String
          Dim endArray()  As Byte
          Dim tmpPosition As Integer
          Dim nowMember   As String
          Dim MemberIndex As Integer
          Dim TempInt     As Integer
          
          ' initialize ;
10        ReDim endArray(1 To PARTY_MAXMEMBERS) As Byte
          
          ' set array ;
20        tempArray = Split(strCompare, ",")
                        Dim pt As String
                        Dim pt1 As String
                        Dim pt2 As String
                        Dim pt3 As String
          Dim j As Long
          
30        For j = 0 To UBound(tempArray())
40            If Not tempArray(j) = vbNullString Then
50               tmpPosition = InStr(1, tempArray(j), "*")
                 
                 
60               If (tmpPosition > 0) Then
                     ' set member name
70                   nowMember = Left$(tempArray(j), (tmpPosition - 1))
                     
                     ' find memberIndex ;
80                   If Not (nowMember = "Nadie") Then
90                      MemberIndex = findMemberIndex(UCase$(nowMember))
100                  Else
110                     MemberIndex = 0
120                  End If

130                   pt = mid$(tempArray(j), (tmpPosition + 1), 2)
                
                     
140                  If (MemberIndex > 0) Then
150                      TempInt = val(mid$(tempArray(j), tmpPosition + 1, 2))
                      '  TempInt = val(Left$(JIji, pt)) matIh.- pT, El fran vieja^^ mira arriba papa
                      
160                      endArray(MemberIndex) = CByte(TempInt)
170                  End If

                     
180              End If
190           End If
200       Next j
          
210       stringToArray = endArray()

End Function

Private Function findMemberIndex(ByRef mName As String) As Integer

          '
          ' @ maTih.-
          
          ' default ;
10        findMemberIndex = 0
          
          Dim j As Long
          Dim N As Integer
          
20        N = NameIndex(mName)
          
30        If (N = 0) Then Exit Function
          
40        For j = 1 To PARTY_MAXMEMBERS
50            If p_members(j).Userindex = N Then
60                findMemberIndex = CInt(j)
70                Exit Function
80            End If
90        Next j

End Function

Public Sub setPorcentajes(ByRef bPorcentajeArray() As Byte)

          '
          ' @ maTih.-
          
          Dim j As Long
          Dim N As Integer
          Dim f As String
          
10        For j = 1 To PARTY_MAXMEMBERS
20            p_members(j).bPorcentaje = bPorcentajeArray(j)
              
30            N = p_members(j).Userindex
              
40            If (N > 0) Then
50                If (f = vbNullString) Then
60                    f = UserList(N).Name & " [" & p_members(j).bPorcentaje & "%] "
70                Else
80                    f = f & UserList(N).Name & " [" & p_members(j).bPorcentaje & "%] "
90                End If
100           End If
110       Next j
          
120       Call MandarMensajeAConsola("Porcentajes de la party actualizados." & vbNewLine & f & ".", "PARTY")
          
End Sub

Private Function getPorcentaje(ByVal bNumUsers As Byte) As Byte

          '
          ' @ maTih.-
          
          Dim bValue As Byte
          
10        Select Case bNumUsers
                 Case 1
20                    bValue = 100
                      
30               Case 2
40                    bValue = 50
                  
50               Case 3
60                    bValue = 33
                  
70               Case 4
80                    bValue = 25
                      
90               Case 5
100                   bValue = 20
110       End Select
          
120       getPorcentaje = bValue

End Function


Public Sub MandarMensajeAConsola(ByVal texto As String, ByVal Sender As String)
      'feo feo, muy feo acceder a senddata desde aca, pero BUEEEEEEEEEEE...
      Dim i As Integer

10    For i = 1 To PARTY_MAXMEMBERS
20        If p_members(i).Userindex > 0 Then
30            Call WriteConsoleMsg(p_members(i).Userindex, " [" & Sender & "] " & texto, FontTypeNames.FONTTYPE_PARTY)
40        End If
50    Next i

End Sub

Public Function EsPartyLeader(ByVal Userindex As Integer) As Boolean
10        EsPartyLeader = (Userindex = p_Fundador)
End Function

Public Function NuevoMiembro(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/08
      'Last Modification By: Marco Vanotti (MarKoxX)
      ' - 09/29/07 There is no level prohibition (Tavo)
      ' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
      '***************************************************

      Dim i As Integer
10    i = 1
20    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex > 0
30        i = i + 1
40    Wend

50    If i <= PARTY_MAXMEMBERS Then
60        p_members(i).Experiencia = 0
70        p_members(i).Userindex = Userindex
80        NuevoMiembro = True
90        p_CantMiembros = p_CantMiembros + 1
100       p_SumaNivelesElevados = p_SumaNivelesElevados + (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
          
110       Call initPorcentajes
120   Else
130       NuevoMiembro = False
140   End If

End Function

Public Function SaleMiembro(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/08
      'Last Modification By: Marco Vanotti (MarKoxX)
      ' - 09/29/07 Experience is round to the biggest number less than that number
      ' - 09/29/07 Now experience is a real-number (Tavo)
      ' - 07/04/08 Added const ExponenteNivelParty. (MarKoxX)
      '***************************************************
      'el valor de retorno representa si se disuelve la party
      Dim i As Integer
      Dim j As Integer
10    i = 1
20    SaleMiembro = False
30    While i <= PARTY_MAXMEMBERS And p_members(i).Userindex <> Userindex
40        i = i + 1
50    Wend

60    If i = 1 Then
          'sale el founder, la party se disuelve
70        SaleMiembro = True
80        Call MandarMensajeAConsola("El lider disuelve la party.", "Servidor")
90        For j = PARTY_MAXMEMBERS To 1 Step -1
100           If p_members(j).Userindex > 0 Then
                  
110               Call WriteConsoleMsg(p_members(j).Userindex, " Abandonas la party liderada por " & UserList(p_members(1).Userindex).Name, FontTypeNames.FONTTYPE_PARTY)
          
120               Call WriteConsoleMsg(p_members(j).Userindex, " Durante la misma has conseguido " & CStr(Fix(p_members(j).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
          
130               If Not PARTY_EXPERIENCIAPORGOLPE Then
140                   UserList(p_members(j).Userindex).Stats.Exp = UserList(p_members(j).Userindex).Stats.Exp + Fix(p_members(j).Experiencia)
150                   If UserList(p_members(j).Userindex).Stats.Exp > MAXEXP Then _
                          UserList(p_members(j).Userindex).Stats.Exp = MAXEXP
160                   Call CheckUserLevel(p_members(j).Userindex)
170                   Call WriteUpdateUserStats(p_members(j).Userindex)
180               End If
190               Call MandarMensajeAConsola(UserList(p_members(j).Userindex).Name & " abandona la party.", "Servidor")
200               UserList(p_members(j).Userindex).PartyIndex = 0
210               p_CantMiembros = p_CantMiembros - 1
220               p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
230               p_members(j).Userindex = 0
240               p_members(j).Experiencia = 0
250           End If
260       Next j
270   Else
280       If i <= PARTY_MAXMEMBERS Then
290           If Not PARTY_EXPERIENCIAPORGOLPE Then
300               UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
310               If UserList(p_members(i).Userindex).Stats.Exp > MAXEXP Then _
                      UserList(p_members(i).Userindex).Stats.Exp = MAXEXP
320               Call CheckUserLevel(p_members(i).Userindex)
330               Call WriteUpdateUserStats(p_members(i).Userindex)
340           End If
350           Call MandarMensajeAConsola(UserList(p_members(i).Userindex).Name & " abandona la party.", "Servidor")
              'TODO: Revisar que esto este bien, y no este faltando/sobrando un mensaje, ahora solo los estoy corrigiendo
360           Call WriteConsoleMsg(p_members(i).Userindex, " Durante la misma has conseguido " & CStr(Fix(p_members(i).Experiencia)) & " puntos de experiencia!", FontTypeNames.FONTTYPE_PARTY)
370           p_CantMiembros = p_CantMiembros - 1
380           p_SumaNivelesElevados = p_SumaNivelesElevados - (UserList(Userindex).Stats.ELV ^ ExponenteNivelParty)
390           p_members(i).Userindex = 0
400           p_members(i).Experiencia = 0
410           CompactMemberList
420       End If
430   End If

End Function

Public Function HacerLeader(ByVal Userindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 09/29/07
      'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
      ' - 09/29/07 There is no level prohibition
      '***************************************************
      Dim i As Integer
      Dim OldLeader As Integer
      Dim oldExp As Double
      Dim UserIndexIndex As Integer

10    UserIndexIndex = 0
20    HacerLeader = True

30    For i = 1 To PARTY_MAXMEMBERS
40        If p_members(i).Userindex > 0 Then
50            If p_members(i).Userindex = Userindex Then
60                UserIndexIndex = i
70            End If
80        End If
90    Next i

100   If Not HacerLeader Then Exit Function

110   If UserIndexIndex = 0 Then
          'catastrofe! esto no deberia pasar nunca! pero como es AO.... :p
120       Call LogError("INCONSISTENCIA DE PARTIES")
130       Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(" Inconsistencia de parties en HACERLEADER (UII = 0), AVISE A UN PROGRAMADOR ESTO ES UNA CATASTROFE!!!!", FontTypeNames.FONTTYPE_GUILD))
140       HacerLeader = False
150       Exit Function
160   End If


      'aca esta todo bien y doy vuelta las collections
170   OldLeader = p_members(1).Userindex
180   oldExp = p_members(1).Experiencia

190   p_members(1).Userindex = p_members(UserIndexIndex).Userindex     'que en realdiad es el userindex, pero no quiero inconsistencias moviendo experiencias
200   p_members(1).Experiencia = p_members(UserIndexIndex).Experiencia

210   p_members(UserIndexIndex).Userindex = OldLeader
220   p_members(UserIndexIndex).Experiencia = oldExp

230   p_Fundador = p_members(1).Userindex

      'no need to compact
End Function


Public Sub ObtenerMiembrosOnline(ByRef MemberList() As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 09/29/07
      'Last Modification By: Marco Vanotti (MarKoxX)
      ' - 09/29/07 Experience is round to the biggest number less than that number
      ' - 09/29/07 Now experience is a real-number (Tavo)
      ' - 08/18/08 Now TotalExperience is fixed (MarKoxX)
      ' - 11/27/09 Rehice la función, ahora devuelve el array con los UI online (Budi)
      '***************************************************

      Dim i As Integer

10    For i = 1 To PARTY_MAXMEMBERS
20        If p_members(i).Userindex > 0 Then
30            MemberList(i) = p_members(i).Userindex
40        End If
50    Next i

End Sub

Public Function ObtenerExperienciaTotal() As Long
      '***************************************************
      'Author: Budi
      'Last Modification: 11/27/09
      'Retrieves the total experience acumulated in the party
      '***************************************************
10    ObtenerExperienciaTotal = p_expTotal
End Function

Public Function PuedeEntrar(ByVal Userindex As Integer, ByRef razon As String) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 09/29/07
      'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
      ' - 09/29/07 There is no level prohibition
      '***************************************************
      'DEFINE LAS REGLAS DEL JUEGO PARA DEJAR ENTRAR A MIEMBROS
      Dim esArmada As Boolean
      Dim esCaos As Boolean
      Dim MyLevel As Integer
      Dim i As Integer
      Dim rv As Boolean
      Dim UI As Integer

10    rv = True
20    esArmada = (UserList(Userindex).Faccion.ArmadaReal = 1)
30    esCaos = (UserList(Userindex).Faccion.FuerzasCaos = 1)
40    MyLevel = UserList(Userindex).Stats.ELV

50    rv = Distancia(UserList(p_members(1).Userindex).Pos, UserList(Userindex).Pos) <= MAXDISTANCIAINGRESOPARTY
60    If rv Then
70        rv = (p_members(PARTY_MAXMEMBERS).Userindex = 0)
80        If rv Then
90            For i = 1 To PARTY_MAXMEMBERS
100               UI = p_members(i).Userindex
                  'pongo los casos que evitarian que pueda entrar
                  'aspirante armada en party crimi
110               If UI > 0 Then
120                   If esArmada And criminal(UI) Then
130                       razon = "Los miembros de la armada no entran a una party con criminales."
140                       rv = False
150                   End If
                      'aspirante caos en party ciuda
160                   If esCaos And Not criminal(UI) Then
170                       razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
180                       rv = False
190                   End If
                      'aspirante crimi en party armada
200                   If UserList(UI).Faccion.ArmadaReal = 1 And criminal(Userindex) Then
210                       razon = "Los criminales no entran a partys con miembros de la armada real."
220                       rv = False
230                   End If
                      'aspirante ciuda en party caos
240                   If UserList(UI).Faccion.FuerzasCaos = 1 And Not criminal(Userindex) Then
250                       razon = "Los miembros de la legión oscura no entran a una party con ciudadanos."
260                       rv = False
270                   End If
                  
280                   If Not rv Then Exit For 'violate una programacion estructurada
290               End If
300           Next i
310       Else
320           razon = "La mayor cantidad de miembros es " & PARTY_MAXMEMBERS
330       End If
340   Else
350       razon = "Te encuentras muy lejos del fundador."
360   End If

370   PuedeEntrar = rv

End Function


Public Sub FlushExperiencia()
      '***************************************************
      'Author: Unknown
      'Last Modification: 09/29/07
      'Last Modification By: Lucas Tavolaro Ortiz (Tavo)
      ' - 09/29/07 Experience is round to the biggest number less than that number
      ' - 09/29/07 Now experience is a real-number
      '***************************************************
      'esta funcion se invoca frente a cerradas del servidor. Flushea la experiencia
      'acumulada a los usuarios.

      Dim i As Integer
10    If Not PARTY_EXPERIENCIAPORGOLPE Then   'esto sirve SOLO cuando acumulamos la experiencia!
20        For i = 1 To PARTY_MAXMEMBERS
30            If p_members(i).Userindex > 0 Then
40                If p_members(i).Experiencia > 0 Then
50                    UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
60                    If UserList(p_members(i).Userindex).Stats.Exp > MAXEXP Then _
                          UserList(p_members(i).Userindex).Stats.Exp = MAXEXP
70                    Call CheckUserLevel(p_members(i).Userindex)
80                Else
90                    If Abs(UserList(p_members(i).Userindex).Stats.Exp) > Abs(Fix(p_members(i).Experiencia)) Then
100                       UserList(p_members(i).Userindex).Stats.Exp = UserList(p_members(i).Userindex).Stats.Exp + Fix(p_members(i).Experiencia)
110                   Else
120                       UserList(p_members(i).Userindex).Stats.Exp = 0
130                   End If
140               End If
150               p_members(i).Experiencia = 0
160               Call WriteUpdateUserStats(p_members(i).Userindex)
170           End If
180       Next i
190   End If

End Sub

Private Sub CompactMemberList()
      Dim i As Integer
      Dim freeIndex As Integer
10    i = 1
20    While i <= PARTY_MAXMEMBERS
30        If p_members(i).Userindex = 0 And freeIndex = 0 Then
40            freeIndex = i
50        ElseIf p_members(i).Userindex > 0 And freeIndex > 0 Then
60            p_members(freeIndex).Experiencia = p_members(i).Experiencia
70            p_members(freeIndex).Userindex = p_members(i).Userindex
80            p_members(i).Userindex = 0
90            p_members(i).Experiencia = 0
              'muevo el de la pos i a freeindex
100           i = freeIndex
110           freeIndex = 0
120       End If
130       i = i + 1
140   Wend

End Sub

Public Function CantMiembros() As Integer
10    CantMiembros = p_CantMiembros
End Function




