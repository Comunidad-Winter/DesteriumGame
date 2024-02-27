Attribute VB_Name = "Mod_AntiEdic"
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

Rem Autor Miqueas
Rem Para Hispano AO
Rem Creado el dia 31/01/14
Rem Controla todo lo que Intervalos
Rem Si un usuario lo supera avisa por consola a los GMs y crea un log por si las dudas

Option Explicit

'// Type de los counters para tener todo bien rodenado

Public Type TimeIntervalos

        UsarItem As Byte
        AtacaArco As Byte
        AtacaComun As Byte
        CastSpell As Byte

End Type

Public Sub ResetAllCount(ByVal Userindex As Integer)

              '// Miqueas, el que lea esto es puto(Solo se aplica a Godker) ...
              '// Reseteamos Todos los counter correspondiente

10            With UserList(Userindex)
              
20                    If (.Counters.Cheat.AtacaArco <> 0) Then
30                            .Counters.Cheat.AtacaArco = 0
40                    End If

50                    If (.Counters.Cheat.AtacaComun <> 0) Then
60                            .Counters.Cheat.AtacaComun = 0
70                    End If

80                    If (.Counters.Cheat.CastSpell <> 0) Then
90                            .Counters.Cheat.CastSpell = 0
100                   End If

110                   If (.Counters.Cheat.UsarItem <> 0) Then
120                           .Counters.Cheat.UsarItem = 0
130                   End If
                      
140           End With

End Sub

Public Sub RestaCount(ByVal Userindex As Integer, _
                      Optional ByVal Flecha As Byte = 0, _
                      Optional ByVal Golpe As Byte = 0, _
                      Optional ByVal Cast As Byte = 0, _
                      Optional ByVal Usar As Byte = 0)
                            
              '// Miqueas, el que lea esto es puto(Solo se aplica a Godker) ...
              '// Reseteamos el counter correcto, segun se lo pidamos

10            With UserList(Userindex)

20                    If (Flecha <> 0) Then
30                            .Counters.Cheat.AtacaArco = 0
40                    End If

50                    If (Golpe <> 0) Then
60                            .Counters.Cheat.AtacaComun = 0
70                    End If

80                    If (Cast <> 0) Then
90                            .Counters.Cheat.CastSpell = 0
100                   End If

110                   If (Usar <> 0) Then
120                           .Counters.Cheat.UsarItem = 0
130                   End If

140           End With

End Sub

Public Sub AddCount(ByVal Userindex As Integer, _
                    Optional ByVal AddFlecha As Byte = 0, _
                    Optional ByVal AddGolpe As Byte = 0, _
                    Optional ByVal AddCast As Byte = 0, _
                    Optional ByVal AddUsar As Byte = 0)
                          
              '// Miqueas, el que lea esto es puto(Solo se aplica a Godker) ...
              '// Sumamos al counter correspondiente

              Dim Msj As String

10            With UserList(Userindex)

20                    If (AddFlecha <> 0) Then
30                            .Counters.Cheat.AtacaArco = (.Counters.Cheat.AtacaArco + 1)

40                            If CheckInt(Userindex, Msj, 1) Then
50                                    MsjCheat Userindex, Msj
60                            End If
                              
70                    End If

80                    If (AddGolpe <> 0) Then
90                            .Counters.Cheat.AtacaComun = (.Counters.Cheat.AtacaComun + 1)

100                           If CheckInt(Userindex, Msj, 2) Then
110                                   MsjCheat Userindex, Msj
120                           End If
                              
130                   End If
             
140                   If (AddCast <> 0) Then
150                           .Counters.Cheat.CastSpell = (.Counters.Cheat.CastSpell + 1)

160                           If CheckInt(Userindex, Msj, 3) Then
170                                   MsjCheat Userindex, Msj
180                           End If
                              
190                   End If

200                   If (AddUsar <> 0) Then
210                           .Counters.Cheat.UsarItem = (.Counters.Cheat.UsarItem + 1)

220                           If CheckInt(Userindex, Msj, 4) Then
230                                   MsjCheat Userindex, Msj
240                           End If
                              
250                   End If
                      
260           End With
              
End Sub

Private Function CheckInt(ByVal Userindex As Integer, _
                          ByRef Msj As String, _
                          ByVal Intervalo As Byte) As Boolean

              '// Miqueas, el que lea esto es puto(Solo se aplica a Godker) ...
              '// Chekeamos los intervalos
              '// Tambien mandamos el msj correspondiente

              Const MaxTol As Byte = 3

10            With UserList(Userindex)

20                    Select Case Intervalo
              
                              Case 1
         
30                                    If (.Counters.Cheat.AtacaArco = MaxTol) Then
40                                            Msj = .Name & ". -" & "Sobrepaso el intervalo de Ataca Arco  " & MaxTol & " veces seguidas." & vbNewLine & "Posible edicion de intervalos."
50                                            .Counters.Cheat.AtacaArco = 0
60                                            CheckInt = True

70                                            Exit Function

80                                    End If

90                            Case 2

100                                   If (.Counters.Cheat.AtacaComun = MaxTol) Then
110                                           Msj = .Name & ". -" & "Sobrepaso el intervalo de Ataca Comun  " & MaxTol & " veces seguidas." & vbNewLine & "Posible edicion de intervalos."
120                                           .Counters.Cheat.AtacaComun = 0
130                                           CheckInt = True
        
140                                           Exit Function

150                                   End If

160                           Case 3

170                                   If (.Counters.Cheat.CastSpell = MaxTol) Then
180                                           Msj = .Name & ". -" & "Sobrepaso el intervalo de Cast Spell " & MaxTol & " veces seguidas." & vbNewLine & "Posible edicion de intervalos."
190                                           .Counters.Cheat.CastSpell = 0
200                                           CheckInt = True

210                                           Exit Function

220                                   End If

230                           Case 4

240                                   If (.Counters.Cheat.UsarItem = MaxTol) Then
250                                           Msj = .Name & ". -" & "Sobrepaso el intervalo de Usar Items " & MaxTol & " veces seguidas." & vbNewLine & "Posible edicion de intervalos."
260                                           .Counters.Cheat.UsarItem = 0
270                                           CheckInt = True
       
280                                           Exit Function

290                                   End If
           
300                   End Select
              
310           End With

320           CheckInt = False
              
End Function

Private Sub MsjCheat(ByVal Userindex As Integer, ByVal Msj As String)

              '// Autor : Miqueas
              '// Mandamos el msj y guardamos en un log el msj por si no estan los GMs del AO Online

              Dim sndData As String

10            With UserList(Userindex)
              
20                    sndData = PrepareMessageConsoleMsg(.Name & Msj, FontTypeNames.FONTTYPE_SERVER)
                      
30                    Call SendData(SendTarget.ToAdmins, 0, sndData)
                      
                      '// Guardamos un log
40                    Call LogIntervalos(.Name, Msj)
                                      
50            End With

End Sub

Private Sub LogIntervalos(ByVal Nombre As String, ByVal Str As String)

10            On Error GoTo Errhandler

              Dim nfile As Integer

20            nfile = FreeFile ' obtenemos un canal
              
30            Open App.Path & "\AntiCheats\" & Nombre & ".log" For Append Shared As #nfile
40            Print #nfile, Date$ & " " & time$ & " " & Str
50            Close #nfile
          
60            Exit Sub

Errhandler:

End Sub


