VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsultasPopulares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' ConsultasPopulares.cls
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
'Modulo de consultas popluares
'En realidad el modulo inicial guardaba los datos de las votaciones
'en memoria, pero no le vi el punto, las votaciones son de vez en cuando nomas
'y el query a un .ini que registra todos los mails q ya votaron, es soportable
'si cuando uno vota y averiguo que el mail ya voto pero el PJ no, entonces seteo
'el flag de yavoto en el charfile ese tambien,
'por lo que la busqueda secuencial en el .dat que tiene todos los mails q ya votaron
'se hara .. 1 vez por PJ nomas.
'
'Hecha por el oso

Private Const ARCHIVOMAILS = "\logs\votaron.dat"
Private Const ARCHIVOCONFIG = "\dat\consultas.dat"

Private pEncuestaActualNum As Integer
Private pEncuestaActualTex As String
Private pNivelRequerido As Integer
Private pOpciones() As Integer


Public Property Let Numero(ByVal NumEncuesta As Integer)
10        pEncuestaActualNum = NumEncuesta
End Property

Public Property Get Numero() As Integer
10        Numero = pEncuestaActualNum
End Property

Public Property Let texto(ByVal Descripcion As String)
10        pEncuestaActualTex = Descripcion
End Property

Public Property Get texto() As String
10        texto = pEncuestaActualTex
End Property


Public Sub LoadData()
      Dim CantOpciones As Integer
      Dim i As Integer

10        pEncuestaActualNum = val(GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActual"))
20        pEncuestaActualTex = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "ConsultaActualTexto")
30        pNivelRequerido = GetVar(App.Path & ARCHIVOCONFIG, "INIT", "NivelRequerido")

40        If pEncuestaActualNum > 0 Then
             'cargo todas las opciones
50            CantOpciones = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "CANTOPCIONES"))
60            ReDim pOpciones(1 To CantOpciones) As Integer
70            For i = 1 To CantOpciones
80                pOpciones(i) = val(GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i))
90            Next i
100       End If
End Sub

Public Function doVotar(ByVal Userindex As Integer, ByVal opcion As Integer) As String
10    On Error GoTo errorh
      Dim YaVoto As Boolean
      Dim CharFile As String
      Dim sufragio As Integer

          'revisar q no haya votado
          'grabar en el charfile el numero de encuesta
          'actualizar resultados encuesta
20        If pEncuestaActualNum = 0 Then
30            doVotar = "No hay consultas populares abiertas"
40            Exit Function
50        End If

60        CharFile = CharPath & UserList(Userindex).Name & ".chr"
          
          
          
70        If (UserList(Userindex).Stats.ELV >= pNivelRequerido) Then
80            If (OpcionValida(opcion)) Then
90                YaVoto = val(GetVar(CharFile, "CONSULTAS", "Voto")) >= pEncuestaActualNum
100               If Not YaVoto Then
110                   If Not MailYaVoto(UserList(Userindex).email) Then
                          'pj apto para votar
120                       sufragio = CLng(val(GetVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion)))
130                       sufragio = sufragio + 1
140                       Call WriteVar(App.Path & ARCHIVOCONFIG, "RESULTADOS" & pEncuestaActualNum, "V" & opcion, Str(sufragio))
150                       doVotar = "Tu voto ha sido computado. Opcion: " & opcion
160                       Call MarcarPjComoQueYaVoto(Userindex)
170                       Call MarcarMailComoQueYaVoto(UserList(Userindex).email)
180                   Else
190                       Call MarcarPjComoQueYaVoto(Userindex)
200                       doVotar = "Este email ya voto en la consulta: " & pEncuestaActualTex
210                   End If
220               Else
230                   doVotar = "Este personaje ya voto en la consulta: " & pEncuestaActualTex
240               End If
250           Else
260               doVotar = "Esa no es una opcion para votar"
270           End If
280       Else
290           doVotar = "Para votar en esta consulta debes ser nivel " & pNivelRequerido & " o superior"
300       End If
         

310   Exit Function
errorh:
320   Call LogError("Error en ConsultasPopularse.doVotar: " & Err.Description)

End Function


Public Function SendInfoEncuesta(ByVal Userindex As Integer) As String
      Dim i As Integer
10        Call WriteConsoleMsg(Userindex, "CONSULTA POPULAR NUMERO " & pEncuestaActualNum, FontTypeNames.FONTTYPE_GUILD)
20        Call WriteConsoleMsg(Userindex, pEncuestaActualTex, FontTypeNames.FONTTYPE_GUILD)
30        Call WriteConsoleMsg(Userindex, " Opciones de voto: ", FontTypeNames.FONTTYPE_GUILDMSG)
40        For i = 1 To UBound(pOpciones)
50            Call WriteConsoleMsg(Userindex, "(Opcion " & i & "): " & GetVar(App.Path & ARCHIVOCONFIG, "ENCUESTA" & pEncuestaActualNum, "OPCION" & i), FontTypeNames.FONTTYPE_GUILDMSG)
60        Next i
70        Call WriteConsoleMsg(Userindex, " Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1. Tu voto no podra ser cambiado.", FontTypeNames.FONTTYPE_VENENO)
End Function


Private Sub MarcarPjComoQueYaVoto(ByVal Userindex As Integer)
10        Call WriteVar(CharPath & UserList(Userindex).Name & ".chr", "CONSULTAS", "Voto", Str(pEncuestaActualNum))
End Sub


Private Function MailYaVoto(ByVal email As String) As Boolean
      'abro el archivo, while not eof levnato 1 linea y comparo. Si da true, cierro
      Dim ArchN As Integer
      Dim tmp As String

10    MailYaVoto = False

20    ArchN = FreeFile()

30    Open App.Path & ARCHIVOMAILS For Input As #ArchN

40    Do While Not EOF(ArchN)
50        Line Input #ArchN, tmp
60        If email = tmp Then
70            MailYaVoto = True
80            Close #ArchN
90            Exit Function
100       End If
110   Loop

120   Close #ArchN
End Function


Private Sub MarcarMailComoQueYaVoto(ByVal email As String)
      Dim ArchN As Integer

10    ArchN = FreeFile()

20    Open App.Path & ARCHIVOMAILS For Append As #ArchN
30    Print #ArchN, email

40    Close #ArchN

End Sub


Private Function OpcionValida(ByVal opcion As Integer) As Boolean
10        OpcionValida = opcion > 0 And opcion <= UBound(pOpciones)
End Function

