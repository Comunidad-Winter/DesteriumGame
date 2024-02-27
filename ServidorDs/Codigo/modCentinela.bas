Attribute VB_Name = "modCentinela"
'*****************************************************************
'modCentinela.bas - ImperiumAO - v1.2
'
'Funciónes de control para usuarios que se encuentran trabajando
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
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

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   ImperiumAO 1.2
'   - First Relase
'
'Juan Martín Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.11.5
'   - Small improvements and added logs to detect possible cheaters
'
'Juan Martín Sotuyo Dodero (juansotuyo@gmail.com)
'   Alkon AO 0.12.0
'   - Added several messages to spam users until they reply
'*****************************************************************

Option Explicit

Private Const NPC_CENTINELA_TIERRA As Integer = 117  'Índice del NPC en el .dat
Private Const NPC_CENTINELA_AGUA As Integer = 117    'Ídem anterior, pero en mapas de agua

Public CentinelaNPCIndex As Integer                'Índice del NPC en el servidor

Private Const TIEMPO_INICIAL As Byte = 2 'Tiempo inicial en minutos. No reducir sin antes revisar el timer que maneja estos datos.

Private Type tCentinela
    RevisandoUserIndex As Integer   '¿Qué índice revisamos?
    TiempoRestante As Integer       '¿Cuántos minutos le quedan al usuario?
    clave As Integer                'Clave que debe escribir
    spawnTime As Long
End Type

Public centinelaActivado As Boolean

Public Centinela As tCentinela

Public Sub CallUserAttention()
      '############################################################
      'Makes noise and FX to call the user's attention.
      '############################################################
10        If (GetTickCount() And &H7FFFFFFF) - Centinela.spawnTime >= 5000 Then
20            If Centinela.RevisandoUserIndex <> 0 And centinelaActivado Then
30                If Not UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then
40                    Call WritePlayWave(Centinela.RevisandoUserIndex, SND_WARP, Npclist(CentinelaNPCIndex).Pos.X, Npclist(CentinelaNPCIndex).Pos.Y)
50                    Call WriteCreateFX(Centinela.RevisandoUserIndex, Npclist(CentinelaNPCIndex).Char.CharIndex, FXIDs.FXWARP, 0)
                      
                      'Resend the key
60                    Call CentinelaSendClave(Centinela.RevisandoUserIndex)
                      
70                    Call FlushBuffer(Centinela.RevisandoUserIndex)
80                End If
90            End If
100       End If
End Sub

Private Sub GoToNextWorkingChar()
      '############################################################
      'Va al siguiente usuario que se encuentre trabajando
      '############################################################
          Dim LoopC As Long
          
10        For LoopC = 1 To LastUser
20            If UserList(LoopC).flags.UserLogged And UserList(LoopC).Counters.Trabajando > 0 And (UserList(LoopC).flags.Privilegios And PlayerType.User) Then
30                If Not UserList(LoopC).flags.CentinelaOK Then
                      'Inicializamos
40                    Centinela.RevisandoUserIndex = LoopC
50                    Centinela.TiempoRestante = TIEMPO_INICIAL
60                    Centinela.clave = RandomNumber(1, 32000)
70                    Centinela.spawnTime = GetTickCount() And &H7FFFFFFF
                      
                      'Ponemos al centinela en posición
80                    Call WarpCentinela(LoopC)
                      
90                    If CentinelaNPCIndex Then
                          'Mandamos el mensaje (el centinela habla y aparece en consola para que no haya dudas)
100                       Call WriteChatOverHead(LoopC, "Saludos " & UserList(LoopC).Name & ", soy el Centinela de estas tierras. Me gustaría que escribas /CENTINELA " & Centinela.clave & " en no más de dos minutos.", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbGreen)
110                       Call WriteConsoleMsg(LoopC, "El centinela intenta llamar tu atención. ¡Respóndele rápido!", FontTypeNames.FONTTYPE_CENTINELA)
120                       Call FlushBuffer(LoopC)
130                   End If
140                   Exit Sub
150               End If
160           End If
170       Next LoopC
          
          'No hay chars trabajando, eliminamos el NPC si todavía estaba en algún lado y esperamos otro minuto
180       If CentinelaNPCIndex Then
190           Call QuitarNPC(CentinelaNPCIndex)
200           CentinelaNPCIndex = 0
210       End If
          
          'No estamos revisando a nadie
220       Centinela.RevisandoUserIndex = 0
End Sub

Private Sub CentinelaFinalCheck()
      '############################################################
      'Al finalizar el tiempo, se retira y realiza la acción
      'pertinente dependiendo del caso
      '############################################################
10    On Error GoTo Error_Handler
          Dim Name As String
          Dim numPenas As Integer
          
20        If Not UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK Then
              'Logueamos el evento
30            Call LogCentinela("Centinela baneo a " & UserList(Centinela.RevisandoUserIndex).Name & " por uso de macro inasistido.")
              
              'Ponemos el ban
40            UserList(Centinela.RevisandoUserIndex).flags.Ban = 1
              
50            Name = UserList(Centinela.RevisandoUserIndex).Name
              
              'Avisamos a los admins
60            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> El centinela ha baneado a " & Name, FontTypeNames.FONTTYPE_SERVER))
              
              'ponemos el flag de ban a 1
70            Call WriteVar(CharPath & Name & ".chr", "FLAGS", "Ban", "1")
              'ponemos la pena
80            numPenas = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
90            Call WriteVar(CharPath & Name & ".chr", "PENAS", "Cant", numPenas + 1)
100           Call WriteVar(CharPath & Name & ".chr", "PENAS", "P" & numPenas + 1, "CENTINELA : BAN POR MACRO INASISTIDO " & Date & " " & time)
              
              'Evitamos loguear el logout
              Dim index As Integer
110           index = Centinela.RevisandoUserIndex
120           Centinela.RevisandoUserIndex = 0
              
130           Call CloseSocket(index)
140       End If
          
150       Centinela.clave = 0
160       Centinela.TiempoRestante = 0
170       Centinela.RevisandoUserIndex = 0
          
180       If CentinelaNPCIndex Then
190           Call QuitarNPC(CentinelaNPCIndex)
200           CentinelaNPCIndex = 0
210       End If
220   Exit Sub

Error_Handler:
230       Centinela.clave = 0
240       Centinela.TiempoRestante = 0
250       Centinela.RevisandoUserIndex = 0
          
260       If CentinelaNPCIndex Then
270           Call QuitarNPC(CentinelaNPCIndex)
280           CentinelaNPCIndex = 0
290       End If
          
300       Call LogError("Error en el checkeo del centinela: " & Err.Description)
End Sub

Public Sub CentinelaCheckClave(ByVal Userindex As Integer, ByVal clave As Integer)
      '############################################################
      'Corrobora la clave que le envia el usuario
      '############################################################
10        If clave = Centinela.clave And Userindex = Centinela.RevisandoUserIndex Then
20            UserList(Centinela.RevisandoUserIndex).flags.CentinelaOK = True
30            Call WriteChatOverHead(Userindex, "¡Muchas gracias " & UserList(Centinela.RevisandoUserIndex).Name & "! Espero no haber sido una molestia.", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbWhite)
40            Centinela.RevisandoUserIndex = 0
50            Call FlushBuffer(Userindex)
60        Else
70            Call CentinelaSendClave(Userindex)
              
              'Logueamos el evento
80            If Userindex <> Centinela.RevisandoUserIndex Then
90                Call LogCentinela("El usuario " & UserList(Userindex).Name & " respondió aunque no se le hablaba a él.")
100           Else
110               Call LogCentinela("El usuario " & UserList(Userindex).Name & " respondió una clave incorrecta: " & clave & " - Se esperaba : " & Centinela.clave)
120           End If
130       End If
End Sub

Public Sub ResetCentinelaInfo()
      '############################################################
      'Cada determinada cantidad de tiempo, volvemos a revisar
      '############################################################
          Dim LoopC As Long
          
10        For LoopC = 1 To LastUser
20            If (LenB(UserList(LoopC).Name) <> 0 And LoopC <> Centinela.RevisandoUserIndex) Then
30                UserList(LoopC).flags.CentinelaOK = False
40            End If
50        Next LoopC
End Sub

Public Sub CentinelaSendClave(ByVal Userindex As Integer)
      '############################################################
      'Enviamos al usuario la clave vía el personaje centinela
      '############################################################
10        If CentinelaNPCIndex = 0 Then Exit Sub
          
20        If Userindex = Centinela.RevisandoUserIndex Then
30            If Not UserList(Userindex).flags.CentinelaOK Then
40                Call WriteChatOverHead(Userindex, "¡La clave que te he dicho es /CENTINELA " & Centinela.clave & ", escríbelo rápido!", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbGreen)
50                Call WriteConsoleMsg(Userindex, "El centinela intenta llamar tu atención. ¡Respondele rápido!", FontTypeNames.FONTTYPE_CENTINELA)
60            Else
                  'Logueamos el evento
70                Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).Name & " respondió más de una vez la contraseña correcta.")
80                Call WriteChatOverHead(Userindex, "Te agradezco, pero ya me has respondido. Me retiraré pronto.", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbGreen)
90            End If
100       Else
110           Call WriteChatOverHead(Userindex, "No es a ti a quien estoy hablando, ¿No ves?", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbWhite)
120       End If
End Sub

Public Sub PasarMinutoCentinela()
      '############################################################
      'Control del timer. Llamado cada un minuto.
      '############################################################
10        If Not centinelaActivado Then Exit Sub
          
20        If Centinela.RevisandoUserIndex = 0 Then
30            Call GoToNextWorkingChar
40        Else
50            Centinela.TiempoRestante = Centinela.TiempoRestante - 1
              
60            If Centinela.TiempoRestante = 0 Then
70                Call CentinelaFinalCheck
80                Call GoToNextWorkingChar
90            Else
                  'Recordamos al user que debe escribir
100               If Matematicas.Distancia(Npclist(CentinelaNPCIndex).Pos, UserList(Centinela.RevisandoUserIndex).Pos) > 5 Then
110                   Call WarpCentinela(Centinela.RevisandoUserIndex)
120               End If
                  
                  'El centinela habla y se manda a consola para que no quepan dudas
130               Call WriteChatOverHead(Centinela.RevisandoUserIndex, "¡" & UserList(Centinela.RevisandoUserIndex).Name & ", tienes un minuto más para responder! Debes escribir /CENTINELA " & Centinela.clave & ".", CStr(Npclist(CentinelaNPCIndex).Char.CharIndex), vbRed)
140               Call WriteConsoleMsg(Centinela.RevisandoUserIndex, "¡" & UserList(Centinela.RevisandoUserIndex).Name & ", tienes un minuto más para responder!", FontTypeNames.FONTTYPE_CENTINELA)
150               Call FlushBuffer(Centinela.RevisandoUserIndex)
160           End If
170       End If
End Sub

Private Sub WarpCentinela(ByVal Userindex As Integer)
      '############################################################
      'Inciamos la revisión del usuario UserIndex
      '############################################################
          'Evitamos conflictos de índices
10        If CentinelaNPCIndex Then
20            Call QuitarNPC(CentinelaNPCIndex)
30            CentinelaNPCIndex = 0
40        End If
          
50        If HayAgua(UserList(Userindex).Pos.map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) Then
60            CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_AGUA, UserList(Userindex).Pos, True, False)
70        Else
80            CentinelaNPCIndex = SpawnNpc(NPC_CENTINELA_TIERRA, UserList(Userindex).Pos, True, False)
90        End If
          
          'Si no pudimos crear el NPC, seguimos esperando a poder hacerlo
100       If CentinelaNPCIndex = 0 Then _
              Centinela.RevisandoUserIndex = 0
End Sub

Public Sub CentinelaUserLogout()
      '############################################################
      'El usuario al que revisabamos se desconectó
      '############################################################
10        If Centinela.RevisandoUserIndex Then
              'Logueamos el evento
20            Call LogCentinela("El usuario " & UserList(Centinela.RevisandoUserIndex).Name & " se desolgueó al pedirsele la contraseña.")
              
              'Reseteamos y esperamos a otro PasarMinuto para ir al siguiente user
30            Centinela.clave = 0
40            Centinela.TiempoRestante = 0
50            Centinela.RevisandoUserIndex = 0
              
60            If CentinelaNPCIndex Then
70                Call QuitarNPC(CentinelaNPCIndex)
80                CentinelaNPCIndex = 0
90            End If
100       End If
End Sub

Private Sub LogCentinela(ByVal texto As String)
      '*************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last modified: 03/15/2006
      'Loguea un evento del centinela
      '*************************************************
10    On Error GoTo Errhandler

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
          
30        Open App.Path & "\logs\Centinela.log" For Append Shared As #nfile
40        Print #nfile, Date & " " & time & " " & texto
50        Close #nfile
60    Exit Sub

Errhandler:
End Sub
