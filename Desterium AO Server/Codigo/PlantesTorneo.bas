Attribute VB_Name = "PlantesTorneo"
Option Explicit
' Codigo: Torneos planticos 100%
' Autor: Joan Calderón - SaturoS.
Public Plantes_Activo As Boolean
Public Plantes_Esperando As Boolean
Private Plantes_Rondas As Integer
Private Plantes_Luchadores() As Integer
 
Private Const mapatorneo As Integer = 199
Private Const mapaespera As Integer = 199
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 58
Private Const esquina1y As Integer = 86
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 61
Private Const esquina2y As Integer = 86
' Donde esperan los tios
Private Const esperax As Integer = 20
Private Const esperay As Integer = 44
' Mapa desconecta
Private Const mapa_fuera As Integer = 1
Private Const fueraesperay As Integer = 55
Private Const fueraesperax As Integer = 57
 ' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1 As Integer = 11
Private Const X2 As Integer = 27
Private Const Y1 As Integer = 42
Private Const Y2 As Integer = 48
 
 Sub Torneoauto_Cancela()
On Error GoTo errorh:
    If (Not Plantes_Activo And Not Plantes_Esperando) Then Exit Sub
    Plantes_Activo = False
    Plantes_Esperando = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Se canceló por falta de participantes.", FontTypeNames.FONTTYPE_GUILD))
    Dim i As Integer
     For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                If (Plantes_Luchadores(i) <> -1) Then
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Plantes_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                      UserList(Plantes_Luchadores(i)).flags.Plantico = False
                End If
        Next i
errorh:
End Sub
Sub RondasP_Cancela()
On Error GoTo errorh
    If (Not Plantes_Activo And Not Plantes_Esperando) Then Exit Sub
    Plantes_Activo = False
    Plantes_Esperando = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Cancelado por un Game Master.", FontTypeNames.FONTTYPE_GUILD))
    Dim i As Integer
    For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                If (Plantes_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Plantes_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                    UserList(Plantes_Luchadores(i)).flags.Plantico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_PlantadorMuere(ByVal Userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_plantadormuere_errorh
        Dim i As Integer, Pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not Plantes_Activo) Then
                Exit Sub
            ElseIf (Plantes_Activo And Plantes_Esperando) Then
                For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                    If (Plantes_Luchadores(i) = Userindex) Then
                        Plantes_Luchadores(i) = -1
                        Call WarpUserChar(Userindex, mapa_fuera, fueraesperay, fueraesperax, True)
                         UserList(Userindex).flags.Plantico = False
                        Exit Sub
                    End If
                Next i
                Exit Sub
            End If
 
        For Pos = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                If (Plantes_Luchadores(Pos) = Userindex) Then Exit For
        Next Pos
 
        ' si no lo ha encontrado
        If (Plantes_Luchadores(Pos) <> Userindex) Then Exit Sub
       
 '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
 
If UserList(Userindex).Pos.X >= X1 And UserList(Userindex).Pos.X <= X2 And UserList(Userindex).Pos.Y >= Y1 And UserList(Userindex).Pos.Y <= Y2 Then
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(Userindex).Name & " se fue del torneo mientras esperaba pelear!", FontTypeNames.FONTTYPE_GUILD))
Call WarpUserChar(Userindex, mapa_fuera, fueraesperax, fueraesperay, True)
UserList(Userindex).flags.Plantico = False
Plantes_Luchadores(Pos) = -1
Exit Sub
End If
 
        combate = 1 + (Pos - 1) \ 2
 
        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1) + 1
        LI2 = LI1 + 1
 
        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(Userindex).Name & " pierde el combate!", FontTypeNames.FONTTYPE_GUILD))
        Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(Userindex).Name & " se fue del combate!", FontTypeNames.FONTTYPE_GUILD))
        End If
 
        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(Userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                 UserList(Userindex).flags.Plantico = False
        ElseIf (Not CambioMapa) Then
             
                 Call WarpUserChar(Userindex, mapa_fuera, fueraesperax, fueraesperay, True)
                  UserList(Userindex).flags.Plantico = False
        End If
 
        'se le borra de la lista y se mueve el segundo a li1
        If (Plantes_Luchadores(LI1) = Userindex) Then
                Plantes_Luchadores(LI1) = Plantes_Luchadores(LI2) 'cambiamos slot
                Plantes_Luchadores(LI2) = -1
        Else
                Plantes_Luchadores(LI2) = -1
        End If
 
    'si es la ultima ronda
    If (Plantes_Rondas = 1) Then
        Call WarpUserChar(Plantes_Luchadores(LI1), mapa_fuera, 51, 51, True)
        'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(UserIndex).Name & " pierde el combate!", FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> El ganador del torneo es " & UserList(Plantes_Luchadores(LI1)).Name & " que se lleva un total de 1 punto de torneo y 100.000 monedas de oro. ¡FELICIDADES!", FontTypeNames.FONTTYPE_CONSEJOVesA))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Torneo finalizado. Gracias por participar.", FontTypeNames.FONTTYPE_GUILD))
        UserList(Plantes_Luchadores(LI1)).Stats.GLD = UserList(Plantes_Luchadores(LI1)).Stats.GLD + 100000
        UserList(Plantes_Luchadores(LI1)).Stats.TorneosGanados = UserList(Plantes_Luchadores(LI1)).Stats.TorneosGanados + 1
        Call WriteUpdateGold(Plantes_Luchadores(LI1))
        UserList(Plantes_Luchadores(LI1)).flags.Plantico = False
        Plantes_Activo = False
       
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Plantes_Luchadores(LI1), mapaespera, esperax, esperay, True)
    End If
 
               
        'si es el ultimo combate de la ronda
        If (2 ^ Plantes_Rondas = 2 * combate) Then
               
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Siguiente ronda!", FontTypeNames.FONTTYPE_GUILD))
                Plantes_Rondas = Plantes_Rondas - 1
 
        'antes de llamar a la proxima ronda hay q copiar a los tipos
        For i = 1 To 2 ^ Plantes_Rondas
                UI1 = Plantes_Luchadores(2 * (i - 1) + 1)
                UI2 = Plantes_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Plantes_Luchadores(i) = UI1
        Next i
ReDim Preserve Plantes_Luchadores(1 To 2 ^ Plantes_Rondas) As Integer
        Call Rondas_Combatep(1)
        Exit Sub
        End If
 
        'vamos al siguiente combate
        Call Rondas_Combatep(combate + 1)
rondas_plantadormuere_errorh:
 
End Sub
 
 
 
Sub Rondas_plantadorDesconecta(ByVal Userindex As Integer)
On Error GoTo errorh
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(Userindex).Name & " ha desconectado en Torneo de Plantes, se le penaliza con 300.000 monedas de oro!", FontTypeNames.FONTTYPE_GUILD))
 If UserList(Userindex).Stats.GLD >= 300000 Then
UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - 300000
End If
Call Rondas_PlantadorMuere(Userindex, False, False)
errorh:
End Sub
 
 
 
Sub Rondas_UsuarioCambiamapa(ByVal Userindex As Integer)
On Error GoTo errorh
        Call Rondas_PlantadorMuere(Userindex, False, True)
errorh:
End Sub
 
Sub torneos_auto(ByVal rondas As Integer)
On Error GoTo errorh
If (Plantes_Activo) Then
               
                Exit Sub
        End If
       
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Torneo iniciado con un cupo de " & val(2 ^ rondas) & " participantes. Para participar teclea /PLANTES - (No cae inventario).", FontTypeNames.FONTTYPE_GUILD))
        Plantes_Rondas = rondas
        Plantes_Activo = True
        Plantes_Esperando = True
 
        ReDim Plantes_Luchadores(1 To 2 ^ rondas) As Integer
        Dim i As Integer
        For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                Plantes_Luchadores(i) = -1
        Next i
errorh:
End Sub
 
Sub plantes_Inicia(ByVal Userindex As Integer, ByVal rondas As Integer)
On Error GoTo errorh
 
        If (Plantes_Activo) Then
                Call WriteConsoleMsg(Userindex, "Ya hay un torneo en curso", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Esta empezando un nuevo torneo de plantes " & val(2 ^ rondas) & " participantes!! para participar teclea /PLANTES - (No cae inventario)", FontTypeNames.FONTTYPE_GUILD))
 
        Plantes_Rondas = rondas
        Plantes_Activo = True
        Plantes_Esperando = True
 
        ReDim Plantes_Luchadores(1 To 2 ^ rondas) As Integer
        Dim i As Integer
        For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                Plantes_Luchadores(i) = -1
        Next i
errorh:
End Sub
 
 
 
Sub plantes_Entra(ByVal Userindex As Integer)
On Error GoTo errorh
        Dim i As Integer
       
                   If Not UserList(Userindex).Pos.Map = 1 Then
    WriteConsoleMsg Userindex, "No puedes ingresar al torneo si no te encuentras en Ullathorpe.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If

            If UserList(Userindex).Stats.ELV < 42 Then
    WriteConsoleMsg Userindex, "No puedes ingresar al torneo si no eres superior a nivel 42.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
     End If
       
       
        If (Not Plantes_Activo) Then
                Call WriteConsoleMsg(Userindex, "¡No hay ningún torneo en curso!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
       
        If (Not Plantes_Esperando) Then
                Call WriteConsoleMsg(Userindex, "Cupos llenos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
       
        For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
                If (Plantes_Luchadores(i) = Userindex) Then
                        Call WriteConsoleMsg(Userindex, "¡Ya estás dentro del torneo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                End If
        Next i
 
        For i = LBound(Plantes_Luchadores) To UBound(Plantes_Luchadores)
        If (Plantes_Luchadores(i) = -1) Then
                Plantes_Luchadores(i) = Userindex
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapaespera
                    FuturePos.X = esperax: FuturePos.Y = esperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                   
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Plantes_Luchadores(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                 UserList(Plantes_Luchadores(i)).flags.Plantico = True
                 
                Call WriteConsoleMsg(Userindex, "¡Has ingresado al torneo!", FontTypeNames.FONTTYPE_INFO)
               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> El personaje " & UserList(Userindex).Name & " ingresó al torneo.", FontTypeNames.FONTTYPE_GUILD))
                If (i = UBound(Plantes_Luchadores)) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Se da inicio al torneo.", FontTypeNames.FONTTYPE_GUILD))
                Plantes_Esperando = False
                Call Rondas_Combatep(1)
     
                End If
                  Exit Sub
        End If
        Next i
errorh:
End Sub
 
 
Sub Rondas_Combatep(combate As Integer)
On Error GoTo errorh
Dim UI1 As Integer, UI2 As Integer
    UI1 = Plantes_Luchadores(2 * (combate - 1) + 1)
    UI2 = Plantes_Luchadores(2 * combate)
   
    If (UI2 = -1) Then
        UI2 = Plantes_Luchadores(2 * (combate - 1) + 1)
        UI1 = Plantes_Luchadores(2 * combate)
    End If
   
    If (UI1 = -1) Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Combate anulado por la desconexión de uno de los dos participantes.", FontTypeNames.FONTTYPE_GUILD))
        If (Plantes_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Torneo terminado, ganador del torneo por eliminación " & UserList(UI2).Name & ".", FontTypeNames.FONTTYPE_GUILD))
                UserList(UI2).flags.Plantico = False
                ' dale_recompensa()
                Plantes_Activo = False
                Exit Sub
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> No hay ganador del evento por la desconexión de todos sus participantes.", FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> El usuario " & UserList(UI2).Name & " pasó a la siguiente ronda.", FontTypeNames.FONTTYPE_GUILD))
           
        If (2 ^ Plantes_Rondas = 2 * combate) Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> Siguiente ronda.", FontTypeNames.FONTTYPE_GUILD))
            Plantes_Rondas = Plantes_Rondas - 1
            'antes de llamar a la proxima ronda hay q copiar a los tipos
            Dim i As Integer, j As Integer
            For i = 1 To 2 ^ Plantes_Rondas
                UI1 = Plantes_Luchadores(2 * (i - 1) + 1)
                UI2 = Plantes_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Plantes_Luchadores(i) = UI1
            Next i
            ReDim Preserve Plantes_Luchadores(1 To 2 ^ Plantes_Rondas) As Integer
            Call Rondas_Combatep(1)
            Exit Sub
        End If
        Call Rondas_Combatep(combate + 1)
        Exit Sub
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo de Plantes> " & UserList(UI1).Name & " vs. " & UserList(UI2).Name & ". ¡Preparados!... ¡A luchar!", FontTypeNames.FONTTYPE_GUILD))
    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
     Call SendData(SendTarget.toMap, 199, PrepareMessagePauseToggle())
                frmMain.ConteoPAutomatico.Enabled = True
errorh:
End Sub


