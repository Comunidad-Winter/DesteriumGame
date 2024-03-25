Attribute VB_Name = "Trabajo"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Const GASTO_ENERGIA_TRABAJADOR    As Byte = 2

Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo DoPermanecerOculto_Err

    '</EhHeader>

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo ErrHandler

    Dim TiempoTranscurrido As Long
    
    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        
        TiempoTranscurrido = (.Counters.TiempoOculto * frmMain.GameTimer.interval)
            
        If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
            Call WriteUpdateGlobalCounter(UserIndex, 1, .Counters.TiempoOculto / 40)

        End If
        
        If .Counters.TiempoOculto <= 0 Then
            If .Clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then

                ' Armaduras que permiten ocultarse por tiempo ilimitado
                If .Invent.ArmourEqpObjIndex > 0 Then
                    If ObjData(.Invent.ArmourEqpObjIndex).Oculto = 1 Then
                        .Counters.TiempoOculto = IntervaloOculto
                        Exit Sub

                    End If
                
                End If

            End If

            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 0 Then

                If .flags.Invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    
                    'Si está en el oscuro no lo hacemos visible
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.zonaOscura Then
                        Call SetInvisible(UserIndex, .Char.charindex, False)

                    End If

                End If

            End If

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoPermanecerOculto")

    '<EhFooter>
    Exit Sub

DoPermanecerOculto_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoPermanecerOculto " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la fórmula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo ErrHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
                    
        If .Clase = eClass.Thief Then
            Suerte = 80

        End If
            
        res = RandomNumber(1, 100)
        
        If .Stats.MaxMan > 0 Then Suerte = Suerte / 2
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .Clase = eClass.Thief Then
                Suerte = Suerte * 2
            Else

                If .Stats.MaxMan > 0 Then
                    Suerte = Suerte / 2

                End If

            End If
            
            .Counters.TiempoOculto = Suerte
             
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.charindex, True)
                
                .PosOculto = .Pos
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                Call WriteUpdateGlobalCounter(UserIndex, 1, .Counters.TiempoOculto / 40)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.Body = iFragataFantasmal

                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, Null)

            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
        Else
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, False)

        End If

        .Counters.Ocultando = .Counters.Ocultando + 1
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer, _
                    Optional ByVal NotRequired As Boolean = False)

    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    '13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
    '16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
    '10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
    '***************************************************
    '<EhHeader>
    On Error GoTo DoNavega_Err

    '</EhHeader>

    With UserList(UserIndex)

        If .Stats.Elv < 25 Then
            Call WriteConsoleMsg(UserIndex, "¡Las clases luchadoras pueden navegar a partir de Nivel 25!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
            
        If NotRequired = False Then
            If .Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    
                Exit Sub
    
            End If
            
            If HayAgua(.Pos.Map, .Pos.X, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X - 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X + 1, .Pos.Y) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y - 1) = True And HayAgua(.Pos.Map, .Pos.X, .Pos.Y + 1) = True Then
                Call WriteConsoleMsg(UserIndex, "¡¡No puedes dejar de navegar en el agua!!", FontTypeNames.FONTTYPE_INFO)
    
                Exit Sub
    
            End If

        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.BarcoSlot = Slot
            
            .Char.Head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
            
                Call ToggleBoatBody(UserIndex)
                
                ' Pierde el ocultar
                If .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    Call SetInvisible(UserIndex, .Char.charindex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If
               
                ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
                If .flags.Invisible = 1 Then
                    Call SetInvisible(UserIndex, .Char.charindex, False)
                    UserList(UserIndex).Counters.DrawersCount = 0

                End If
                
                ' Esta muerto
            Else
                .Char.Body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
            ' Estaba navegando
        Else
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
        
            ' No esta muerto
            If .flags.Muerto = 0 Then
                .Char.Head = .OrigChar.Head

                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
                Else
                    Call DarCuerpoDesnudo(UserIndex)

                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = GetShieldAnim(UserIndex, .Invent.EscudoEqpObjIndex)

                If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)

                If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = GetHelmAnim(UserIndex, .Invent.CascoEqpObjIndex)
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.Invisible = 1 Then
                    Call SetInvisible(UserIndex, .Char.charindex, True)
                    UserList(UserIndex).Counters.DrawersCount = RandomNumberPower(1, 200)

                End If
                
                ' Esta muerto
            Else
                .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
                .Char.Head = iCabezaMuerto(Escriminal(UserIndex))

                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco

            End If
            
            ' Termina de navegar
            .flags.Navegando = 0

        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

    End With
    
    Call WriteNavigateToggle(UserIndex)

    '<EhFooter>
    Exit Sub

DoNavega_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoNavega " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(UserIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) Then
                Call DoLingotes(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End If

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
    '***************************************************
    '<EhHeader>
    On Error GoTo TieneObjetos_Err

    '</EhHeader>

    Dim i     As Integer

    Dim Total As Long

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True

        Exit Function

    End If
        
    '<EhFooter>
    Exit Function

TieneObjetos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.TieneObjetos " & "at line " & Erl
        
    '</EhFooter>
End Function

'# Los objetos [BRONCE] [PLATA] [ORO] [PREMIUM] Se consideran especiales.
Function TieneObjetos_Especiales(ByVal UserIndex As Integer, _
                                 ByVal Bronce As Byte, _
                                 ByVal Plata As Byte, _
                                 ByVal Oro As Byte, _
                                 ByVal Premium As Byte) As String

    '<EhHeader>
    On Error GoTo TieneObjetos_Especiales_Err

    '</EhHeader>

    Dim A        As Integer

    Dim ObjIndex As Integer
    
    Dim Total    As Long

    For A = 1 To UserList(UserIndex).CurrentInventorySlots
        ObjIndex = UserList(UserIndex).Invent.Object(A).ObjIndex
        
        If ObjIndex > 0 Then
            If Bronce = 0 And ObjData(ObjIndex).Bronce = 1 Then
                TieneObjetos_Especiales = "El evento no permite los objetos [AVENTURERO]"
                Exit Function

            End If
            
            If Plata = 0 And ObjData(ObjIndex).Plata = 1 Then
                TieneObjetos_Especiales = "El evento no permite los objetos [HEROE]"
                Exit Function

            End If
            
            If Oro = 0 And ObjData(ObjIndex).Oro = 1 Then
                TieneObjetos_Especiales = "El evento no permite los objetos [LEYENDA]"
                Exit Function

            End If
            
            If Premium = 0 And ObjData(ObjIndex).Premium = 1 Then
                TieneObjetos_Especiales = "El evento no permite los objetos [PREMIUM]"
                Exit Function

            End If

        End If

    Next A
        
    '<EhFooter>
    Exit Function

TieneObjetos_Especiales_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.TieneObjetos_Especiales " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Long, _
                         ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarObjetos_Err

    '</EhHeader>

    Dim i As Integer

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        With UserList(UserIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
                .Amount = .Amount - cant
                
                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, UserIndex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

    '<EhFooter>
    Exit Sub

QuitarObjetos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.QuitarObjetos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub QuitarObjetoEspecifico(ByVal ItemIndex As Integer, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo QuitarObjetoEspecifico_Err

    '</EhHeader>
    Dim i As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        With UserList(UserIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Equipped = 1 Then Call Desequipar(UserIndex, i)

                UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                .Amount = 0
                .ObjIndex = 0
                
                Call UpdateUserInv(False, UserIndex, i)
                
            End If

        End With

    Next i

    '<EhFooter>
    Exit Sub

QuitarObjetoEspecifico_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.QuitarObjetoEspecifico " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '<EhHeader>
    On Error GoTo MineralesParaLingote_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 13

        Case iMinerales.PlataCruda
            MineralesParaLingote = 25

        Case iMinerales.OroCrudo
            MineralesParaLingote = 50

        Case Else
            MineralesParaLingote = 10000

    End Select

    '<EhFooter>
    Exit Function

MineralesParaLingote_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.MineralesParaLingote " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo DoLingotes_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim Slot           As Integer

    Dim obji           As Integer

    Dim CantidadItems  As Integer

    Dim TieneMinerales As Boolean

    Dim OtroUserIndex  As Integer
    
    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)

            End If

        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.Elv - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0

        End If
        
        Dim MiObj As Obj

        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    '<EhFooter>
    Exit Sub

DoLingotes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoLingotes " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo DoFundir_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2010
    '03/06/2010 - Pato: Si es el último ítem a fundir y está equipado lo desequipamos.
    '11/03/2010 - ZaMa: Reemplazo división por producto para uan mejor performanse.
    '***************************************************
    Dim i             As Integer

    Dim Num           As Integer

    Dim Slot          As Byte

    Dim Lingotes(2)   As Integer

    Dim OtroUserIndex As Integer

    With UserList(UserIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)

            End If

        End If
        
        Slot = .flags.TargetObjInvSlot
        
        With .Invent.Object(Slot)
            .Amount = .Amount - 1
            
            If .Amount < 1 Then
                If .Equipped = 1 Then Call Desequipar(UserIndex, Slot)
                
                .Amount = 0
                .ObjIndex = 0

            End If

        End With
        
        Num = RandomNumber(10, 25)
        
        Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
        Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
        Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
    
        Dim MiObj(2) As Obj
        
        For i = 0 To 2
            MiObj(i).Amount = Lingotes(i)
            MiObj(i).ObjIndex = LingoteHierro + i 'Una gran negrada pero práctica
            
            If MiObj(i).Amount > 0 Then
                If Not MeterItemEnInventario(UserIndex, MiObj(i)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(i))

                End If

            End If

        Next i
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido el " & Num & "% de los lingotes utilizados para la construcción del objeto!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    '<EhFooter>
    Exit Sub

DoFundir_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoFundir " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub DoAdminInvisible(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    '<EhHeader>
    On Error GoTo DoAdminInvisible_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .flags.AdminInvisible = 0 Then

            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False

            End If
            
            .flags.AdminInvisible = 1
            .flags.Invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.Body
            .flags.OldHead = .Char.Head
            .Char.Body = 0
            .Char.Head = 0
                
            ' Solo el admin sabe que se hace invi
            Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.charindex))
            ' Call ModAreas.DeleteEntity(UserIndex, ENTITY_TYPE_PLAYER)
        
        Else
            .flags.AdminInvisible = 0
            .flags.Invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.Body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            ' Solo el admin sabe que se hace visible
            Call SendData(ToOne, UserIndex, PrepareMessageCharacterChange(.Char.Body, 0, .Char.Head, .Char.Heading, .Char.charindex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim, .Char.AuraIndex, .flags.ModoStream, False, False))
            Call SendData(ToOne, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
            
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, True, True)
            
            ' Se lo mando a los demas
            Call ModAreas.CreateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos, ModAreas.DEFAULT_ENTITY_WIDTH, ModAreas.DEFAULT_ENTITY_HEIGHT)
        
        End If

    End With
    
    '<EhFooter>
    Exit Sub

DoAdminInvisible_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoAdminInvisible " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 28/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim CantidadItems As Integer

    With UserList(UserIndex)

        ' Si estaba oculto, se vuelve visible
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
                
            If .flags.Invisible = 0 Then
                Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                Call SetInvisible(UserIndex, .Char.charindex, False)

            End If

        End If
            
        Call QuitarSta(UserIndex, RandomNumber(0, EsfuerzoTalarLeñador))
    
        Dim Skill As Integer

        Skill = .Stats.UserSkills(eSkill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= 5 Then

            Dim MiObj As Obj
        
            If .flags.TargetObj = 0 Then Exit Sub
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
            CantidadItems = MaxItemsExtraibles(.Stats.Elv)
            
            MiObj.Amount = RandomNumber(1, CantidadItems + ObjData(.Invent.WeaponEqpObjIndex).ProbPesca)
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
        
            Call SubirSkill(UserIndex, eSkill.Mineria, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(UserIndex, eSkill.Mineria, False)

        End If
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_MINERO, .Pos.X, .Pos.Y))
        
        If Not Escriminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, ByVal WeaponIndex As Integer)

    '<EhHeader>
    On Error GoTo DoPescar_Err

    '</EhHeader>

    Dim iSkill        As Integer

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim CantidadItems As Integer
    
    Dim LastFish      As Byte    ' Ultimo pescado disponible

    Dim MaxSuerte     As Byte   ' Mejora la cantidad de suerte segun el barco que tenga
     
    With UserList(UserIndex)
            
        If MapInfo(.Pos.Map).Pesca = 0 Then Exit Sub ' # No hay peces en el mapa
            
        Select Case WeaponIndex

            Case CAÑA_PESCA

                If .Invent.BarcoObjIndex > 0 Then
                    MaxSuerte = ObjData(.Invent.BarcoObjIndex).ProbPesca

                End If
                      
            Case RED_PESCA

                If .Invent.BarcoObjIndex <> 475 And .Invent.BarcoObjIndex <> 476 Then Exit Sub

                If Abs(.Pos.X - .flags.TargetX) + Abs(.Pos.Y - .flags.TargetY) > 5 Then
                    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                            
                If .Pos.X = .flags.TargetX And .Pos.Y = .flags.TargetY Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde allí.", FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If
                
                MaxSuerte = ObjData(.Invent.BarcoObjIndex).ProbPesca

            Case Else
                Exit Sub

        End Select
                            
        'Play sound!
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_PESCAR, .Pos.X, .Pos.Y, .Char.charindex))
                              
        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
        
        iSkill = .Stats.UserSkills(eSkill.Pesca)
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 55)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= (6 + MaxSuerte) Then
            
                Dim MiObj   As Obj

                Dim A       As Long

                Dim N       As Long
                
                Dim SacoPez As Boolean
                    
                Dim Slot    As Integer
                    
                Slot = RandomNumber(1, MapInfo(.Pos.Map).Pesca)

                MiObj.ObjIndex = MapInfo(.Pos.Map).PescaItem(Slot)
                MiObj.Amount = RandomNumber(1, ObjData(WeaponIndex).ProbPesca)
                        
                If RandomNumber(1, 100) <= ObjData(MiObj.ObjIndex).ProbPesca Then

                    If Not MeterItemEnInventario(UserIndex, MiObj) Then
                        Call TirarItemAlPiso(.Pos, MiObj)

                    End If
                            
                    Call SubirSkill(UserIndex, eSkill.Pesca, True)

                End If

            Else
                
                Call SubirSkill(UserIndex, eSkill.Pesca, False)

            End If

        End If
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    '<EhFooter>
    Exit Sub

DoPescar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoPescar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************
    '<EhHeader>
    On Error GoTo DoRobar_Err

    '</EhHeader>

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    ' Caos robando a caos?
    If UserList(LadrOnIndex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(LadrOnIndex, "¡No puedes robar o hurtar objetos si no te encuentras oculto!", FontTypeNames.FONTTYPE_FIGHT)

        Exit Sub

    End If
        
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)

        Exit Sub

    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not Escriminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)

                Exit Sub

            End If

        Else

            If .Faction.Status = r_Armada Then
                If Not Escriminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejército real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)

                    Exit Sub

                End If

            End If

        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faction.Status = r_Caos And .Faction.Status = r_Caos Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legión oscura.", FontTypeNames.FONTTYPE_FIGHT)

            Exit Sub

        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
            
        ' ¿La victima tiene energia para ser robado?
        If UserList(VictimaIndex).Stats.MinSta < 15 Then
            Call WriteConsoleMsg(LadrOnIndex, "El Personaje está muy cansado para poder defenderse del hurto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.WeaponEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer

            Dim res        As Integer

            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)
                        Call Protocol.FlushBuffer(OtroUserIndex)

                    End If

                End If
               
                If (RandomNumber(1, 100) < 35) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.Gld > 0 Then

                        Dim N As Long
                        
                        If .Clase = eClass.Thief Then

                            ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                N = RandomNumber((.Stats.Elv) * 50, (.Stats.Elv) * 100)
                            Else
                                N = RandomNumber(.Stats.Elv * 25, (.Stats.Elv) * 50)

                            End If
                            
                            If UserList(VictimaIndex).flags.Paralizado = 1 Or UserList(VictimaIndex).flags.Inmovilizado = 1 Then
                                N = N * 1.3

                            End If
                            
                        Else
                            N = RandomNumber(1, 100)

                        End If

                        If N > UserList(VictimaIndex).Stats.Gld Then
                            N = UserList(VictimaIndex).Stats.Gld
                            Call WriteConsoleMsg(LadrOnIndex, "¡Le has robado todo el Oro a " & UserList(VictimaIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                            Call WriteConsoleMsg(VictimaIndex, "¡Te han robado todo el Oro!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                            Call WriteConsoleMsg(VictimaIndex, "Te han robado " & N & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - N
                        
                        .Stats.Gld = .Stats.Gld + N

                        If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                      
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Call FlushBuffer(VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(VictimaIndex)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

            End If
        
            If Not Escriminal(LadrOnIndex) Then
                If Not Escriminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)

                End If

            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If Escriminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

                If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

DoRobar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoRobar " & "at line " & Erl
        
    '</EhFooter>
End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregué los barcos
    ' Esta funcion determina qué objetos son robables.
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************
    '<EhHeader>
    On Error GoTo ObjEsRobable_Err

    '</EhHeader>

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otMonturas And ObjData(OI).Bronce <> 1 And ObjData(OI).Premium <> 1 And ObjData(OI).Plata <> 1 And ObjData(OI).Oro <> 1 And Not ItemNewbie(OI) And ObjData(OI).NoNada <> 1 And Not ObjData(OI).OBJType = otGemaTelep And Not ObjData(OI).OBJType = otTransformVIP

    '<EhFooter>
    Exit Function

ObjEsRobable_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.ObjEsRobable " & "at line " & Erl
        
    '</EhFooter>
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************
    '<EhHeader>
    On Error GoTo RobarObjeto_Err

    '</EhHeader>

    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    With UserList(VictimaIndex)

        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
            i = 1

            Do While Not flag And i <= .CurrentInventorySlots

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i + 1
            Loop

        Else
            i = .CurrentInventorySlots

            Do While Not flag And i > 0

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i - 1
            Loop

        End If
    
        If flag Then

            Dim MiObj     As Obj

            Dim Num       As Integer

            Dim ObjAmount As Integer
        
            ObjAmount = .Invent.Object(i).Amount

            If UserList(VictimaIndex).flags.Paralizado = 1 Or UserList(VictimaIndex).flags.Inmovilizado = 1 Then
                'Cantidad al azar entre el 15% y el 20% del total, con minimo 1.
                Num = MaximoInt(1, RandomNumber(ObjAmount * 0.15, ObjAmount * 0.2))
            Else
                'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
                Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))

            End If

            MiObj.Amount = Num
            MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
            .Invent.Object(i).Amount = ObjAmount - Num
                    
            If .Invent.Object(i).Amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

            End If
                
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
            If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
        
            If UserList(LadrOnIndex).Clase = eClass.Thief Then
                Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "Te han robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)

    End With

    '<EhFooter>
    Exit Sub

RobarObjeto_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.RobarObjeto " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long, ByRef exito As Boolean)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el daño
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Suerte   As Integer

    Dim Skill    As Integer

    Dim ObjIndex As Integer
    
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

    Select Case UserList(UserIndex).Clase

        Case eClass.Assasin
            Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.078) * Skill + 4.45)
          
        Case eClass.Cleric, eClass.Paladin
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
          
        Case eClass.Bard, eClass.Hunter, eClass.Warrior
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
          
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    ' Si no es Daga Arrojadiza
    
    Dim WeaponIndex As Integer

    Dim CanDaga     As Boolean
    
    WeaponIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
    CanDaga = True
    
    If WeaponIndex > 0 Then
        If ObjData(WeaponIndex).proyectil = 1 And ObjData(WeaponIndex).Apuñala = 1 Then
            CanDaga = False

        End If

    End If
    
    If CanDaga Then
        If VictimUserIndex Then
            If UserList(UserIndex).Clase = eClass.Assasin Then
                If UserList(UserIndex).Char.Heading = UserList(VictimUserIndex).Char.Heading Then Suerte = 60

            End If

        End If

    End If
    
    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            daño = Round(daño * 1.5, 0)

            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
            
            SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, UserList(UserIndex).DañoApu + daño, eDamageType.d_Apuñalar)
            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessagePlayEffect(eSound.sApuñaladaEspalda, UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y))
            
            Call WriteConsoleMsg(VictimUserIndex, "¡Te han dado una apuñalada por " & Int(UserList(UserIndex).DañoApu + daño) & "!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(UserIndex, "¡Apuñalada por " & Int(UserList(UserIndex).DañoApu + daño) & "!", FontTypeNames.FONTTYPE_FIGHT)
            
            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateFX(UserList(VictimUserIndex).Char.charindex, FXIDs.FX_APUÑALADA, 1))
            
            exito = True
        Else
            ObjIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
            
            daño = daño + ObjData(ObjIndex).NpcBonusDamage
            
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            SendData SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y, Int(UserList(UserIndex).DañoApu + daño), eDamageType.d_Apuñalar)
            Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessagePlayEffect(eSound.sApuñaladaEspalda, Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y))
            Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(UserList(UserIndex).DañoApu + daño), FontTypeNames.FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateFX(Npclist(VictimNpcIndex).Char.charindex, FXIDs.FX_APUÑALADA, 1))

            '[Alejo]
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
            
            exito = True

        End If
          
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
    Else
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
        '  Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        'SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, DAMAGE_NORMAL)
        Call SubirSkill(UserIndex, eSkill.Apuñalar, True)

    End If

    Exit Sub
ErrHandler:

End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 12/01/2010
    '***************************************************
    '<EhHeader>
    On Error GoTo DoAcuchillar_Err

    '</EhHeader>

    If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
        daño = Int(daño * DAÑO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - daño
                Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & .Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & daño, FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
        
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
        
        End If

    End If
    
    '<EhFooter>
    Exit Sub

DoAcuchillar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoAcuchillar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoGolpeCritico_Npcs(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal daño As Long)

    '***************************************************
    'Autor: Lautaro
    'Last Modification:
    ' Lo hacemos aparte porque queremos dejar el otro
    '***************************************************
    '<EhHeader>
    On Error GoTo DoGolpeCritico_Npcs_Err

    '</EhHeader>
   
    daño = UserList(UserIndex).Stats.Elv * 2
        
    With UserList(UserIndex)
        
        If .Clase <> eClass.Warrior And .Clase <> eClass.Hunter And .Clase <> eClass.Thief Then Exit Sub
    
    End With
    
    With Npclist(NpcIndex)
    
        .Stats.MinHp = .Stats.MinHp - daño
        
        ' Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_CRITICO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y - 1, daño, eDamageType.d_DañoNpc_Critical))
        Call CalcularDarExp(UserIndex, NpcIndex, daño)

    End With
    
    '<EhFooter>
    Exit Sub

DoGolpeCritico_Npcs_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoGolpeCritico_Npcs " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)

    '<EhHeader>
    On Error GoTo DoGolpeCritico_Err

    '</EhHeader>

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
    '***************************************************
    Dim Suerte      As Integer

    Dim Skill       As Integer

    Dim WeaponIndex As Integer
    
    Exit Sub
    
    With UserList(UserIndex)
        ' Es bandido?
        'If .Clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Armas)

    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        daño = Int(daño * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - daño
                Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & .Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
        
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
            
        End If
        
    End If

    '<EhFooter>
    Exit Sub

DoGolpeCritico_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoGolpeCritico " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal cantidad As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarSta_Err

    '</EhHeader>

    On Error GoTo ErrHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarSta. Error " & Err.number & " : " & Err.description)
    
    '<EhFooter>
    Exit Sub

QuitarSta_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.QuitarSta " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 28/05/2010
    '16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo ErrHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim CantidadItems As Integer

    Dim Skill         As Integer

    With UserList(UserIndex)

        Call QuitarSta(UserIndex, RandomNumber(0, EsfuerzoTalarLeñador))
    
        Skill = .Stats.UserSkills(eSkill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= 4 Then

            Dim MiObj As Obj

            CantidadItems = MaxItemsExtraibles(.Stats.Elv)
            
            MiObj.Amount = RandomNumber(1, CantidadItems)
            MiObj.ObjIndex = ObjIndex
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            Call SubirSkill(UserIndex, eSkill.Talar, True)
        Else

            '[/CDT]
            Call SubirSkill(UserIndex, eSkill.Talar, False)

        End If
    
        If Not Escriminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo DoMeditar_Err

    '</EhHeader>

    Dim Mana As Long
        
    With UserList(UserIndex)

        .Counters.TimerMeditar = .Counters.TimerMeditar + 1
        .Counters.TiempoInicioMeditar = .Counters.TiempoInicioMeditar + 1
            
        If .Counters.TimerMeditar >= IntervaloMeditar Then

            Mana = Porcentaje(.Stats.MaxMan, Porcentaje(Balance.PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Magia) * 0.5))

            If Mana <= 0 Then Mana = 1

            If .Stats.MinMan + Mana >= .Stats.MaxMan Then

                .Stats.MinMan = .Stats.MaxMan
                .flags.Meditando = False
                .Char.FX = 0

                Call WriteUpdateMana(UserIndex)

                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
                
            Else
                    
                .Stats.MinMan = .Stats.MinMan + Mana
                Call WriteUpdateMana(UserIndex)

            End If

            .Counters.TimerMeditar = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

DoMeditar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoMeditar " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************
    '<EhHeader>
    On Error GoTo DoDesequipar_Err

    '</EhHeader>

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte

    Dim AlgoEquipado   As Boolean
    
    With UserList(UserIndex)

        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Armas)
        
        Probabilidad = WrestlingSkill * 0.2 + (.Stats.Elv) * 0.66

    End With
   
    With UserList(VictimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

    '<EhFooter>
    Exit Sub

DoDesequipar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoDesequipar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '<EhHeader>
    On Error GoTo DoHurtar_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
    '***************************************************
    Dim OtroUserIndex As Integer

    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    Exit Sub

    'If UserList(UserIndex).Clase <> eClass.Bandit Then Exit Sub
    'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
    'Uso el slot de los anillos para "equipar" los guantes.
    'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
    If UserList(UserIndex).Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub

    Dim res As Integer

    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
    
            If UserList(VictimaIndex).flags.Comerciando Then
                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                    Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                    Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                    Call LimpiarComercioSeguro(VictimaIndex)
                    Call Protocol.FlushBuffer(OtroUserIndex)

                End If

            End If
                
            Call RobarObjeto(UserIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

    '<EhFooter>
    Exit Sub

DoHurtar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoHurtar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)

    '<EhHeader>
    On Error GoTo DoHandInmo_Err

    '</EhHeader>

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 17/02/2007
    'Implements the special Skill of the Thief
    '***************************************************
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    If UserList(UserIndex).Clase <> eClass.Thief Then Exit Sub
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
    Dim res As Integer

    res = RandomNumber(0, 100)

    If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) / 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
        UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)

    End If

    '<EhFooter>
    Exit Sub

DoHandInmo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.DoHandInmo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************
    '<EhHeader>
    On Error GoTo Desarmar_Err

    '</EhHeader>

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Armas)
        
        Probabilidad = WrestlingSkill * 0.2 + (.Stats.Elv) * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            Call FlushBuffer(VictimIndex)

        End If

    End With
    
    '<EhFooter>
    Exit Sub

Desarmar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.Desarmar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '***************************************************
    '<EhHeader>
    On Error GoTo MaxItemsConstruibles_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        MaxItemsConstruibles = MaximoInt(1, CInt(((.Stats.Elv) - 2) * 0.2))

    End With

    '<EhFooter>
    Exit Function

MaxItemsConstruibles_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.MaxItemsConstruibles " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/05/2010
    '***************************************************
    '<EhHeader>
    On Error GoTo MaxItemsExtraibles_Err

    '</EhHeader>
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1
    '<EhFooter>
    Exit Function

MaxItemsExtraibles_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.MaxItemsExtraibles " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Copies body, head and desc from previously clicked npc.
    '***************************************************
    '<EhHeader>
    On Error GoTo ImitateNpc_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        ' Copy desc
        .DescRM = Npclist(NpcIndex).Name
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.Invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NpcIndex).Char.Body
            .flags.OldHead = Npclist(NpcIndex).Char.Head
        Else
            .Char.Body = Npclist(NpcIndex).Char.Body
            .Char.Head = Npclist(NpcIndex).Char.Head
            
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)

        End If
    
    End With
    
    '<EhFooter>
    Exit Sub

ImitateNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Trabajo.ImitateNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

