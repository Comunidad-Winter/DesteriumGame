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

Private Const GASTO_ENERGIA_TRABAJADOR As Byte = 2
Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6

Public Sub DoPermanecerOculto(ByVal Userindex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 11/19/2009
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
'********************************************************
On Error GoTo Errhandler
    With UserList(Userindex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        If .Counters.TiempoOculto <= 0 Then
            If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 612 Or .Invent.ArmourEqpObjIndex = 360 Or .Invent.ArmourEqpObjIndex = 671 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub
                End If
            End If
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(Userindex)
                    Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
'13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
'***************************************************

On Error GoTo Errhandler

    Dim Suerte As Double
    Dim res As Integer
    Dim Skill As Integer
    
    With UserList(Userindex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            .Counters.TiempoOculto = Suerte
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(Userindex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(Userindex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
            ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                    NingunEscudo, NingunCasco)
            End If
            
            Call SubirSkill(Userindex, eSkill.Ocultarse, True)
        Else
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(Userindex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
            '[/CDT]
            
            Call SubirSkill(Userindex, eSkill.Ocultarse, False)
        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal Userindex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
'***************************************************

    Dim ModNave As Single
    
    With UserList(Userindex)
        ModNave = ModNavegacion(.clase, Userindex)
If HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) = True And HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y) = True And HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y) = True And _
HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1) = True And HayAgua(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1) = True Then
    Call WriteConsoleMsg(Userindex, "No puedes dejar de navegar en el agua!!", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
                                         If .flags.Montando = 1 Then
Call WriteConsoleMsg(Userindex, "¡No puedes navegar si estás montando!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
End If

                                 If .flags.Mimetizado = 1 Then
Call WriteConsoleMsg(Userindex, "¡No puedes mimetizarte si estás navegando!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
End If
        
    UserList(Userindex).Invent.BarcoObjIndex = UserList(Userindex).Invent.Object(Slot).objindex
UserList(Userindex).Invent.BarcoSlot = Slot

If UserList(Userindex).flags.Navegando = 0 Then
End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).objindex
            .Invent.BarcoSlot = Slot
            
            .Char.Head = 0
            
               ' No esta muerto
            If .flags.Muerto = 0 Then
            
                Call ToggleBoatBody(Userindex)
                
                ' Pierde el ocultar
                If .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
               
                ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
                If .flags.invisible = 1 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)
                End If
                
            ' Esta muerto
            Else
                .Char.body = iFragataFantasmal
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
                
                If .clase = eClass.Pirat Then
                    If .flags.Oculto = 1 Then
                        ' Al desequipar barca, perdió el ocultar
                        .flags.Oculto = 0
                        .Counters.Ocultando = 0
                        Call WriteConsoleMsg(Userindex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(Userindex)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then _
                    .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then _
                    .Char.WeaponAnim = GetWeaponAnim(Userindex, .Invent.WeaponEqpObjIndex)
                If .Invent.CascoEqpObjIndex > 0 Then _
                    .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.invisible = 1 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, True)
                End If
                
            ' Esta muerto
            Else
                .Char.body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            
            ' Termina de navegar
            .flags.Navegando = 0
        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
    
    Call WriteNavigateToggle(Userindex)

End Sub

Public Sub FundirMineral(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    With UserList(Userindex)
        If .flags.TargetObjInvIndex > 0 Then
           
           If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) / ModFundicion(.clase) Then
                Call DoLingotes(Userindex)
           Else
                Call WriteConsoleMsg(Userindex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
           End If
        
        End If
    End With

    Exit Sub

Errhandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub


Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    Dim Total As Long
    For i = 1 To UserList(Userindex).CurrentInventorySlots
        If UserList(Userindex).Invent.Object(i).objindex = ItemIndex Then
            Total = Total + UserList(Userindex).Invent.Object(i).Amount
        End If
    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function
    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************

    Dim i As Integer
    For i = 1 To UserList(Userindex).CurrentInventorySlots
        With UserList(Userindex).Invent.Object(i)
            If .objindex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(Userindex, i)
                
                .Amount = .Amount - cant
                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
                    .Amount = 0
                    .objindex = 0
                Else
                    cant = 0
                End If
                
                Call UpdateUserInv(False, Userindex, i)
                
                If cant = 0 Then Exit Sub
            End If
        End With
    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, Userindex)
End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Function TieneMaterialesUpgrade(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'
'***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .Madera > 0 Then
            If Not TieneObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), Userindex) Then
                Call WriteConsoleMsg(Userindex, "No tienes suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
    End With
    
    TieneMaterialesUpgrade = True
End Function

Sub QuitarMaterialesUpgrade(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'
'***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .Madera > 0 Then Call QuitarObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), Userindex)
    End With
    
    Call QuitarObjetos(ItemIndex, 1, Userindex)
End Sub

Public Function PuedeConstruir(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex) And UserList(Userindex).Stats.UserSkills(eSkill.herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function

Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

If PuedeConstruir(Userindex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(Userindex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(Userindex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otescudo Then
        Call WriteConsoleMsg(Userindex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otcasco Then
        Call WriteConsoleMsg(Userindex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otarmadura Then
        Call WriteConsoleMsg(Userindex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.objindex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.objindex).LOG = 1 Then
        Call LogDesarrollo(UserList(Userindex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.objindex).Name)
    End If
    
    'Call SubirSkill(UserIndex, herreria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)

If CarpinteroTieneMateriales(Userindex, ItemIndex) And _
   UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(Userindex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
    
    Call CarpinteroQuitarMateriales(Userindex, ItemIndex)
    Call WriteConsoleMsg(Userindex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.objindex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.objindex).LOG = 1 Then
        Call LogDesarrollo(UserList(Userindex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.objindex).Name)
    End If
    
    'Call SubirSkill(UserIndex, Carpinteria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))


    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

End If
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 25
        Case iMinerales.PlataCruda
            MineralesParaLingote = 35
        Case iMinerales.OroCrudo
            MineralesParaLingote = 50
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer
    Dim CantidadItems As Integer
    Dim TieneMinerales As Boolean
    Dim OtroUserIndex As Integer
    
    With UserList(Userindex)
        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(Userindex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(Userindex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).objindex
        
        While CantidadItems > 0 And Not TieneMinerales
            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems
        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).objindex = 0
        End If
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.objindex = ObjData(.flags.TargetObjInvIndex).LingoteIndex
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, Userindex, Slot)
        Call WriteConsoleMsg(Userindex, "¡Has obtenido " & CantidadItems & " lingote" & _
                            IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
End Sub

Public Sub DoUpgrade(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'12/08/2009: Pato - Implementado nuevo sistema de mejora de items
'***************************************************
Dim ItemUpgrade As Integer
Dim WeaponIndex As Integer
Dim OtroUserIndex As Integer

ItemUpgrade = ObjData(ItemIndex).Upgrade

With UserList(Userindex)
    If .flags.Comerciando Then
        OtroUserIndex = .ComUsu.DestUsu
            
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(Userindex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
            Call LimpiarComercioSeguro(Userindex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
        
    'Sacamos energía
    If .clase = eClass.Worker Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If ItemUpgrade <= 0 Then Exit Sub
    If Not TieneMaterialesUpgrade(Userindex, ItemIndex) Then Exit Sub
    
    If PuedeConstruirHerreria(ItemUpgrade) Then
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> MARTILLO_HERRERO Then
            Call WriteConsoleMsg(Userindex, "Debes equiparte el martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Round(.Stats.UserSkills(eSkill.herreria) / ModHerreriA(.clase), 0) < ObjData(ItemUpgrade).SkHerreria Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(Userindex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otescudo 'Todavía no hay, pero just in case
                Call WriteConsoleMsg(Userindex, "Has mejorado el escudo!", FontTypeNames.FONTTYPE_INFO)
            
            Case eOBJType.otcasco
                Call WriteConsoleMsg(Userindex, "Has mejorado el casco!", FontTypeNames.FONTTYPE_INFO)
            
            Case eOBJType.otarmadura
                Call WriteConsoleMsg(Userindex, "Has mejorado la armadura!", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Call SubirSkill(Userindex, eSkill.herreria, True)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.X, .Pos.Y))
    
    ElseIf PuedeConstruirCarpintero(ItemUpgrade) Then
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(Userindex, "Debes equiparte un serrucho.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.clase), 0) < ObjData(ItemUpgrade).SkCarpinteria Then
            Call WriteConsoleMsg(Userindex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otFlechas
                Call WriteConsoleMsg(Userindex, "Has mejorado la flecha!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(Userindex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otBarcos
                Call WriteConsoleMsg(Userindex, "Has mejorado el barco!", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Call SubirSkill(Userindex, eSkill.Carpinteria, True)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.X, .Pos.Y))
    Else
        Exit Sub
    End If
    
    Call QuitarMaterialesUpgrade(Userindex, ItemIndex)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.objindex = ItemUpgrade
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(.Pos, MiObj)
    End If
    
    If ObjData(ItemIndex).LOG = 1 Then _
        Call LogDesarrollo(.Name & " ha mejorado el ítem " & ObjData(ItemIndex).Name & " a " & ObjData(ItemUpgrade).Name)
        
    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta
    If .Reputacion.PlebeRep > MAXREP Then _
        .Reputacion.PlebeRep = MAXREP
        
    .Counters.Trabajando = .Counters.Trabajando + 1
End With
End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal Userindex As Integer) As Single
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 27/11/2009
'27/11/2009: ZaMa - A worker can navigate before only if it's an expert fisher
'12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
'***************************************************
Select Case clase
    Case eClass.Pirat
        ModNavegacion = 1
    Case eClass.Worker
        If UserList(Userindex).Stats.UserSkills(eSkill.Pesca) = 100 Then
            ModNavegacion = 1.71
        Else
            ModNavegacion = 2
        End If
    Case Else
        ModNavegacion = 2
End Select

End Function


Function ModFundicion(ByVal clase As eClass) As Single
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Select Case clase
    Case eClass.Worker
        ModFundicion = 1
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Select Case clase
    Case eClass.Worker
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
Select Case clase
    Case eClass.Worker
        ModHerreriA = 1
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Select Case clase
        Case eClass.Druid
            ModDomar = 6
        Case eClass.Hunter
            ModDomar = 6
        Case eClass.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 02/03/2009
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'***************************************************

On Error GoTo Errhandler

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer
Dim CanStay As Boolean
Dim petType As Integer
Dim NroPets As Integer


If Npclist(NpcIndex).MaestroUser = Userindex Then
    Call WriteConsoleMsg(Userindex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(Userindex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
  '  If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
   '     Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
    '    Exit Sub
 '   End If
    
    puntosDomar = CInt(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(Userindex).Stats.UserSkills(eSkill.Domar))
    If UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAELFICA And UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloPlata Then
        puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
    Else
        puntosRequeridos = Npclist(NpcIndex).flags.Domable
    End If
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(Userindex).NroMascotas = UserList(Userindex).NroMascotas + 1
        index = FreeMascotaIndex(Userindex)
        UserList(Userindex).MascotasIndex(index) = NpcIndex
        UserList(Userindex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = Userindex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(Userindex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (MapInfo(UserList(Userindex).Pos.Map).Pk = True)
        
        If Not CanStay Then
            petType = Npclist(NpcIndex).Numero
            NroPets = UserList(Userindex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
           UserList(Userindex).MascotasType(index) = petType
            UserList(Userindex).NroMascotas = NroPets
            
          '  Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(Userindex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(Userindex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.UltimoMensaje = 5
        End If
    End If
    
    'Entreno domar. Es un 30% más dificil si no sos druida.
    If UserList(Userindex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
        Call SubirSkill(Userindex, Domar, True)
    End If
Else
    Call WriteConsoleMsg(Userindex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub

Errhandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub
''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'***************************************************
    
    With UserList(Userindex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
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
            .flags.invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            .Char.body = 0
            .Char.Head = 0
            
            ' Solo el admin sabe que se hace invi
            Call EnviarDatosASlot(Userindex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            ' Solo el admin sabe que se hace visible
            Call EnviarDatosASlot(Userindex, PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.Heading, _
            .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim))
            Call EnviarDatosASlot(Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
             
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.Map, Userindex, .Pos.Map, .Pos.X, .Pos.Y, True)
        End If
    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If MapData(Map, X, Y).ObjInfo.objindex <> 58 Then
    Call WriteConsoleMsg(Userindex, "Necesitas clickear sobre leña para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(Userindex).Pos) > 2 Then
    Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(Userindex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim SupervivenciaSkill As Byte

SupervivenciaSkill = UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia)

If SupervivenciaSkill >= 0 And SupervivenciaSkill < 6 Then
    Suerte = 3
ElseIf SupervivenciaSkill >= 6 And SupervivenciaSkill <= 34 Then
    Suerte = 2
ElseIf SupervivenciaSkill >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.objindex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
    Call WriteConsoleMsg(Userindex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, Map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(Userindex).flags.TargetObj = FOGATA_APAG
    
    Call SubirSkill(Userindex, eSkill.Supervivencia, True)
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(Userindex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
    
    Call SubirSkill(Userindex, eSkill.Supervivencia, False)
End If

End Sub

Public Sub DoPescar(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'***************************************************
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

If UserList(Userindex).clase = eClass.Worker Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).clase = eClass.Worker Then
        With UserList(Userindex)
            CantidadItems = 1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))
        End With
        
        MiObj.Amount = RandomNumber(1, CantidadItems)
    Else
        MiObj.Amount = 1
    End If
    MiObj.objindex = Pescado
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(Userindex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
    Call SubirSkill(Userindex, eSkill.Pesca, True)
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(Userindex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(Userindex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
    
    Call SubirSkill(Userindex, eSkill.Pesca, False)
End If

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.description)
End Sub

Public Sub DoPescarRed(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo Errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UserList(Userindex).clase = eClass.Worker Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(Userindex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1
        End If
        MiObj.objindex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(Userindex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
        Call SubirSkill(Userindex, eSkill.Pesca, True)
    Else
        Call WriteConsoleMsg(Userindex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(Userindex, eSkill.Pesca, False)
    End If
End If

    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP
        
Exit Sub

Errhandler:
    Call LogError("Error en DoPescarRed")
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

On Error GoTo Errhandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        Else
            If .Faccion.ArmadaReal = 1 Then
                If Not criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejército real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
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
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte As Integer
            Dim res As Integer
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
               
                If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else 'Roba oro
                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim N As Long
                        
                        If .clase = eClass.Thief Then
                        ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                N = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                N = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)
                            End If
                        Else
                            N = RandomNumber(1, 100)
                        End If
                        If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                        
                        .Stats.GLD = .Stats.GLD + N
                        If .Stats.GLD > MaxOro Then _
                            .Stats.GLD = MaxOro
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
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
        
            If Not criminal(LadrOnIndex) Then
                If Not criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)
                End If
            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron
                If .Reputacion.LadronesRep > MAXREP Then _
                    .Reputacion.LadronesRep = MAXREP
            End If
        End If
    End With

Exit Sub

Errhandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
' Agregué los barcos
' Esta funcion determina qué objetos son robables.
'***************************************************

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).objindex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otMonturas And _
ObjData(OI).OBJType <> eOBJType.otMonturasDraco And _
ObjData(OI).VIP = 0 And _
ObjData(OI).VIPP = 0 And _
ObjData(OI).VIPB = 0 And _
ObjData(OI).UM = 0 And _
ObjData(OI).HM = 0 And _
ObjData(OI).NoSeCae = 0 And _
ObjData(OI).Newbie = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

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

Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= UserList(VictimaIndex).CurrentInventorySlots
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).objindex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).objindex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim Num As Byte
    Dim ObjAmount As Integer
    
    ObjAmount = UserList(VictimaIndex).Invent.Object(i).Amount
    
    'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
    Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                
    MiObj.Amount = Num
    MiObj.objindex = UserList(VictimaIndex).Invent.Object(i).objindex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = ObjAmount - Num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).clase = eClass.Thief Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.objindex).Name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.objindex).Name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer
Dim pt As Long
pt = CalcularDaño(Userindex, VictimNpcIndex)

Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(Userindex).clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Cleric, eClass.Paladin, eClass.Pirat
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(Userindex).clase = eClass.Assasin Then
            daño = Round(daño * 1.4, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
 UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
 SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, UserList(Userindex).Dañoapu + daño, DAMAGE_PUÑAL)
        Call WriteConsoleMsg(Userindex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(Userindex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Su golpe total ha sido de " & Int(UserList(Userindex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT)
Call WriteConsoleMsg(Userindex, "Tu golpe total es de " & Int(UserList(Userindex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT)
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
         SendData SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y, Int(UserList(Userindex).Dañoapu + daño), DAMAGE_PUÑAL)
        Call WriteConsoleMsg(Userindex, "Has apuñalado la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        WriteConsoleMsg Userindex, "Tu golpe total es de " & Int(UserList(Userindex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, daño)
End If
    
    Call SubirSkill(Userindex, eSkill.Apuñalar, True)
Else
Call SubirSkill(Userindex, eSkill.Apuñalar, True)
    Call WriteConsoleMsg(Userindex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
    'SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, DAMAGE_NORMAL)
    Call SubirSkill(Userindex, eSkill.Apuñalar, True)
End If

End Sub

Public Sub DoAcuchillar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 12/01/2010
'***************************************************

    If UserList(Userindex).clase <> eClass.Pirat Then Exit Sub
    If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then Exit Sub

    If RandomNumber(0, 100) < PROB_ACUCHILLAR Then
        daño = Int(daño * DAÑO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(Userindex, "Has acuchillado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).Name & " te ha acuchillado por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(Userindex, "Has acuchillado a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(Userindex, VictimNpcIndex, daño)
        End If
    End If
    
End Sub

Public Sub DoGolpeCritico(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub
If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(Userindex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.75)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, daño)
    End If
End If

End Sub
Public Sub DoGolpeArco(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer
Dim pt As Byte
If UserList(Userindex).clase <> eClass.Hunter Then Exit Sub
If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then Exit Sub
'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Arco Compuesto Reforzado" Or ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Arco de Cazador" Then Exit Sub
pt = RandomNumber(0, 100)


Skill = UserList(Userindex).Stats.UserSkills(eSkill.Proyectiles)
Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
If pt > Suerte Then
    daño = Int(daño * 0.3)
   If VictimNpcIndex <> 0 Then
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, daño)
    End If
    End If

End Sub

Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Cantidad
    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call WriteUpdateSta(Userindex)
    
Exit Sub

Errhandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).clase = eClass.Worker Then
    Call QuitarSta(Userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(Userindex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).clase = eClass.Worker And UserList(Userindex).Invent.WeaponEqpObjIndex = HACHA_LEÑADOR And ObjData(ArbT).OBJType = otarboles Then
        MiObj.Amount = RandomNumber(8, 18)
        MiObj.objindex = Leña
        Else
         If UserList(Userindex).clase = eClass.Worker And UserList(Userindex).Invent.WeaponEqpObjIndex = HACHA_DORADA And ObjData(ArbT).OBJType = 38 Then
        MiObj.Amount = RandomNumber(5, 13)
        MiObj.objindex = 642
    Else
    MiObj.objindex = Leña
        MiObj.Amount = 1
    End If
    End If
    
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(Userindex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(Userindex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, talar, True)

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).clase = eClass.Worker Then
    Call QuitarSta(Userindex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(Userindex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.objindex = ObjData(UserList(Userindex).flags.TargetObj).MineralIndex
    
    If UserList(Userindex).clase = eClass.Worker Then
        MiObj.Amount = RandomNumber(5, 13) '(NicoNZ) 04/25/2008
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then _
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    
    Call WriteConsoleMsg(Userindex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 9 Then
        Call WriteConsoleMsg(Userindex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Mineria, True)

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(Userindex)
        .Counters.IdleCount = 0
        
        Dim Suerte As Integer
        Dim res As Integer
        Dim cant As Integer
        Dim MeditarSkill As Byte

        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(Userindex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(Userindex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Exit Sub
        End If
        
MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 And MeditarSkill >= -1 Then
            Suerte = 35
        ElseIf MeditarSkill <= 30 And MeditarSkill >= 11 Then
            Suerte = 30
        ElseIf MeditarSkill <= 40 And MeditarSkill >= 21 Then
            Suerte = 28
        ElseIf MeditarSkill <= 50 And MeditarSkill >= 31 Then
            Suerte = 24
        ElseIf MeditarSkill <= 60 And MeditarSkill >= 41 Then
            Suerte = 22
        ElseIf MeditarSkill <= 70 And MeditarSkill >= 51 Then
            Suerte = 20
        ElseIf MeditarSkill <= 80 And MeditarSkill >= 61 Then
            Suerte = 18
        ElseIf MeditarSkill <= 90 And MeditarSkill >= 71 Then
            Suerte = 15
        ElseIf MeditarSkill <= 100 And MeditarSkill >= 81 Then
            Suerte = 10
        ElseIf MeditarSkill < 110 And MeditarSkill >= 91 Then
            Suerte = 7
        ElseIf MeditarSkill = 100 Then
            Suerte = 5
        End If
        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)
            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant
            If .Stats.MinMAN > .Stats.MaxMAN Then _
                .Stats.MinMAN = .Stats.MaxMAN
            
                Call WriteConsoleMsg(Userindex, "¡Has recuperado " & cant & " puntos de maná!", FontTypeNames.FONTTYPE_INFO)

            
            Call WriteUpdateMana(Userindex)
            Call WriteUpdateFollow(Userindex)
            Call SubirSkill(Userindex, eSkill.Meditar, True)
        Else
            Call SubirSkill(Userindex, eSkill.Meditar, False)
        End If
    End With
End Sub

Public Sub DoDesequipar(ByVal Userindex As Integer, ByVal victimIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modif: 15/04/2010
'Unequips either shield, weapon or helmet from target user.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado As Boolean
    
    With UserList(Userindex)
        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
   End With
   
   With UserList(victimIndex)
        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(victimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(victimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(Userindex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(victimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(Userindex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    End With


End Sub

Public Sub DoHurtar(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 03/03/2010
'Implements the pick pocket skill of the Bandit :)
'03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
'***************************************************
Dim OtroUserIndex As Integer

If TriggerZonaPelea(Userindex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(Userindex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

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
                
        Call RobarObjeto(Userindex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(Userindex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
If UserList(Userindex).clase <> eClass.Thief Then Exit Sub
    

If UserList(Userindex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(Userindex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal Userindex As Integer, ByVal victimIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/04/2010 (ZaMa)
'02/04/2010: ZaMa - Nueva formula para desarmar.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    
    With UserList(Userindex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(victimIndex, UserList(victimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(victimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(victimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call FlushBuffer(victimIndex)
        End If
    End With
    
End Sub
Public Function MaxItemsConstruibles(ByVal Userindex As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    MaxItemsConstruibles = MaximoInt(1, CInt((UserList(Userindex).Stats.ELV - 4) / 5))
End Function
Public Sub ImitateNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Copies body, head and desc from previously clicked npc.
'***************************************************
    
    With UserList(Userindex)
        
        ' Copy desc
        .DescRM = Npclist(NpcIndex).Name

        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NpcIndex).Char.body
            .flags.OldHead = Npclist(NpcIndex).Char.Head
        Else
            .Char.body = Npclist(NpcIndex).Char.body
            .Char.Head = Npclist(NpcIndex).Char.Head
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    
    End With
    
End Sub
Public Sub DoEquita(ByVal Userindex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)
 
Dim ModEqui As Long
  ModEqui = ModEquitacion(UserList(Userindex).clase)
 With UserList(Userindex)
   If .Stats.UserSkills(eSkill.Equitacion) / ModEqui < Montura.MinSkill Then
       Call WriteConsoleMsg(Userindex, "Para usar esta montura necesitas " & Montura.MinSkill * ModEqui & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
  End If
.Invent.MonturaObjIndex = .Invent.Object(Slot).objindex
.Invent.MonturaSlot = Slot
 
     If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
        Exit Sub
    End If
 
   If .flags.Montando = 0 Then
       .Char.Head = 0
       If .flags.Muerto = 0 Then
           .Char.body = Montura.Ropaje
       Else
           .Char.body = iCuerpoMuerto
           .Char.Head = iCabezaMuerto
       End If
       .Char.Head = UserList(Userindex).OrigChar.Head
       .Char.ShieldAnim = NingunEscudo
       .Char.WeaponAnim = NingunArma
       .Char.CascoAnim = .Char.CascoAnim
       .flags.Montando = 1
   Else
     .flags.Montando = 0
       If .flags.Muerto = 0 Then
          .Char.Head = UserList(Userindex).OrigChar.Head
           If .Invent.ArmourEqpObjIndex > 0 Then
              .Char.body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
           Else
               Call DarCuerpoDesnudo(Userindex)
           End If
         If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim
         If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim
         If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim
     Else
        'Ds AO
        If UserList(Userindex).Faccion.FuerzasCaos <> 0 Then
        UserList(Userindex).Char.body = iCuerpoMuertoCrimi
        UserList(Userindex).Char.Head = iCabezaMuertoCrimi
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
        Else
        UserList(Userindex).Char.body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
        End If
     End If
 End If
 
 
 Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
 Call WriteMontateToggle(Userindex)
 End With
End Sub

Function ModEquitacion(ByVal clase As String) As Integer
Select Case UCase$(clase)
    Case "1"
        ModEquitacion = 1
    Case Else
        ModEquitacion = 1
End Select
 
End Function
