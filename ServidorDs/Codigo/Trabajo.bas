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

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
      '********************************************************
      'Autor: Nacho (Integer)
      'Last Modif: 11/19/2009
      'Chequea si ya debe mostrarse
      'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
      '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
      '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
      '********************************************************
10    On Error GoTo Errhandler
20        With UserList(UserIndex)
30            .Counters.TiempoOculto = .Counters.TiempoOculto - 1
40            If .Counters.TiempoOculto <= 0 Then
50                If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
60                    If .Invent.ArmourEqpObjIndex = 612 Or .Invent.ArmourEqpObjIndex = 360 Or .Invent.ArmourEqpObjIndex = 671 Then
70                        .Counters.TiempoOculto = IntervaloOculto
80                        Exit Sub
90                    End If
100               End If
110               .Counters.TiempoOculto = 0
120               .flags.Oculto = 0
                  
130               If .flags.Navegando = 1 Then
140                   If .clase = eClass.Pirat Then
                          ' Pierde la apariencia de fragata fantasmal
150                       Call ToggleBoatBody(UserIndex)
160                       Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
170                       Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                              NingunEscudo, NingunCasco)
180                   End If
190               Else
200                   If .flags.invisible = 0 Then
210                       Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
220                       Call SetInvisible(UserIndex, .Char.CharIndex, False)
230                   End If
240               End If
250           End If
260       End With
          
270       Exit Sub

Errhandler:
280       Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 13/01/2010 (ZaMa)
      'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
      'Modifique la fórmula y ahora anda bien.
      '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
      '***************************************************

10    On Error GoTo Errhandler

          Dim Suerte As Double
          Dim res As Integer
          Dim Skill As Integer
          
20        With UserList(UserIndex)
30            Skill = .Stats.UserSkills(eSkill.Ocultarse)
              
40            Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
              
50            res = RandomNumber(1, 100)
              
60            If res <= Suerte Then
              
70                .flags.Oculto = 1
80                Suerte = (-0.000001 * (100 - Skill) ^ 3)
90                Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
100               Suerte = Suerte + (-0.0088 * (100 - Skill))
110               Suerte = Suerte + (0.9571)
120               Suerte = Suerte * IntervaloOculto
130               .Counters.TiempoOculto = Suerte
                  
                  ' No es pirata o es uno sin barca
140               If .flags.Navegando = 0 Then
150                   Call SetInvisible(UserIndex, .Char.CharIndex, True)
              
160                   Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                  ' Es un pirata navegando
170               Else
                      ' Le cambiamos el body a galeon fantasmal
180                   .Char.body = iFragataFantasmal
                      ' Actualizamos clientes
190                   Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, _
                                          NingunEscudo, NingunCasco)
200               End If
                  
210               Call SubirSkill(UserIndex, eSkill.Ocultarse, True)
220           Else
                  '[CDT 17-02-2004]
230               If Not .flags.UltimoMensaje = 4 Then
240                   Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
250                   .flags.UltimoMensaje = 4
260               End If
                  '[/CDT]
                  
270               Call SubirSkill(UserIndex, eSkill.Ocultarse, False)
280           End If
              
290           .Counters.Ocultando = .Counters.Ocultando + 1
300       End With
          
310       Exit Sub

Errhandler:
320       Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 13/01/2010 (ZaMa)
      '13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
      '16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
      '10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
      '***************************************************

          Dim ModNave As Single
          
10        With UserList(UserIndex)
20            ModNave = ModNavegacion(.clase, UserIndex)
30    If HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) = True And HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y) = True And HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y) = True And _
      HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1) = True And HayAgua(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1) = True Then
40        Call WriteConsoleMsg(UserIndex, "No puedes dejar de navegar en el agua!!", FontTypeNames.FONTTYPE_INFO)
50        Exit Sub
60    End If
70            If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
80                Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
90                Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
100               Exit Sub
110           End If
              
120                                            If .flags.Montando = 1 Then
130   Call WriteConsoleMsg(UserIndex, "¡No puedes navegar si estás montando!", FontTypeNames.FONTTYPE_INFO)
140                   Exit Sub
150   End If

160                                    If .flags.Mimetizado = 1 Then
170   Call WriteConsoleMsg(UserIndex, "¡No puedes mimetizarte si estás navegando!", FontTypeNames.FONTTYPE_INFO)
180                   Exit Sub
190   End If
              
200       UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
210   UserList(UserIndex).Invent.BarcoSlot = Slot

220   If UserList(UserIndex).flags.Navegando = 0 Then
230   End If
              
              ' No estaba navegando
240           If .flags.Navegando = 0 Then
250               .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
260               .Invent.BarcoSlot = Slot
                  
270               .Char.Head = 0
                  
                     ' No esta muerto
280               If .flags.Muerto = 0 Then
                  
290                   Call ToggleBoatBody(UserIndex)
                      
                      ' Pierde el ocultar
300                   If .flags.Oculto = 1 Then
310                       .flags.Oculto = 0
320                       Call SetInvisible(UserIndex, .Char.CharIndex, False)
330                       Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
340                   End If
                     
                      ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
350                   If .flags.invisible = 1 Then
360                       Call SetInvisible(UserIndex, .Char.CharIndex, False)
370                   End If
                      
                  ' Esta muerto
380               Else
390                   .Char.body = iFragataFantasmal
400                   .Char.ShieldAnim = NingunEscudo
410                   .Char.WeaponAnim = NingunArma
420                   .Char.CascoAnim = NingunCasco
430               End If
                  
                  ' Comienza a navegar
440               .flags.Navegando = 1
              
              ' Estaba navegando
450           Else
460               .Invent.BarcoObjIndex = 0
470               .Invent.BarcoSlot = 0
              
                  ' No esta muerto
480               If .flags.Muerto = 0 Then
490                   .Char.Head = .OrigChar.Head
                      
500                   If .clase = eClass.Pirat Then
510                       If .flags.Oculto = 1 Then
                              ' Al desequipar barca, perdió el ocultar
520                           .flags.Oculto = 0
530                           .Counters.Ocultando = 0
540                           Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
550                       End If
560                   End If
                      
570                   If .Invent.ArmourEqpObjIndex > 0 Then
580                       .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
590                   Else
600                       Call DarCuerpoDesnudo(UserIndex)
610                   End If
                      
620                   If .Invent.EscudoEqpObjIndex > 0 Then _
                          .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
630                   If .Invent.WeaponEqpObjIndex > 0 Then _
                          .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)
640                   If .Invent.CascoEqpObjIndex > 0 Then _
                          .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                      
                      
                      ' Al dejar de navegar, si estaba invisible actualizo los clientes
650                   If .flags.invisible = 1 Then
660                       Call SetInvisible(UserIndex, .Char.CharIndex, True)
670                   End If
                      
                  ' Esta muerto
680               Else
690                   .Char.body = iCuerpoMuerto
700                   .Char.Head = iCabezaMuerto
710                   .Char.ShieldAnim = NingunEscudo
720                   .Char.WeaponAnim = NingunArma
730                   .Char.CascoAnim = NingunCasco
740               End If
                  
                  ' Termina de navegar
750               .flags.Navegando = 0
760           End If
              
              ' Actualizo clientes
770           Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
780       End With
          
790       Call WriteNavigateToggle(UserIndex)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

20        With UserList(UserIndex)
30            If .flags.TargetObjInvIndex > 0 Then
                 
40               If ObjData(.flags.TargetObjInvIndex).ObjType = eOBJType.otMinerales And _
                      ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) / ModFundicion(.clase) Then
50                    Call DoLingotes(UserIndex)
60               Else
70                    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
80               End If
              
90            End If
100       End With

110       Exit Sub

Errhandler:
120       Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.Description)

End Sub

Function TieneObjetosOffline(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserName As String) As Boolean
      '***************************************************
      'Author: -
      'Last Modification: 24/04/2019
      ' Copy of the TieneObjetos Default AO, adapted to offline
      '***************************************************

          Dim i As Integer
          Dim FilePath As String
          Dim Temp As String
          Dim Obj As Obj
          Dim Total As Long
          
          FilePath = CharPath & UCase$(UserName) & ".chr"
          
10        For i = 1 To MAX_INVENTORY_SLOTS
              Temp = GetVar(FilePath, "INVENTORY", "OBJ" & i)
              Obj.ObjIndex = val(ReadField(1, Temp, Asc("-")))
              Obj.Amount = val(ReadField(2, Temp, Asc("-")))
              
20            If Obj.ObjIndex = ItemIndex Then
30                Total = Total + Obj.Amount
40            End If
50        Next i
          
60        If cant <= Total Then
70            TieneObjetosOffline = True
80            Exit Function
90        End If
              
End Function
Public Sub QuitarObjetosOffline(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserName As String)
      '***************************************************
      'Author: -
      'Last Modification: 24/04/2019
      '***************************************************

        Dim i As Integer
        Dim Temp As String
        Dim Obj As Obj
        Dim FilePath As String
        Dim Equipped As Byte
        Dim NroItems As Byte
        Dim ObjType As eOBJType
        Dim tmpObjType As String
        
        FilePath = CharPath & UCase$(UserName) & ".chr"
          
        NroItems = val(GetVar(FilePath, "INVENTORY", "CANTIDADITEMS"))
        
        For i = 1 To MAX_INVENTORY_SLOTS
            Temp = GetVar(FilePath, "INVENTORY", "OBJ" & i)
            Obj.ObjIndex = val(ReadField(1, Temp, Asc("-")))
            Obj.Amount = val(ReadField(2, Temp, Asc("-")))
            Equipped = val(ReadField(3, Temp, Asc("-")))
              
            If Obj.ObjIndex = ItemIndex Then
                If Obj.Amount <= cant And Equipped = 1 Then Equipped = 0
                    
                Obj.Amount = Obj.Amount - cant
                
                If Obj.Amount <= 0 Then
                    cant = Abs(Obj.Amount)
                    
                    Select Case ObjType
                        Case eOBJType.otAnilloNpc: tmpObjType = "ANILLONPCSLOT"
                        Case eOBJType.otAnillo: tmpObjType = "ANILLOSLOT"
                        Case eOBJType.otArmadura: tmpObjType = "ARMOUREQPSLOT"
                        Case eOBJType.otBarcos: tmpObjType = "BARCOSLOT"
                        Case eOBJType.otCasco: tmpObjType = "CASCOEQPSLOT"
                        Case eOBJType.otEscudo: tmpObjType = "ESCUDOEQPSLOT"
                        Case eOBJType.otMochilas: tmpObjType = "MOCHILASLOT"
                        Case eOBJType.otMonturas: tmpObjType = "MONTURASLOT"
                        Case eOBJType.otFlechas: tmpObjType = "MUNICIONSLOT"
                    End Select
                    
                    If tmpObjType <> vbNullString Then WriteVar FilePath, "INVENTORY", tmpObjType, "0"
                    WriteVar FilePath, "INVENTORY", "CANTIDADITEMS", (NroItems - 1)
                    WriteVar FilePath, "INVENTORY", "OBJ" & i, "0-0-0"
                    
                    Obj.Amount = 0
                    Obj.ObjIndex = 0
                Else
                    cant = 0
                    WriteVar FilePath, "INVENTORY", "OBJ" & i, Obj.ObjIndex & "-" & Obj.Amount & "-" & Equipped
                    Exit For
                End If
                
                'If cant = 0 Then Exit Sub
            End If
        Next i
End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Integer
          Dim Total As Long
10        For i = 1 To UserList(UserIndex).CurrentInventorySlots
20            If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
30                Total = Total + UserList(UserIndex).Invent.Object(i).Amount
40            End If
50        Next i
          
60        If cant <= Total Then
70            TieneObjetos = True
80            Exit Function
90        End If
              
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 05/08/09
      '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
      '***************************************************

          Dim i As Integer
10        For i = 1 To UserList(UserIndex).CurrentInventorySlots
20            With UserList(UserIndex).Invent.Object(i)
30                If .ObjIndex = ItemIndex Then
40                    If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                      
50                    .Amount = .Amount - cant
60                    If .Amount <= 0 Then
70                        cant = Abs(.Amount)
80                        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
90                        .Amount = 0
100                       .ObjIndex = 0
110                   Else
120                       cant = 0
130                   End If
                      
140                   Call UpdateUserInv(False, UserIndex, i)
                      
150                   If cant = 0 Then Exit Sub
160               End If
170           End With
180       Next i

End Sub

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
10        If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
20        If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
30        If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
10        If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
          
10        If ObjData(ItemIndex).Madera > 0 Then
20                If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
30                        Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
40                        CarpinteroTieneMateriales = False
50                        Exit Function
60                End If
70        End If
          
80        CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
10        If ObjData(ItemIndex).LingH > 0 Then
20                If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
30                        Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
40                        HerreroTieneMateriales = False
50                        Exit Function
60                End If
70        End If
80        If ObjData(ItemIndex).LingP > 0 Then
90                If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
100                       Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
110                       HerreroTieneMateriales = False
120                       Exit Function
130               End If
140       End If
150       If ObjData(ItemIndex).LingO > 0 Then
160               If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
170                       Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
180                       HerreroTieneMateriales = False
190                       Exit Function
200               End If
210       End If
220       HerreroTieneMateriales = True
End Function

Function TieneMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 12/08/2009
      '
      '***************************************************
          Dim ItemUpgrade As Integer
          
10        ItemUpgrade = ObjData(ItemIndex).Upgrade
          
20        With ObjData(ItemUpgrade)
30            If .LingH > 0 Then
40                If Not TieneObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
50                    Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
60                    TieneMaterialesUpgrade = False
70                    Exit Function
80                End If
90            End If
              
100           If .LingP > 0 Then
110               If Not TieneObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
120                   Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
130                   TieneMaterialesUpgrade = False
140                   Exit Function
150               End If
160           End If
              
170           If .LingO > 0 Then
180               If Not TieneObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
190                   Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
200                   TieneMaterialesUpgrade = False
210                   Exit Function
220               End If
230           End If
              
240           If .Madera > 0 Then
250               If Not TieneObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
260                   Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
270                   TieneMaterialesUpgrade = False
280                   Exit Function
290               End If
300           End If
              
310           If .MaderaElfica > 0 Then
320               If Not TieneObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
330                   Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
340                   TieneMaterialesUpgrade = False
350                   Exit Function
360               End If
370           End If
380       End With
          
390       TieneMaterialesUpgrade = True
End Function

Sub QuitarMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 12/08/2009
      '
      '***************************************************
          Dim ItemUpgrade As Integer
          
10        ItemUpgrade = ObjData(ItemIndex).Upgrade
          
20        With ObjData(ItemUpgrade)
30            If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
40            If .LingP > 0 Then Call QuitarObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
50            If .LingO > 0 Then Call QuitarObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
60            If .Madera > 0 Then Call QuitarObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
70            If .MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
80        End With
          
90        Call QuitarObjetos(ItemIndex, 1, UserIndex)
End Sub

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
10    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.herreria) >= _
       ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
      Dim i As Long

10    For i = 1 To UBound(ArmasHerrero)
20        If ArmasHerrero(i) = ItemIndex Then
30            PuedeConstruirHerreria = True
40            Exit Function
50        End If
60    Next i
70    For i = 1 To UBound(ArmadurasHerrero)
80        If ArmadurasHerrero(i) = ItemIndex Then
90            PuedeConstruirHerreria = True
100           Exit Function
110       End If
120   Next i
130   PuedeConstruirHerreria = False
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

10    If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
20        Call HerreroQuitarMateriales(UserIndex, ItemIndex)
          ' AGREGAR FX
30        If ObjData(ItemIndex).ObjType = eOBJType.otWeapon Then
40            Call WriteConsoleMsg(UserIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
50        ElseIf ObjData(ItemIndex).ObjType = eOBJType.otEscudo Then
60            Call WriteConsoleMsg(UserIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
70        ElseIf ObjData(ItemIndex).ObjType = eOBJType.otCasco Then
80            Call WriteConsoleMsg(UserIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
90        ElseIf ObjData(ItemIndex).ObjType = eOBJType.otArmadura Then
100           Call WriteConsoleMsg(UserIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
110       End If
          Dim MiObj As Obj
120       MiObj.Amount = 1
130       MiObj.ObjIndex = ItemIndex
140       If Not MeterItemEnInventario(UserIndex, MiObj) Then
150           Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
160       End If
          
          'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
170       If ObjData(MiObj.ObjIndex).LOG = 1 Then
180           Call LogDesarrollo(UserList(UserIndex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
190       End If
          
          'Call SubirSkill(UserIndex, herreria)
200       Call UpdateUserInv(True, UserIndex, 0)
210       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

220       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
230       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
              UserList(UserIndex).Reputacion.PlebeRep = MAXREP

240       UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
250   End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
      Dim i As Long

10    For i = 1 To UBound(ObjCarpintero)
20        If ObjCarpintero(i) = ItemIndex Then
30            PuedeConstruirCarpintero = True
40            Exit Function
50        End If
60    Next i
70    PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

10    If CarpinteroTieneMateriales(UserIndex, ItemIndex) And _
         UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
         ObjData(ItemIndex).SkCarpinteria And _
         PuedeConstruirCarpintero(ItemIndex) And _
         UserList(UserIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
          
20        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
30        Call WriteConsoleMsg(UserIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
          
          Dim MiObj As Obj
40        MiObj.Amount = 1
50        MiObj.ObjIndex = ItemIndex
60        If Not MeterItemEnInventario(UserIndex, MiObj) Then
70                        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
80        End If
          
          'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
90        If ObjData(MiObj.ObjIndex).LOG = 1 Then
100           Call LogDesarrollo(UserList(UserIndex).Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
110       End If
          
          'Call SubirSkill(UserIndex, Carpinteria)
120       Call UpdateUserInv(True, UserIndex, 0)
130       Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))


140       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
150       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
              UserList(UserIndex).Reputacion.PlebeRep = MAXREP

160       UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

170   End If
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
10        Select Case Lingote
              Case iMinerales.HierroCrudo
20                MineralesParaLingote = 25
30            Case iMinerales.PlataCruda
40                MineralesParaLingote = 35
50            Case iMinerales.OroCrudo
60                MineralesParaLingote = 50
70            Case Else
80                MineralesParaLingote = 10000
90        End Select
End Function


Public Sub DoLingotes(ByVal UserIndex As Integer)
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
          
10        With UserList(UserIndex)
20            If .flags.Comerciando Then
30                OtroUserIndex = .ComUsu.DestUsu
                      
40                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
50                    Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
60                    Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                      
70                    Call LimpiarComercioSeguro(UserIndex)
80                    Call Protocol.FlushBuffer(OtroUserIndex)
90                End If
100           End If
              
110           CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

120           Slot = .flags.TargetObjInvSlot
130           obji = .Invent.Object(Slot).ObjIndex
              
140           While CantidadItems > 0 And Not TieneMinerales
150               If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
160                   TieneMinerales = True
170               Else
180                   CantidadItems = CantidadItems - 1
190               End If
200           Wend
              
210           If Not TieneMinerales Or ObjData(obji).ObjType <> eOBJType.otMinerales Then
220               Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
230               Exit Sub
240           End If
              
250           .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems
260           If .Invent.Object(Slot).Amount < 1 Then
270               .Invent.Object(Slot).Amount = 0
280               .Invent.Object(Slot).ObjIndex = 0
290           End If
              
              Dim MiObj As Obj
300           MiObj.Amount = CantidadItems
310           MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex
320           If Not MeterItemEnInventario(UserIndex, MiObj) Then
330               Call TirarItemAlPiso(.Pos, MiObj)
340           End If
              
350           Call UpdateUserInv(False, UserIndex, Slot)
360           Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & _
                                  IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
          
370           .Counters.Trabajando = .Counters.Trabajando + 1
380       End With
End Sub

Public Sub DoUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modification: 12/08/2009
      '12/08/2009: Pato - Implementado nuevo sistema de mejora de items
      '***************************************************
      Dim ItemUpgrade As Integer
      Dim WeaponIndex As Integer
      Dim OtroUserIndex As Integer

10    ItemUpgrade = ObjData(ItemIndex).Upgrade

20    With UserList(UserIndex)
30        If .flags.Comerciando Then
40            OtroUserIndex = .ComUsu.DestUsu
                  
50            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
60                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
70                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                  
80                Call LimpiarComercioSeguro(UserIndex)
90                Call Protocol.FlushBuffer(OtroUserIndex)
100           End If
110       End If
              
          'Sacamos energía
120       If .clase = eClass.Worker Then
              'Chequeamos que tenga los puntos antes de sacarselos
130           If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
140               .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
150               Call WriteUpdateSta(UserIndex)
160           Else
170               Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
180               Exit Sub
190           End If
200       Else
              'Chequeamos que tenga los puntos antes de sacarselos
210           If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
220               .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
230               Call WriteUpdateSta(UserIndex)
240           Else
250               Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
260               Exit Sub
270           End If
280       End If
          
290       If ItemUpgrade <= 0 Then Exit Sub
300       If Not TieneMaterialesUpgrade(UserIndex, ItemIndex) Then Exit Sub
          
310       If PuedeConstruirHerreria(ItemUpgrade) Then
              
320           WeaponIndex = .Invent.WeaponEqpObjIndex
          
330           If WeaponIndex <> MARTILLO_HERRERO Then
340               Call WriteConsoleMsg(UserIndex, "Debes equiparte el martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
350               Exit Sub
360           End If
              
370           If Round(.Stats.UserSkills(eSkill.herreria) / ModHerreriA(.clase), 0) < ObjData(ItemUpgrade).SkHerreria Then
380               Call WriteConsoleMsg(UserIndex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
390               Exit Sub
400           End If
              
410           Select Case ObjData(ItemIndex).ObjType
                  Case eOBJType.otWeapon
420                   Call WriteConsoleMsg(UserIndex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                      
430               Case eOBJType.otEscudo 'Todavía no hay, pero just in case
440                   Call WriteConsoleMsg(UserIndex, "Has mejorado el escudo!", FontTypeNames.FONTTYPE_INFO)
                  
450               Case eOBJType.otCasco
460                   Call WriteConsoleMsg(UserIndex, "Has mejorado el casco!", FontTypeNames.FONTTYPE_INFO)
                  
470               Case eOBJType.otArmadura
480                   Call WriteConsoleMsg(UserIndex, "Has mejorado la armadura!", FontTypeNames.FONTTYPE_INFO)
490           End Select
              
500           Call SubirSkill(UserIndex, eSkill.herreria, True)
510           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.X, .Pos.Y))
          
520       ElseIf PuedeConstruirCarpintero(ItemUpgrade) Then
              
530           WeaponIndex = .Invent.WeaponEqpObjIndex
540           If WeaponIndex <> SERRUCHO_CARPINTERO Then
550               Call WriteConsoleMsg(UserIndex, "Debes equiparte un serrucho.", FontTypeNames.FONTTYPE_INFO)
560               Exit Sub
570           End If
              
580           If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.clase), 0) < ObjData(ItemUpgrade).SkCarpinteria Then
590               Call WriteConsoleMsg(UserIndex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
600               Exit Sub
610           End If
              
620           Select Case ObjData(ItemIndex).ObjType
                  Case eOBJType.otFlechas
630                   Call WriteConsoleMsg(UserIndex, "Has mejorado la flecha!", FontTypeNames.FONTTYPE_INFO)
                      
640               Case eOBJType.otWeapon
650                   Call WriteConsoleMsg(UserIndex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                      
660               Case eOBJType.otBarcos
670                   Call WriteConsoleMsg(UserIndex, "Has mejorado el barco!", FontTypeNames.FONTTYPE_INFO)
680           End Select
              
690           Call SubirSkill(UserIndex, eSkill.Carpinteria, True)
700           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.X, .Pos.Y))
710       Else
720           Exit Sub
730       End If
          
740       Call QuitarMaterialesUpgrade(UserIndex, ItemIndex)
          
          Dim MiObj As Obj
750       MiObj.Amount = 1
760       MiObj.ObjIndex = ItemUpgrade
          
770       If Not MeterItemEnInventario(UserIndex, MiObj) Then
780           Call TirarItemAlPiso(.Pos, MiObj)
790       End If
          
800       If ObjData(ItemIndex).LOG = 1 Then _
              Call LogDesarrollo(.Name & " ha mejorado el ítem " & ObjData(ItemIndex).Name & " a " & ObjData(ItemUpgrade).Name)
              
810       .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta
820       If .Reputacion.PlebeRep > MAXREP Then _
              .Reputacion.PlebeRep = MAXREP
              
830       .Counters.Trabajando = .Counters.Trabajando + 1
840   End With
End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal UserIndex As Integer) As Single
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 27/11/2009
      '27/11/2009: ZaMa - A worker can navigate before only if it's an expert fisher
      '12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
      '***************************************************
10    Select Case clase
          Case eClass.Pirat
20            ModNavegacion = 1
30        Case eClass.Worker
40            If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) = 100 Then
50                ModNavegacion = 1.71
60            Else
70                ModNavegacion = 2
80            End If
90        Case Else
100           ModNavegacion = 2
110   End Select

End Function


Function ModFundicion(ByVal clase As eClass) As Single
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    Select Case clase
          Case eClass.Worker
20            ModFundicion = 1
30        Case Else
40            ModFundicion = 3
50    End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    Select Case clase
          Case eClass.Worker
20            ModCarpinteria = 1
30        Case Else
40            ModCarpinteria = 3
50    End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
10    Select Case clase
          Case eClass.Worker
20            ModHerreriA = 1
30        Case Else
40            ModHerreriA = 4
50    End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
10        Select Case clase
              Case eClass.Druid
20                ModDomar = 6
30            Case eClass.Hunter
40                ModDomar = 6
50            Case eClass.Cleric
60                ModDomar = 7
70            Case Else
80                ModDomar = 10
90        End Select
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: 02/03/09
      '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
      '***************************************************
          Dim j As Integer
10        For j = 1 To MAXMASCOTAS
20            If UserList(UserIndex).MascotasType(j) = 0 Then
30                FreeMascotaIndex = j
40                Exit Function
50            End If
60        Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Nacho (Integer)
      'Last Modification: 02/03/2009
      '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
      '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
      '***************************************************

10    On Error GoTo Errhandler

      Dim puntosDomar As Integer
      Dim puntosRequeridos As Integer
      Dim CanStay As Boolean
      Dim petType As Integer
      Dim NroPets As Integer


      If UserList(UserIndex).clase <> eClass.Druid Then
            WriteConsoleMsg UserIndex, "Solo los Druidas tienen el poder de la doma de animales.", FontTypeNames.FONTTYPE_INFO
            Exit Sub
      End If
        
20    If Npclist(NpcIndex).MaestroUser = UserIndex Then
30        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
40        Exit Sub
50    End If

60    If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
          
70        If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
80            Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
90            Exit Sub
100       End If
          
        '  If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
         '     Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
          '    Exit Sub
       '   End If
          
110       puntosDomar = CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(UserIndex).Stats.UserSkills(eSkill.Domar))
120       If UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAELFICA And UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
130           puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
140       Else
150           puntosRequeridos = Npclist(NpcIndex).flags.Domable
160       End If
          
170       If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
              Dim Index As Integer
180           UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
190           Index = FreeMascotaIndex(UserIndex)
200           UserList(UserIndex).MascotasIndex(Index) = NpcIndex
210           UserList(UserIndex).MascotasType(Index) = Npclist(NpcIndex).Numero
              
220           Npclist(NpcIndex).MaestroUser = UserIndex
              
230           Call FollowAmo(NpcIndex)
240           Call ReSpawnNpc(Npclist(NpcIndex))
              
250           Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
              
              ' Es zona segura?
260           CanStay = (MapInfo(UserList(UserIndex).Pos.map).Pk = True)
              
270           If Not CanStay Then
280               petType = Npclist(NpcIndex).Numero
290               NroPets = UserList(UserIndex).NroMascotas
                  
300               Call QuitarNPC(NpcIndex)
                  
310              UserList(UserIndex).MascotasType(Index) = petType
320               UserList(UserIndex).NroMascotas = NroPets
                  
                '  Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
330           End If

340       Else
350           If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
360               Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
370               UserList(UserIndex).flags.UltimoMensaje = 5
380           End If
390       End If
          
          'Entreno domar. Es un 30% más dificil si no sos druida.
400       If UserList(UserIndex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
410           Call SubirSkill(UserIndex, Domar, True)
420       End If
430   Else
440       Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
450   End If

460   Exit Sub

Errhandler:
470       Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)

End Sub
''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: ZaMa
      'This function checks how many NPCs of the same type have
      'been tamed by the user.
      'Returns True if that amount is less than two.
      '***************************************************
          Dim i As Long
          Dim numMascotas As Long
          
10        For i = 1 To MAXMASCOTAS
20            If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
30                numMascotas = numMascotas + 1
40            End If
50        Next i
          
60        If numMascotas <= 1 Then PuedeDomarMascota = True
          
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 12/01/2010 (ZaMa)
      'Makes an admin invisible o visible.
      '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
      '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
      '***************************************************
          
10        With UserList(UserIndex)
20            If .flags.AdminInvisible = 0 Then
                  ' Sacamos el mimetizmo
30                If .flags.Mimetizado = 1 Then
40                    .Char.body = .CharMimetizado.body
50                    .Char.Head = .CharMimetizado.Head
60                    .Char.CascoAnim = .CharMimetizado.CascoAnim
70                    .Char.ShieldAnim = .CharMimetizado.ShieldAnim
80                    .Char.WeaponAnim = .CharMimetizado.WeaponAnim
90                    .Counters.Mimetismo = 0
100                   .flags.Mimetizado = 0
                      ' Se fue el efecto del mimetismo, puede ser atacado por npcs
110                   .flags.Ignorado = False
120               End If
                  
130               .flags.AdminInvisible = 1
140               .flags.invisible = 1
150               .flags.Oculto = 1
160               .flags.OldBody = .Char.body
170               .flags.OldHead = .Char.Head
180               .Char.body = 0
190               .Char.Head = 0
                  
                  ' Solo el admin sabe que se hace invi
200               Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                  'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
210               Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
220           Else
230               .flags.AdminInvisible = 0
240               .flags.invisible = 0
250               .flags.Oculto = 0
260               .Counters.TiempoOculto = 0
270               .Char.body = .flags.OldBody
280               .Char.Head = .flags.OldHead
                  
                  ' Solo el admin sabe que se hace visible
290               Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.Heading, _
                  .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim))
300               Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                   
                  'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
310               Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
320           End If
330       End With
          
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim Suerte As Byte
      Dim exito As Byte
      Dim Obj As Obj
      Dim posMadera As WorldPos

10    If Not LegalPos(map, X, Y) Then Exit Sub

20    With posMadera
30        .map = map
40        .X = X
50        .Y = Y
60    End With

70    If MapData(map, X, Y).ObjInfo.ObjIndex <> 58 Then
80        Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre leña para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
90        Exit Sub
100   End If

110   If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
120       Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
130       Exit Sub
140   End If

150   If UserList(UserIndex).flags.Muerto = 1 Then
160       Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
170       Exit Sub
180   End If

190   If MapData(map, X, Y).ObjInfo.Amount < 3 Then
200       Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
210       Exit Sub
220   End If

      Dim SupervivenciaSkill As Byte

230   SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

240   If SupervivenciaSkill >= 0 And SupervivenciaSkill < 6 Then
250       Suerte = 3
260   ElseIf SupervivenciaSkill >= 6 And SupervivenciaSkill <= 34 Then
270       Suerte = 2
280   ElseIf SupervivenciaSkill >= 35 Then
290       Suerte = 1
300   End If

310   exito = RandomNumber(1, Suerte)

320   If exito = 1 Then
330       Obj.ObjIndex = FOGATA_APAG
340       Obj.Amount = MapData(map, X, Y).ObjInfo.Amount \ 3
          
350       Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
          
360       Call MakeObj(Obj, map, X, Y)
          
          'Seteamos la fogata como el nuevo TargetObj del user
370       UserList(UserIndex).flags.TargetObj = FOGATA_APAG
          
380       Call SubirSkill(UserIndex, eSkill.Supervivencia, True)
390   Else
          '[CDT 17-02-2004]
400       If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
410           Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
420           UserList(UserIndex).flags.UltimoMensaje = 10
430       End If
          '[/CDT]
          
440       Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
450   End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 16/11/2009
      '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
      '***************************************************
10    On Error GoTo Errhandler

      Dim Suerte As Integer
      Dim res As Integer
      Dim CantidadItems As Integer
    

      
20    If UserList(UserIndex).clase = eClass.Worker Then
30        Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
40    Else
50        Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
60    End If

      Dim Skill As Integer
70    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)
80    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

90    res = RandomNumber(1, Suerte)

100   If res <= 6 Then
          Dim MiObj As Obj
          
110       If UserList(UserIndex).clase = eClass.Worker Then
120           With UserList(UserIndex)
130               CantidadItems = 1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))
140           End With
              
150           MiObj.Amount = RandomNumber(1, CantidadItems)
160       Else
170           MiObj.Amount = 1
180       End If

190       MiObj.ObjIndex = Pescado
          
200       If Not MeterItemEnInventario(UserIndex, MiObj) Then
210           Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
220       End If
          
230       Call WriteConsoleMsg(UserIndex, "¡Has sacado un pez del agua!", FontTypeNames.FONTTYPE_INFO)
          
240       Call SubirSkill(UserIndex, eSkill.Pesca, True)
250   Else
          '[CDT 17-02-2004]
260       If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
270         Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
280         UserList(UserIndex).flags.UltimoMensaje = 6
290       End If
          '[/CDT]
          
300       Call SubirSkill(UserIndex, eSkill.Pesca, False)
310   End If

320   UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
330   If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
          UserList(UserIndex).Reputacion.PlebeRep = MAXREP

340   UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

350   Exit Sub

Errhandler:
360       Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer, Optional ByVal CofreDrop As Boolean = False)

      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
10    On Error GoTo Errhandler

      Dim iSkill As Integer
      Dim Suerte As Integer
      Dim res As Integer
      Dim EsPescador As Boolean

      If CofreDrop Then
            mCofres.UsuarioPescaCofres (UserIndex)
      End If
      
20    If UserList(UserIndex).clase = eClass.Worker Then
30        Call QuitarSta(UserIndex, 1)
40        EsPescador = True
50    Else
60        Call QuitarSta(UserIndex, 1)
70        EsPescador = False
80    End If

      If UserList(UserIndex).Stats.MinSta <= 0 Then
        If UserList(UserIndex).Counters.Trabajando Then _
              UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
      End If

90    iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

      ' m = (60-11)/(1-10)
      ' y = mx - m*10 + 11

100   Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        
110   If Suerte > 0 Then
120       res = RandomNumber(1, Suerte)
          
130       If res < 6 Then
              Dim MiObj As Obj
              Dim PecesPosibles(1 To 5) As Integer
              
140           PecesPosibles(1) = PESCADO1
150           PecesPosibles(2) = PESCADO2
160           PecesPosibles(3) = PESCADO3
170           PecesPosibles(4) = PESCADO4
              PecesPosibles(5) = PESCADO5
               
180           If EsPescador = True Then
190               MiObj.Amount = RandomNumber(1, 5)
200           Else
210               MiObj.Amount = 1
220           End If

230           MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
              
240           If Not MeterItemEnInventario(UserIndex, MiObj) Then
250               Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
260           End If
              
270           Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
              
280           Call SubirSkill(UserIndex, eSkill.Pesca, True)
290       Else
300           Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
310           Call SubirSkill(UserIndex, eSkill.Pesca, False)
320       End If
330   End If

340       UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
350       If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
              UserList(UserIndex).Reputacion.PlebeRep = MAXREP
              
360   Exit Sub

Errhandler:
370       Call LogError("Error en DoPescarRed")
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

10    On Error GoTo Errhandler

          Dim OtroUserIndex As Integer

20        If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub
          
30        If UserList(VictimaIndex).flags.EnConsulta Then
40            Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
50            Exit Sub
60        End If
          
70        With UserList(LadrOnIndex)
          
80            If .flags.Seguro Then
90                If Not criminal(VictimaIndex) Then
100                   Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
110                   Exit Sub
120               End If
130           Else
140               If .Faccion.ArmadaReal = 1 Then
150                   If Not criminal(VictimaIndex) Then
160                       Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejército real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
170                       Exit Sub
180                   End If
190               End If
200           End If
              
              ' Caos robando a caos?
210           If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
220               Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legión oscura.", FontTypeNames.FONTTYPE_FIGHT)
230               Exit Sub
240           End If
              
250           If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
          
              
              ' Tiene energia?
260           If .Stats.MinSta < 15 Then
270               If .Genero = eGenero.Hombre Then
280                   Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
290               Else
300                   Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)
310               End If
                  
320               Exit Sub
330           End If
              
              ' Quito energia
340           Call QuitarSta(LadrOnIndex, 15)
              
              Dim GuantesHurto As Boolean
          
350           If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
              
360           If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
                  
                  Dim Suerte As Integer
                  Dim res As Integer
                  Dim RobarSkill As Byte
                  
370               RobarSkill = .Stats.UserSkills(eSkill.Robar)
                      
380               If RobarSkill <= 10 Then
390                   Suerte = 35
400               ElseIf RobarSkill <= 20 Then
410                   Suerte = 30
420               ElseIf RobarSkill <= 30 Then
430                   Suerte = 28
440               ElseIf RobarSkill <= 40 Then
450                   Suerte = 24
460               ElseIf RobarSkill <= 50 Then
470                   Suerte = 22
480               ElseIf RobarSkill <= 60 Then
490                   Suerte = 20
500               ElseIf RobarSkill <= 70 Then
510                   Suerte = 18
520               ElseIf RobarSkill <= 80 Then
530                   Suerte = 15
540               ElseIf RobarSkill <= 90 Then
550                   Suerte = 10
560               ElseIf RobarSkill < 100 Then
570                   Suerte = 7
580               Else
590                   Suerte = 5
600               End If
                  
610               res = RandomNumber(1, Suerte)
                      
620               If res < 3 Then 'Exito robo
630                   If UserList(VictimaIndex).flags.Comerciando Then
640                       OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                              
650                       If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
660                           Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
670                           Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                              
680                           Call LimpiarComercioSeguro(VictimaIndex)
690                           Call Protocol.FlushBuffer(OtroUserIndex)
700                       End If
710                   End If
                     
720                   If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
730                       If TieneObjetosRobables(VictimaIndex) Then
740                           Call RobarObjeto(LadrOnIndex, VictimaIndex)
750                       Else
760                           Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
770                       End If
780                   Else 'Roba oro
790                       If UserList(VictimaIndex).Stats.Gld > 0 Then
                              Dim N As Long
                              
800                           If .clase = eClass.Thief Then
                              ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
810                               If GuantesHurto Then
820                                   N = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
830                               Else
840                                   N = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)
850                               End If
860                           Else
870                               N = RandomNumber(1, 100)
880                           End If
890                           If N > UserList(VictimaIndex).Stats.Gld Then N = UserList(VictimaIndex).Stats.Gld
900                           UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - N
                              
910                           .Stats.Gld = .Stats.Gld + N
920                           If .Stats.Gld > MaxOro Then _
                                  .Stats.Gld = MaxOro
                              
930                           Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
940                           Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                              
950                           Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
960                           Call FlushBuffer(VictimaIndex)
970                       Else
980                           Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
990                       End If
1000                  End If
                      
1010                  Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
1020              Else
1030                  Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
1040                  Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
1050                  Call FlushBuffer(VictimaIndex)
                      
1060                  Call SubirSkill(LadrOnIndex, eSkill.Robar, False)
1070              End If
              
1080              If Not criminal(LadrOnIndex) Then
1090                  If Not criminal(VictimaIndex) Then
1100                      Call VolverCriminal(LadrOnIndex)
1110                  End If
1120              End If
                  
                  ' Se pudo haber convertido si robo a un ciuda
1130              If criminal(LadrOnIndex) Then
1140                  .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron
1150                  If .Reputacion.LadronesRep > MAXREP Then _
                          .Reputacion.LadronesRep = MAXREP
1160              End If
1170          End If
1180      End With

1190  Exit Sub

Errhandler:
1200      Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

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

10    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

20    ObjEsRobable = _
      ObjData(OI).ObjType <> eOBJType.otLlaves And _
      UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
      ObjData(OI).Real = 0 And _
      ObjData(OI).Caos = 0 And _
      ObjData(OI).ObjType <> eOBJType.otMonturas And _
      ObjData(OI).ObjType <> eOBJType.otMonturasDraco And _
      ObjData(OI).VIP = 0 And _
      ObjData(OI).VIPP = 0 And _
      ObjData(OI).VIPB = 0 And _
      ObjData(OI).UM = 0 And _
      ObjData(OI).HM = 0 And _
      ObjData(OI).NoSeCae = 0 And _
      ObjData(OI).Newbie = 0 And _
      ObjData(OI).ObjType <> eOBJType.otBarcos

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
10    flag = False

20    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
30        i = 1
40        Do While Not flag And i <= UserList(VictimaIndex).CurrentInventorySlots
              'Hay objeto en este slot?
50            If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
60               If ObjEsRobable(VictimaIndex, i) Then
70                     If RandomNumber(1, 10) < 4 Then flag = True
80               End If
90            End If
100           If Not flag Then i = i + 1
110       Loop
120   Else
130       i = 20
140       Do While Not flag And i > 0
            'Hay objeto en este slot?
150         If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
160            If ObjEsRobable(VictimaIndex, i) Then
170                  If RandomNumber(1, 10) < 4 Then flag = True
180            End If
190         End If
200         If Not flag Then i = i - 1
210       Loop
220   End If

230   If flag Then
          Dim MiObj As Obj
          Dim Num As Byte
          Dim ObjAmount As Integer
          
240       ObjAmount = UserList(VictimaIndex).Invent.Object(i).Amount
          
          'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
250       Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                      
260       MiObj.Amount = Num
270       MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
          
280       UserList(VictimaIndex).Invent.Object(i).Amount = ObjAmount - Num
                      
290       If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
300             Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
310       End If
                  
320       Call UpdateUserInv(False, VictimaIndex, CByte(i))
                      
330       If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
340           Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
350       End If
          
360       If UserList(LadrOnIndex).clase = eClass.Thief Then
370           Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
380       Else
390           Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
400       End If
410   Else
420       Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
430   End If

      'If exiting, cancel de quien es robado
440   Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)
      '***************************************************
      'Autor: Nacho (Integer) & Unknown (orginal version)
      'Last Modification: 04/17/08 - (NicoNZ)
      'Simplifique la cuenta que hacia para sacar la suerte
      'y arregle la cuenta que hacia para sacar el daño
      '***************************************************
      Dim Suerte As Integer
      Dim Skill As Integer
      Dim pt As Long
10    pt = CalcularDaño(UserIndex, VictimNpcIndex)

20    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

30    Select Case UserList(UserIndex).clase
          Case eClass.Assasin
40            Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
          
50        Case eClass.Cleric, eClass.Paladin, eClass.Pirat
60            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
          
70        Case eClass.Bard
80            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
          
90        Case Else
100           Suerte = Int(0.0361 * Skill + 4.39)
110   End Select


120   If RandomNumber(0, 100) < Suerte Then
130       If VictimUserIndex <> 0 Then
140           If UserList(UserIndex).clase = eClass.Assasin Then
150               daño = Round(daño * 1.4, 0)
160           Else
170               daño = Round(daño * 1.5, 0)
180           End If
190    UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
200    SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, UserList(UserIndex).Dañoapu + daño, DAMAGE_PUÑAL)
210           Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
220           Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
230           Call WriteConsoleMsg(VictimUserIndex, "Su golpe total ha sido de " & Int(UserList(UserIndex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT)
240   Call WriteConsoleMsg(UserIndex, "Tu golpe total es de " & Int(UserList(UserIndex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT)
250           Call FlushBuffer(VictimUserIndex)
260       Else
270           Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
280            SendData SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageCreateDamage(Npclist(VictimNpcIndex).Pos.X, Npclist(VictimNpcIndex).Pos.Y, Int(UserList(UserIndex).Dañoapu + daño), DAMAGE_PUÑAL)
290           Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
300           WriteConsoleMsg UserIndex, "Tu golpe total es de " & Int(UserList(UserIndex).Dañoapu + daño), FontTypeNames.FONTTYPE_FIGHT
              '[Alejo]
310           Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
320   End If
          
330       Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
340   Else
350   Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
360       Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
          'SendData SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, daño, DAMAGE_NORMAL)
370       Call SubirSkill(UserIndex, eSkill.Apuñalar, True)
380   End If

End Sub

Public Sub DoAcuchillar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 12/01/2010
      '***************************************************

10        If UserList(UserIndex).clase <> eClass.Pirat Then Exit Sub
20        If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub

30        If RandomNumber(0, 100) < PROB_ACUCHILLAR Then
40            daño = Int(daño * DAÑO_ACUCHILLAR)
              
50            If VictimUserIndex <> 0 Then
60                UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
70                Call WriteConsoleMsg(UserIndex, "Has acuchillado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
80                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha acuchillado por " & daño, FontTypeNames.FONTTYPE_FIGHT)
90            Else
100               Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
110               Call WriteConsoleMsg(UserIndex, "Has acuchillado a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
120               Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
130           End If
140       End If
          
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
      '***************************************************
      'Autor: Pablo (ToxicWaste)
      'Last Modification: 28/01/2007
      '***************************************************
      Dim Suerte As Integer
      Dim Skill As Integer

10    If UserList(UserIndex).clase <> eClass.Pirat Then Exit Sub
20    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
30    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Sable" Then Exit Sub


40    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

50    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)

60    If RandomNumber(0, 100) < Suerte Then
70        daño = Int(daño * 0.75)
80        If VictimUserIndex <> 0 Then
90            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
100           Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
110           Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
120       Else
130           Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
140           Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
              '[Alejo]
150           Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
160       End If
170   End If

End Sub
Public Sub DoGolpeArco(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Long)
      '***************************************************
      'Autor: Pablo (ToxicWaste)
      'Last Modification: 28/01/2007
      '***************************************************
      Dim Suerte As Integer
      Dim Skill As Integer
      Dim pt As Byte
10    If UserList(UserIndex).clase <> eClass.Hunter Then Exit Sub
20    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
      'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Arco Compuesto Reforzado" Or ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Arco de Cazador" Then Exit Sub
30    pt = RandomNumber(0, 100)


40    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles)
50    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
60    If pt > Suerte Then
70        daño = Int(daño * 0.3)
80       If VictimNpcIndex <> 0 Then
90            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
100           Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
              '[Alejo]
110           Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
120       End If
130       End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

20        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
30        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
40        Call WriteUpdateSta(UserIndex)
          
50    Exit Sub

Errhandler:
60        Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.Description)
          
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer)
10    On Error GoTo Errhandler

      Dim Suerte As Integer
      Dim res As Integer

20    If UserList(UserIndex).clase = eClass.Worker Then
30        Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
40    Else
50        Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
60    End If

      Dim Skill As Integer
70    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.talar)
80    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

90    res = RandomNumber(1, Suerte)

100   If res <= 6 Then
          Dim MiObj As Obj
          
110       If UserList(UserIndex).clase = eClass.Worker And UserList(UserIndex).Invent.WeaponEqpObjIndex = HACHA_LEÑADOR And ObjData(ArbT).ObjType = otarboles Then
120           MiObj.Amount = RandomNumber(8, 18)
130           MiObj.ObjIndex = Leña
140           Else
150            If UserList(UserIndex).clase = eClass.Worker And UserList(UserIndex).Invent.WeaponEqpObjIndex = HACHA_DORADA And ObjData(ArbT).ObjType = 38 Then
160           MiObj.Amount = RandomNumber(5, 13)
170           MiObj.ObjIndex = 642
180       Else
190       MiObj.ObjIndex = Leña
200           MiObj.Amount = 1
210       End If
220       End If
          
          
230       If Not MeterItemEnInventario(UserIndex, MiObj) Then
              
240           Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
              
250       End If
          
260       Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
          
270   Else
          '[CDT 17-02-2004]
280       If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
290           Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
300           UserList(UserIndex).flags.UltimoMensaje = 8
310       End If
          '[/CDT]
320   End If

330   Call SubirSkill(UserIndex, talar, True)

340   UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
350   If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
          UserList(UserIndex).Reputacion.PlebeRep = MAXREP

360   UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

370   Exit Sub

Errhandler:
380       Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
10    On Error GoTo Errhandler

      Dim Suerte As Integer
      Dim res As Integer

20    If UserList(UserIndex).clase = eClass.Worker Then
30        Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
40    Else
50        Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
60    End If

      Dim Skill As Integer
70    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Mineria)
80    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

90    res = RandomNumber(1, Suerte)

100   If res <= 5 Then
          Dim MiObj As Obj
          
110       If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
          
120       MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
          
130       If UserList(UserIndex).clase = eClass.Worker Then
140           MiObj.Amount = RandomNumber(5, 13) '(NicoNZ) 04/25/2008
150       Else
160           MiObj.Amount = 1
170       End If
          
180       If Not MeterItemEnInventario(UserIndex, MiObj) Then _
              Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
          
190       Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
          
200   Else
          '[CDT 17-02-2004]
210       If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
220           Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
230           UserList(UserIndex).flags.UltimoMensaje = 9
240       End If
          '[/CDT]
250   End If

260   Call SubirSkill(UserIndex, Mineria, True)

270   UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
280   If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
          UserList(UserIndex).Reputacion.PlebeRep = MAXREP

290   UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

300   Exit Sub

Errhandler:
310       Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With UserList(UserIndex)
20            .Counters.IdleCount = 0
              
              Dim Suerte As Integer
              Dim res As Integer
              Dim cant As Integer
              Dim MeditarSkill As Byte

30            If .Stats.MinMAN >= .Stats.MaxMAN Then
40                Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
50                Call WriteMeditateToggle(UserIndex)
60                .flags.Meditando = False
70                .Char.FX = 0
80                .Char.loops = 0
90                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
100               Exit Sub
110           End If
              
120   MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
              
130           If MeditarSkill <= 10 And MeditarSkill >= -1 Then
140               Suerte = 35
150           ElseIf MeditarSkill <= 30 And MeditarSkill >= 11 Then
160               Suerte = 30
170           ElseIf MeditarSkill <= 40 And MeditarSkill >= 21 Then
180               Suerte = 28
190           ElseIf MeditarSkill <= 50 And MeditarSkill >= 31 Then
200               Suerte = 24
210           ElseIf MeditarSkill <= 60 And MeditarSkill >= 41 Then
220               Suerte = 22
230           ElseIf MeditarSkill <= 70 And MeditarSkill >= 51 Then
240               Suerte = 20
250           ElseIf MeditarSkill <= 80 And MeditarSkill >= 61 Then
260               Suerte = 18
270           ElseIf MeditarSkill <= 90 And MeditarSkill >= 71 Then
280               Suerte = 15
290           ElseIf MeditarSkill <= 100 And MeditarSkill >= 81 Then
300               Suerte = 10
310           ElseIf MeditarSkill < 110 And MeditarSkill >= 91 Then
320               Suerte = 7
330           ElseIf MeditarSkill = 100 Then
340               Suerte = 5
350           End If
360           res = RandomNumber(1, Suerte)
              
370           If res = 1 Then
                  
380               cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)
390               If cant <= 0 Then cant = 1
400               .Stats.MinMAN = .Stats.MinMAN + cant
410               If .Stats.MinMAN > .Stats.MaxMAN Then _
                      .Stats.MinMAN = .Stats.MaxMAN
                  
420                   Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de maná!", FontTypeNames.FONTTYPE_INFO)

                  
430               Call WriteUpdateMana(UserIndex)
440               Call WriteUpdateFollow(UserIndex)
450               Call SubirSkill(UserIndex, eSkill.Meditar, True)
460           Else
470               Call SubirSkill(UserIndex, eSkill.Meditar, False)
480           End If
490       End With
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modif: 15/04/2010
      'Unequips either shield, weapon or helmet from target user.
      '***************************************************

          Dim Probabilidad As Integer
          Dim Resultado As Integer
          Dim WrestlingSkill As Byte
          Dim AlgoEquipado As Boolean
          
10        With UserList(UserIndex)
              ' Si no tiene guantes de hurto no desequipa.
20            If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
              
              ' Si no esta solo con manos, no desequipa tampoco.
30            If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
              
40            WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
              
50            Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
60       End With
         
70       With UserList(VictimIndex)
              ' Si tiene escudo, intenta desequiparlo
80            If .Invent.EscudoEqpObjIndex > 0 Then
                  
90                Resultado = RandomNumber(1, 100)
                  
100               If Resultado <= Probabilidad Then
                      ' Se lo desequipo
110                   Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                      
120                   Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                      
130                   If .Stats.ELV < 20 Then
140                       Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
150                   End If
                      
160                   Call FlushBuffer(VictimIndex)
                      
170                   Exit Sub
180               End If
                  
190               AlgoEquipado = True
200           End If
              
              ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
210           If .Invent.WeaponEqpObjIndex > 0 Then
                  
220               Resultado = RandomNumber(1, 100)
                  
230               If Resultado <= Probabilidad Then
                      ' Se lo desequipo
240                   Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                      
250                   Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                      
260                   If .Stats.ELV < 20 Then
270                       Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
280                   End If
                      
290                   Call FlushBuffer(VictimIndex)
                      
300                   Exit Sub
310               End If
                  
320               AlgoEquipado = True
330           End If
              
              ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
340           If .Invent.CascoEqpObjIndex > 0 Then
                  
350               Resultado = RandomNumber(1, 100)
                  
360               If Resultado <= Probabilidad Then
                      ' Se lo desequipo
370                   Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                      
380                   Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                      
390                   If .Stats.ELV < 20 Then
400                       Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
410                   End If
                      
420                   Call FlushBuffer(VictimIndex)
                      
430                   Exit Sub
440               End If
                  
450               AlgoEquipado = True
460           End If
          
470           If AlgoEquipado Then
480               Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
490           Else
500               Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
510           End If
          
520       End With


End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modif: 03/03/2010
      'Implements the pick pocket skill of the Bandit :)
      '03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
      '***************************************************
      Dim OtroUserIndex As Integer

10    If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

      'If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub
      'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
      'Uso el slot de los anillos para "equipar" los guantes.
      'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
20    If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

      Dim res As Integer
30    res = RandomNumber(1, 100)
40    If (res < 20) Then
50        If TieneObjetosRobables(VictimaIndex) Then
          
60            If UserList(VictimaIndex).flags.Comerciando Then
70                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                      
80                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
90                    Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
100                   Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                      
110                   Call LimpiarComercioSeguro(VictimaIndex)
120                   Call Protocol.FlushBuffer(OtroUserIndex)
130               End If
140           End If
                      
150           Call RobarObjeto(UserIndex, VictimaIndex)
160           Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
170       Else
180           Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
190       End If
200   End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste)
      'Last Modif: 17/02/2007
      'Implements the special Skill of the Thief
      '***************************************************
10    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
20    If UserList(UserIndex).clase <> eClass.Thief Then Exit Sub
          

30    If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
          
      Dim res As Integer
40    res = RandomNumber(0, 100)
50    If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
60        UserList(VictimaIndex).flags.Paralizado = 1
70        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
80        Call WriteParalizeOK(VictimaIndex)
90        Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
100       Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
110   End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 02/04/2010 (ZaMa)
      '02/04/2010: ZaMa - Nueva formula para desarmar.
      '***************************************************

          Dim Probabilidad As Integer
          Dim Resultado As Integer
          Dim WrestlingSkill As Byte
          
10        With UserList(UserIndex)
20            WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
              
30            Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
              
40            Resultado = RandomNumber(1, 100)
              
50            If Resultado <= Probabilidad Then
60                Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
70                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
80                If UserList(VictimIndex).Stats.ELV < 20 Then
90                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
100               End If
110               Call FlushBuffer(VictimIndex)
120           End If
130       End With
          
End Sub
Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
      '***************************************************
      'Author: ZaMa
      'Last Modification: 29/01/2010
      '
      '***************************************************
10        MaxItemsConstruibles = MaximoInt(1, CInt((UserList(UserIndex).Stats.ELV - 4) / 5))
End Function
Public Sub ImitateNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 20/11/2010
      'Copies body, head and desc from previously clicked npc.
      '***************************************************
          
10        With UserList(UserIndex)
              
              ' Copy desc
20            .DescRM = Npclist(NpcIndex).Name

              ' Remove Anims (Npcs don't use equipment anims yet)
30            .Char.CascoAnim = NingunCasco
40            .Char.ShieldAnim = NingunEscudo
50            .Char.WeaponAnim = NingunArma
              
              ' If admin is invisible the store it in old char
60            If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
                  
70                .flags.OldBody = Npclist(NpcIndex).Char.body
80                .flags.OldHead = Npclist(NpcIndex).Char.Head
90            Else
100               .Char.body = Npclist(NpcIndex).Char.body
110               .Char.Head = Npclist(NpcIndex).Char.Head
120               Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
130           End If
          
140       End With
          
End Sub
Public Sub DoEquita(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)
       
      Dim ModEqui As Long
10
        ModEqui = ModEquitacion(UserList(UserIndex).clase)
        
20     With UserList(UserIndex)
30       If .Stats.UserSkills(eSkill.Equitacion) / ModEqui < Montura.MinSkill Then
40           Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill * ModEqui & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
50           Exit Sub
60      End If


70    .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
80    .Invent.MonturaSlot = Slot
       
90         If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
100           Exit Sub
110       End If
       
120      If .flags.Montando = 0 Then
130          .Char.Head = 0
140          If .flags.Muerto = 0 Then
150              .Char.body = Montura.Ropaje
160          Else
170              .Char.body = iCuerpoMuerto
180              .Char.Head = iCabezaMuerto
190          End If
200          .Char.Head = UserList(UserIndex).OrigChar.Head
210          .Char.ShieldAnim = NingunEscudo
220          .Char.WeaponAnim = NingunArma
230          .Char.CascoAnim = .Char.CascoAnim
240          .flags.Montando = 1
250      Else
260        .flags.Montando = 0
270          If .flags.Muerto = 0 Then
280             .Char.Head = UserList(UserIndex).OrigChar.Head
290              If .Invent.ArmourEqpObjIndex > 0 Then
300                 .Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
310              Else
320                  Call DarCuerpoDesnudo(UserIndex)
330              End If
340                    If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
350                    If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
360                    If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
370            Else
                  'Ds AO
380               If UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then
390                   UserList(UserIndex).Char.body = iCuerpoMuertoCrimi
400                   UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
410                   UserList(UserIndex).Char.ShieldAnim = NingunEscudo
420                   UserList(UserIndex).Char.WeaponAnim = NingunArma
430                 UserList(UserIndex).Char.CascoAnim = NingunCasco
440               Else
450                   UserList(UserIndex).Char.body = iCuerpoMuerto
460                   UserList(UserIndex).Char.Head = iCabezaMuerto
470                   UserList(UserIndex).Char.ShieldAnim = NingunEscudo
480                   UserList(UserIndex).Char.WeaponAnim = NingunArma
490                   UserList(UserIndex).Char.CascoAnim = NingunCasco
500               End If
510        End If
520    End If
       
       
530    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
540    Call WriteMontateToggle(UserIndex)
550    End With
End Sub

Function ModEquitacion(ByVal clase As String) As Integer
10    Select Case UCase$(clase)
          Case "1"
20            ModEquitacion = 1
30        Case Else
40            ModEquitacion = 1
50    End Select
       
End Function
