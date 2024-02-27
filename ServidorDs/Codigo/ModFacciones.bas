Attribute VB_Name = "ModFacciones"
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

Public Gld As Long
Public ArmaduraImperial1 As Integer
Public ArmaduraImperial2 As Integer
Public ArmaduraImperial3 As Integer
Public TunicaMagoImperial As Integer
Public TunicaMagoImperialEnanos As Integer
Public ArmaduraCaos1 As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer



Public Const NUM_RANGOS_FACCION As Integer = 5
Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3

Public Enum eTipoDefArmors
    ieBaja
    ieMedia
    ieAlta
End Enum

Public Type tFaccionArmaduras
    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer
End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION) As Long

Private Function GetArmourAmount(ByVal Rango As Integer, ByVal TipoDef As eTipoDefArmors) As Integer
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 15/04/2010
      'Returns the amount of armours to give, depending on the specified rank
      '***************************************************

10        Select Case TipoDef
              
              Case eTipoDefArmors.ieBaja
20                GetArmourAmount = 1
                  
30            Case eTipoDefArmors.ieMedia
40                GetArmourAmount = 1
                  
50            Case eTipoDefArmors.ieAlta
60                GetArmourAmount = 1
                  
70        End Select
          
End Function

Private Sub GiveFactionArmours(ByVal Userindex As Integer, ByVal IsCaos As Boolean)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 15/04/2010
      'Gives faction armours to user
      '***************************************************
          
          Dim ObjArmour As Obj
          Dim Rango As Integer
          
10        With UserList(Userindex)
          
20            Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1
          
          
              ' Entrego armaduras de defensa baja
              'ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)
              'If IsCaos = True And ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta And eTipoDefArmors.ieMedia And eTipoDefArmors.ieBaja) >= 1 Then Exit Sub
              'If IsCaos = False And ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta And eTipoDefArmors.ieMedia And eTipoDefArmors.ieBaja) >= 1 Then Exit Sub
             ' If IsCaos Then
             '     ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
            '  Else
             '     ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)
             ' End If
              
             ' If Not MeterItemEnInventario(userIndex, ObjArmour) Then
             '     Call TirarItemAlPiso(.Pos, ObjArmour)
             ' End If
              
              
              ' Entrego armaduras de defensa media
30            ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)
40            If IsCaos = True Then
50          If .Faccion.RecibioArmaduraCaos = 0 Then
       Dim MiObj As Obj
60        MiObj.Amount = 1
         
         'CAOS
70        Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
80                Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
90                        MiObj.objindex = 734
100                   Case eClass.Cleric
110                       MiObj.objindex = 736
120                   Case eClass.Paladin, eClass.Warrior
130                       MiObj.objindex = 738
140                   Case eClass.Mage
150                       Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
160                               MiObj.objindex = 740
170                           Case eGenero.Hombre
180                               MiObj.objindex = 741
190                       End Select
200                 End Select
210           Case eRaza.Gnomo, eRaza.Enano
220               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
230                       MiObj.objindex = 735
240                   Case eClass.Cleric
250                       MiObj.objindex = 737
260                   Case eClass.Paladin, eClass.Warrior
270                       MiObj.objindex = 739
280                   Case eClass.Mage
290                       MiObj.objindex = 742
300           End Select
310       End Select
          
320       If Not MeterItemEnInventario(Userindex, MiObj) Then
330               Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
340      End If
350       End If
360       UserList(Userindex).Faccion.RecibioArmaduraCaos = 1
370       UserList(Userindex).Faccion.NivelIngreso = UserList(Userindex).Stats.ELV
380       UserList(Userindex).Faccion.FechaIngreso = Date
390   ElseIf IsCaos = False Then
400   If UserList(Userindex).Faccion.RecibioArmaduraReal = 0 Then
410       MiObj.Amount = 1
              
          'ARMADA
420       Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
430               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
440                       MiObj.objindex = 779
450                   Case eClass.Cleric
460                       MiObj.objindex = 781
470                   Case eClass.Paladin, eClass.Warrior
480                       MiObj.objindex = 783
490                   Case eClass.Mage
500                       Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
510                               MiObj.objindex = 785
520                           Case eGenero.Hombre
530                               MiObj.objindex = 786
540                       End Select
550                 End Select
560           Case eRaza.Gnomo, eRaza.Enano
570               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
580                       MiObj.objindex = 780
590                   Case eClass.Cleric
600                       MiObj.objindex = 782
610                   Case eClass.Paladin, eClass.Warrior
620                       MiObj.objindex = 784
630                   Case eClass.Mage
640                       MiObj.objindex = 787
650           End Select
660       End Select
          
670       If Not MeterItemEnInventario(Userindex, MiObj) Then
680               Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
690       End If
700       End If
710       UserList(Userindex).Faccion.RecibioArmaduraReal = 1
720       UserList(Userindex).Faccion.NivelIngreso = UserList(Userindex).Stats.ELV
730       UserList(Userindex).Faccion.FechaIngreso = Date
          'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
740       UserList(Userindex).Faccion.MatadosIngreso = UserList(Userindex).Faccion.CiudadanosMatados

750   End If

760       End With

End Sub

Public Sub GiveExpReward(ByVal Userindex As Integer, ByVal Rango As Long)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 15/04/2010
      'Gives reward exp to user
      '***************************************************
          
          Dim GivenExp As Long
          
10        With UserList(Userindex)
              
20            GivenExp = RecompensaFacciones(Rango)
              
30            .Stats.Exp = .Stats.Exp + GivenExp
              
40            If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
              
50            Call WriteConsoleMsg(Userindex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

60            Call CheckUserLevel(Userindex)
              
70        End With
          
End Sub

Public Sub EnlistarArmadaReal(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 15/04/2010
      'Handles the entrance of users to the "Armada Real"
      '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
      '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
      '15/04/2010: ZaMa - Cambio en recompensas iniciales.
      '***************************************************

10    With UserList(Userindex)
20        If .Faccion.ArmadaReal = 1 Then
30            Call WriteChatOverHead(Userindex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
40            Exit Sub
50        End If
          
60        If .Faccion.FuerzasCaos = 1 Then
70            Call WriteChatOverHead(Userindex, "¡¡¡Maldito insolente!!! Vete de aquí seguidor de las sombras.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
80            Exit Sub
90        End If
          
100       If criminal(Userindex) Then
110           Call WriteChatOverHead(Userindex, "¡¡¡No se permiten criminales en el ejército real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
120           Exit Sub
130       End If
          
140       If .Faccion.CriminalesMatados < 50 Then
150           Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 50 criminales, sólo has matado " & .Faccion.CriminalesMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
160           Exit Sub
170       End If
          
180       If .Stats.ELV < 25 Then
190           Call WriteChatOverHead(Userindex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
200           Exit Sub
210       End If
           
220       If .Faccion.CiudadanosMatados > 0 Then
230           Call WriteChatOverHead(Userindex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
240           Exit Sub
250       End If
          
260       If .Faccion.Reenlistadas > 4 Then
270           Call WriteChatOverHead(Userindex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
280           Exit Sub
290       End If
          
300       If .Reputacion.NobleRep < 0 Then
310           Call WriteChatOverHead(Userindex, "Necesitas ser aún más noble para integrar el ejército real, sólo tienes " & .Reputacion.NobleRep & "/20.000 puntos de nobleza", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
320           Exit Sub
330       End If
          
340       If .GuildIndex > 0 Then
350           If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
360               Call WriteChatOverHead(Userindex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
370               Exit Sub
380           End If
390       End If
          
400       .Faccion.ArmadaReal = 1
410       .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
          
420       Call WriteChatOverHead(Userindex, "¡¡¡Bienvenido al ejército real!!! Aquí tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargaré de recompensarte.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
430       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Rey de Banderbill> Ahora le ofreceré estas vestimentas a " & .Name & " por haberse enlistado a la Armada Real. Espero grandes logros de este noble guerrero.", FontTypeNames.FONTTYPE_CONSEJOVesA))
          
          ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
440       If .Faccion.RecibioArmaduraReal = 0 Then
              
             
          Dim LiObj As Obj
450       LiObj.Amount = 1
          
      '[Wizard 03/09/05] no se quien hizo lo que estaba aca, pero por dios mandenlo a un curso de redaccion
      'Habia 3 cases diciendo lo mismo, 1 If clause que nunca se accedia por suerte porque si se accedia daba armadura del caos
      'ademas usan los Ucase$ para esto, que son cosas que los escribe el codigo y no pueden cambiar, gastan memoria ram al pedo.
460   Select Case .raza
          Case Drow, Elfo, Humano
470           If .clase = Cleric Or .clase = Druid Or .clase = Bard Then
480               LiObj.objindex = 372
490           ElseIf .Genero = Hombre And .clase = Mage Then
500               LiObj.objindex = 517
510           ElseIf .Genero = Mujer And .clase = Mage Then
520               LiObj.objindex = 516
530           ElseIf (.Genero = Mujer) And (.clase = Paladin Or .clase = Warrior Or .clase = Assasin Or .clase = Hunter) Then
540               LiObj.objindex = 520
550           ElseIf (.Genero = Hombre) And (.clase = Paladin Or .clase = Warrior Or .clase = Assasin Or .clase = Hunter) Then
560               LiObj.objindex = 521
570           End If
          
580       Case Gnomo, Enano
590           If .clase = Warrior Or .clase = Paladin Or .clase = Hunter Or .clase = Assasin Then
600               LiObj.objindex = 492
610           ElseIf .clase = Mage Or .clase = Bard Or .clase = Druid Or .clase = Cleric Then
620               LiObj.objindex = 549
630           Else 'Trabajadoras
640               LiObj.objindex = 678
650           End If
660   End Select
              
670           If Not MeterItemEnInventario(Userindex, LiObj) Then
680               Call TirarItemAlPiso(.Pos, LiObj)
690           End If
700    .Faccion.RecibioArmaduraReal = 1
       
710           Call GiveExpReward(Userindex, 0)
              
720           .Faccion.RecibioArmaduraReal = 1
730           .Faccion.NivelIngreso = .Stats.ELV
740           .Faccion.FechaIngreso = Date
              'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
750           .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados
              
760           .Faccion.RecibioExpInicialReal = 1
770           .Faccion.RecompensasReal = 0
780           .Faccion.NextRecompensa = 100
              
790       End If
          
800       If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
          
810       Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
820   End With

End Sub

Public Sub RecompensaArmadaReal(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 23/01/2007
      'Handles the way of gaining new ranks in the "Armada Real"
      '***************************************************
      Dim Crimis As Long
      Dim Lvl As Byte
      Dim NextRecom As Long
      Dim Nobleza As Long
      Dim MiObj As Obj
10    MiObj.Amount = 1
20    Lvl = UserList(Userindex).Stats.ELV
30    Crimis = UserList(Userindex).Faccion.CriminalesMatados
40    NextRecom = UserList(Userindex).Faccion.NextRecompensa
50    Nobleza = UserList(Userindex).Reputacion.NobleRep

60    If Crimis < NextRecom Then
70        Call WriteChatOverHead(Userindex, "Ya has recibido tu recompensa, mata " & NextRecom - Crimis & " criminaales mas para recibir la proxima!!!", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
80        Exit Sub
90    End If

100   Select Case NextRecom
              Case 100:
110            If Lvl < 32 Then
120                   Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 32 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
130                   Exit Sub
140               End If
150               If (UserList(Userindex).Stats.Gld >= 500000) Then
160           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 500000
170           Call WriteUpdateGold(Userindex)
180           Else
190           Call WriteChatOverHead(Userindex, "Necesitas 500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
200           Exit Sub
210           End If
220               UserList(Userindex).Faccion.RecompensasReal = 1
230               UserList(Userindex).Faccion.NextRecompensa = 175
                  'Call PerderItemsFaccionarios(Userindex)
                  '2doARMY
240       Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
250               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
260                       MiObj.objindex = 788
270                   Case eClass.Cleric
280                       MiObj.objindex = 790
290                   Case eClass.Paladin, eClass.Warrior
300                       MiObj.objindex = 792
310                   Case eClass.Mage
320                       Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
330                               MiObj.objindex = 794
340                           Case eGenero.Hombre
350                               MiObj.objindex = 795
360                       End Select
370                 End Select
380           Case eRaza.Gnomo, eRaza.Enano
390               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
400                       MiObj.objindex = 789
410                   Case eClass.Cleric
420                       MiObj.objindex = 791
430                   Case eClass.Paladin, eClass.Warrior
440                       MiObj.objindex = 793
450                   Case eClass.Mage
460                       MiObj.objindex = 796
470           End Select
480       End Select
              
490           Case 175:
500            If Lvl < 36 Then
510                   Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
520                   Exit Sub
530               End If
540               If (UserList(Userindex).Stats.Gld >= 1500000) Then
550           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 1500000
560           Call WriteUpdateGold(Userindex)
570           Else
580           Call WriteChatOverHead(Userindex, "Necesitas 1.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
590           Exit Sub
600           End If
610               UserList(Userindex).Faccion.RecompensasReal = 2
620               UserList(Userindex).Faccion.NextRecompensa = 250
                  'Call PerderItemsFaccionarios(Userindex)
                  '3cerARMY
630   Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
640               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
650                       MiObj.objindex = 797
660                   Case eClass.Cleric
670                       MiObj.objindex = 799
680                   Case eClass.Paladin, eClass.Warrior
690                       MiObj.objindex = 801
700                   Case eClass.Mage
710                       Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
720                               MiObj.objindex = 803
730                           Case eGenero.Hombre
740                               MiObj.objindex = 804
750                       End Select
760                 End Select
770           Case eRaza.Gnomo, eRaza.Enano
780               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
790                       MiObj.objindex = 798
800                   Case eClass.Cleric
810                       MiObj.objindex = 800
820                   Case eClass.Paladin, eClass.Warrior
830                       MiObj.objindex = 802
840                   Case eClass.Mage
850                       MiObj.objindex = 805
860           End Select
870       End Select
              
880           Case 250:
890            If Lvl < 38 Then
900                   Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 38 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
910                   Exit Sub
920               End If
930               If (UserList(Userindex).Stats.Gld >= 2500000) Then
940           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 2500000
950           Call WriteUpdateGold(Userindex)
960           Else
970           Call WriteChatOverHead(Userindex, "Necesitas 2.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
980           Exit Sub
990           End If
1000              UserList(Userindex).Faccion.RecompensasReal = 3
1010              UserList(Userindex).Faccion.NextRecompensa = 325
                  'Call PerderItemsFaccionarios(Userindex)
                  
      '4toARMY
1020  Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
1030              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1040                      MiObj.objindex = 806
1050                  Case eClass.Cleric
1060                      MiObj.objindex = 808
1070                  Case eClass.Paladin, eClass.Warrior
1080                      MiObj.objindex = 810
1090                  Case eClass.Mage
1100                      Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
1110                              MiObj.objindex = 812
1120                          Case eGenero.Hombre
1130                              MiObj.objindex = 813
1140                      End Select
1150                End Select
1160          Case eRaza.Gnomo, eRaza.Enano
1170              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1180                      MiObj.objindex = 807
1190                  Case eClass.Cleric
1200                      MiObj.objindex = 809
1210                  Case eClass.Paladin, eClass.Warrior
1220                      MiObj.objindex = 811
1230                  Case eClass.Mage
1240                      MiObj.objindex = 814
1250          End Select
1260      End Select
              
1270          Case 325:
1280           If Lvl < 40 Then
1290                  Call WriteChatOverHead(Userindex, "Mataste suficientes criminales, pero te faltan " & 40 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
1300                  Exit Sub
1310              End If
1320              UserList(Userindex).Faccion.RecompensasReal = 4
1330              UserList(Userindex).Faccion.NextRecompensa = 415
                 ' Call PerderItemsFaccionarios(Userindex)
1340              If (UserList(Userindex).Stats.Gld >= 4000000) Then
1350          UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 4000000
1360          Call WriteUpdateGold(Userindex)
1370          Else
1380          Call WriteChatOverHead(Userindex, "Necesitas 4.000.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
1390          Exit Sub
1400          End If
                  '5toARMY
1410  Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
1420              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1430                      MiObj.objindex = 815
1440                  Case eClass.Cleric
1450                      MiObj.objindex = 817
1460                  Case eClass.Paladin, eClass.Warrior
1470                      MiObj.objindex = 819
1480                  Case eClass.Mage
1490                      Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
1500                              MiObj.objindex = 821
1510                          Case eGenero.Hombre
1520                              MiObj.objindex = 822
1530                      End Select
1540                End Select
1550          Case eRaza.Gnomo, eRaza.Enano
1560              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1570                      MiObj.objindex = 816
1580                  Case eClass.Cleric
1590                      MiObj.objindex = 818
1600                  Case eClass.Paladin, eClass.Warrior
1610                      MiObj.objindex = 820
1620                  Case eClass.Mage
1630                      MiObj.objindex = 823
1640          End Select
1650      End Select
                  '.Faccion.NextRecompensa = 17001
            
              
1660          Case 415:
1670              Exit Sub
1680      End Select
1690      If Not MeterItemEnInventario(Userindex, MiObj) Then
1700              Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
1710     End If
       
1720  Call WriteChatOverHead(Userindex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(Userindex) + "!!!", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
      'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
      'If UserList(UserIndex).Stats.Exp > MAXEXP Then
      '    UserList(UserIndex).Stats.Exp = MAXEXP
      'End If
      'Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

      'Call CheckUserLevel(UserIndex)


End Sub

Public Sub ExpulsarFaccionReal(ByVal Userindex As Integer, Optional Expulsado As Boolean = True)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    With UserList(Userindex)
20        .Faccion.ArmadaReal = 0
30        Call PerderItemsFaccionarios(Userindex)
40        If Expulsado Then
50            Call WriteConsoleMsg(Userindex, "¡¡¡Has sido expulsado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
60        Else
70            Call WriteConsoleMsg(Userindex, "¡¡¡Te has retirado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
80            .Faccion.RecibioArmaduraReal = 0
90        End If
          
100       If .Invent.ArmourEqpObjIndex <> 0 Then
              'Desequipamos la armadura real si está equipada
110           If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpObjIndex)
         'Call QuitarObjetos(.Invent.ArmourEqpObjIndex, 1, UserIndex)
120       End If
          
130       If .Invent.EscudoEqpObjIndex <> 0 Then
              'Desequipamos el escudo de caos si está equipado
140           If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpObjIndex)
          'Call QuitarObjetos(ObjData(.Invent.EscudoEqpObjIndex).Real = 1, 1, UserIndex)
150       End If
          
160       If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
170   End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal Userindex As Integer, Optional Expulsado As Boolean = True)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    With UserList(Userindex)
20        .Faccion.FuerzasCaos = 0
30        Call PerderItemsFaccionarios(Userindex)
40        If Expulsado Then
50            Call WriteConsoleMsg(Userindex, "¡¡¡Has sido expulsado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
60        Else
70            Call WriteConsoleMsg(Userindex, "¡¡¡Te has retirado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
80        End If
          
90        If .Invent.ArmourEqpObjIndex <> 0 Then
              'Desequipamos la armadura de caos si está equipada
100           If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
         ' Call QuitarObjetos(.Invent.ArmourEqpObjIndex, 1, UserIndex)
110       End If
          
120       If .Invent.EscudoEqpObjIndex <> 0 Then
              'Desequipamos el escudo de caos si está equipado
130           If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpObjIndex)
          'Call QuitarObjetos(.Invent.EscudoEqpObjIndex, 1, UserIndex)
140       End If
          
          
          
          
150       If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
160   End With

End Sub

Public Function TituloReal(ByVal Userindex As Integer) As String
      '***************************************************
      'Autor: Unknown
      'Last Modification: 23/01/2007 Pablo (ToxicWaste)
      'Handles the titles of the members of the "Armada Real"
      '***************************************************

10    Select Case UserList(Userindex).Faccion.RecompensasReal
         
          Case 0
20            TituloReal = "Aprendiz"
30        Case 1
40            TituloReal = "Caballero"
50        Case 2
60            TituloReal = "Capitán"
70        Case 3
80            TituloReal = "Guardián"
90        Case Else
100           TituloReal = "Campeón de la Luz"
110   End Select


End Function

Public Sub EnlistarCaos(ByVal Userindex As Integer)
      '***************************************************
      'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 27/11/2009
      '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
      '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
      'Handles the entrance of users to the "Legión Oscura"
      '***************************************************

10    With UserList(Userindex)
20        If Not criminal(Userindex) Then
30            Call WriteChatOverHead(Userindex, "¡¡¡Lárgate de aquí, bufón!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
40            Exit Sub
50        End If
          
60        If .Faccion.FuerzasCaos = 1 Then
70            Call WriteChatOverHead(Userindex, "¡¡¡Ya perteneces a la legión oscura!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
80            Exit Sub
90        End If
          
100       If .Faccion.ArmadaReal = 1 Then
110           Call WriteChatOverHead(Userindex, "Las sombras reinarán en las tierras Desterianas. ¡¡¡Fuera de aquí insecto real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
120           Exit Sub
130       End If
          
          '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
140       If .Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
150           Call WriteChatOverHead(Userindex, "No permitiré que ningún insecto real ingrese a mis tropas.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
160           Exit Sub
170       End If
          '[/Barrin]
          
180       If Not criminal(Userindex) Then
190           Call WriteChatOverHead(Userindex, "¡¡Ja ja ja!! Tú no eres bienvenido aquí asqueroso ciudadano.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
200           Exit Sub
210       End If
          
220      If .Faccion.CiudadanosMatados < 60 Then
230           Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 60 ciudadanos, sólo has matado " & .Faccion.CiudadanosMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
240           Exit Sub
250       End If
          
260       If .Stats.ELV < 25 Then
270           Call WriteChatOverHead(Userindex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
280           Exit Sub
290       End If
          
300       If .GuildIndex > 0 Then
310           If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
320               Call WriteChatOverHead(Userindex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
330               Exit Sub
340           End If
350       End If
          
          
360       If .Faccion.Reenlistadas > 4 Then
370           If .Faccion.Reenlistadas = 200 Then
380               Call WriteChatOverHead(Userindex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
390           Else
400               Call WriteChatOverHead(Userindex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
410           End If
420           Exit Sub
430       End If
          
440       .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
450       .Faccion.FuerzasCaos = 1
          
460       Call WriteChatOverHead(Userindex, "¡¡¡Bienvenido al lado oscuro!!! Aquí tienes tus armaduras. Derrama sangre ciudadana y real, y serás recompensado, lo prometo.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
470        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Señor del Miedo> Es un gusto anunciar que " & .Name & " se ha enlistado en la Legión Oscura. Su alma me pertenece y su equipamiento es la recompensa para sembrar el miedo en estas tierras.", FontTypeNames.FONTTYPE_CONSEJOCAOSVesA))
          
480       If .Faccion.RecibioArmaduraCaos = 0 Then
                      
490           Call GiveFactionArmours(Userindex, True)
500           Call GiveExpReward(Userindex, 0)
              
510           .Faccion.RecibioArmaduraCaos = 1
520           .Faccion.NivelIngreso = .Stats.ELV
530           .Faccion.FechaIngreso = Date
          
540           .Faccion.RecibioExpInicialCaos = 1
550           .Faccion.RecompensasCaos = 0
560           .Faccion.NextRecompensa = 110
570       End If
          
580       If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

590       Call LogEjercitoCaos(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
600   End With

End Sub

Public Sub RecompensaCaos(ByVal Userindex As Integer)
      '***************************************************
      'Author: Pablo (ToxicWaste) & Unknown (orginal version)
      'Last Modification: 23/01/2007
      'Handles the way of gaining new ranks in the "Legión Oscura"
      '***************************************************
      Dim Ciudas As Long
      Dim Lvl As Byte
      Dim NextRecom As Long
      Dim MiObj As Obj
10    MiObj.Amount = 1
20    Lvl = UserList(Userindex).Stats.ELV
30    Ciudas = UserList(Userindex).Faccion.CiudadanosMatados
40    NextRecom = UserList(Userindex).Faccion.NextRecompensa

50    If Ciudas < NextRecom Then
60        Call WriteChatOverHead(Userindex, "Ya has recibido tu recompensa, mata " & NextRecom - Ciudas & "  ciudadanos mas para recibir la proxima!!!", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
70        Exit Sub
80    End If

90    Select Case NextRecom
          Case 110:
100           If Lvl < 27 Then
110                   Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
120                   Exit Sub
130               End If
140           If (UserList(Userindex).Stats.Gld >= 500000) Then
150           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 500000
160           Call WriteUpdateGold(Userindex)
170           Else
180           Call WriteChatOverHead(Userindex, "Necesitas 500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
190           Exit Sub
200           End If
210               UserList(Userindex).Faccion.RecompensasCaos = 1
220               UserList(Userindex).Faccion.NextRecompensa = 180
                 ' Call PerderItemsFaccionarios(Userindex)
                  '2doCAOS
230   Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
240               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
250                       MiObj.objindex = 743
260                   Case eClass.Cleric
270                       MiObj.objindex = 745
280                   Case eClass.Paladin, eClass.Warrior
290                       MiObj.objindex = 747
300                   Case eClass.Mage
310                       If UserList(Userindex).Genero = 2 Then
320                       MiObj.objindex = 749
330                       ElseIf UserList(Userindex).Genero = 1 Then
340                       MiObj.objindex = 750
350                       End If
360                 End Select
370           Case eRaza.Gnomo, eRaza.Enano
380               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
390                       MiObj.objindex = 744
400                   Case eClass.Cleric
410                       MiObj.objindex = 746
420                   Case eClass.Paladin, eClass.Warrior
430                       MiObj.objindex = 748
440                   Case eClass.Mage
450                       MiObj.objindex = 751
460           End Select
470       End Select
              
480           Case 180:
490           If Lvl < 30 Then
500                   Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
510                   Exit Sub
520               End If
530               If (UserList(Userindex).Stats.Gld >= 1500000) Then
540           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 1500000
550           Call WriteUpdateGold(Userindex)
560           Else
570           Call WriteChatOverHead(Userindex, "Necesitas 1.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
580           Exit Sub
590           End If
600               UserList(Userindex).Faccion.RecompensasCaos = 2
610               UserList(Userindex).Faccion.NextRecompensa = 270
                '  Call PerderItemsFaccionarios(Userindex)
                  '3cerCAOS
620   Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
630               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
640                       MiObj.objindex = 752
650                   Case eClass.Cleric
660                       MiObj.objindex = 754
670                   Case eClass.Paladin, eClass.Warrior
680                       MiObj.objindex = 756
690                   Case eClass.Mage
700                       If UserList(Userindex).Genero = 1 Then
710                       MiObj.objindex = 758
720                       ElseIf UserList(Userindex).Genero = 2 Then
730                       MiObj.objindex = 759
740                       End If
750                 End Select
760           Case eRaza.Gnomo, eRaza.Enano
770               Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
780                       MiObj.objindex = 753
790                   Case eClass.Cleric
800                       MiObj.objindex = 755
810                   Case eClass.Paladin, eClass.Warrior
820                       MiObj.objindex = 757
830                   Case eClass.Mage
840                       MiObj.objindex = 760
850           End Select
860       End Select
              
870           Case 270:
880           If Lvl < 34 Then
890                   Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 34 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
900                   Exit Sub
910               End If
920               If (UserList(Userindex).Stats.Gld >= 2500000) Then
930           UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 2500000
940           Call WriteUpdateGold(Userindex)
950           Else
960           Call WriteChatOverHead(Userindex, "Necesitas 2.500.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
970           Exit Sub
980           End If
990               UserList(Userindex).Faccion.RecompensasCaos = 3
1000              UserList(Userindex).Faccion.NextRecompensa = 350
                ' Call PerderItemsFaccionarios(Userindex)
                  '4toCAOS
1010  Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
1020              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1030                      MiObj.objindex = 761
1040                  Case eClass.Cleric
1050                      MiObj.objindex = 763
1060                  Case eClass.Paladin, eClass.Warrior
1070                      MiObj.objindex = 765
1080                  Case eClass.Mage
1090                      Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
1100                              MiObj.objindex = 767
1110                          Case eGenero.Hombre
1120                              MiObj.objindex = 768
1130                      End Select
1140                End Select
1150          Case eRaza.Gnomo, eRaza.Enano
1160              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1170                      MiObj.objindex = 762
1180                  Case eClass.Cleric
1190                      MiObj.objindex = 764
1200                  Case eClass.Paladin, eClass.Warrior
1210                      MiObj.objindex = 766
1220                  Case eClass.Mage
1230                      MiObj.objindex = 769
1240          End Select
1250      End Select
                  
              
1260          Case 350:
1270           If Lvl < 37 Then
1280                  Call WriteChatOverHead(Userindex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
1290                  Exit Sub
1300              End If
1310              If (UserList(Userindex).Stats.Gld >= 4000000) Then
1320          UserList(Userindex).Stats.Gld = UserList(Userindex).Stats.Gld - 4000000
1330          Call WriteUpdateGold(Userindex)
1340          Else
1350          Call WriteChatOverHead(Userindex, "Necesitas 4.000.000 monedas de oro para poder recibir la próxima recompensa.", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
1360          Exit Sub
1370          End If
1380              UserList(Userindex).Faccion.RecompensasCaos = 4
1390              UserList(Userindex).Faccion.NextRecompensa = 425
                  'Call PerderItemsFaccionarios(Userindex)
                  '5toCAOS
1400  Select Case UserList(Userindex).raza
              Case eRaza.Humano, eRaza.Elfo, eRaza.Drow
1410              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1420                      MiObj.objindex = 770
1430                  Case eClass.Cleric
1440                      MiObj.objindex = 772
1450                  Case eClass.Paladin, eClass.Warrior
1460                      MiObj.objindex = 774
1470                  Case eClass.Mage
1480                      Select Case UserList(Userindex).Genero
                              Case eGenero.Mujer
1490                              MiObj.objindex = 776
1500                          Case eGenero.Hombre
1510                              MiObj.objindex = 777
1520                      End Select
1530                End Select
1540          Case eRaza.Gnomo, eRaza.Enano
1550              Select Case UserList(Userindex).clase
                      Case eClass.Bard, eClass.Druid, eClass.Hunter, eClass.Assasin
1560                      MiObj.objindex = 771
1570                  Case eClass.Cleric
1580                      MiObj.objindex = 773
1590                  Case eClass.Paladin, eClass.Warrior
1600                      MiObj.objindex = 775
1610                  Case eClass.Mage
1620                      MiObj.objindex = 778
1630          End Select
1640      End Select
1650  Case 425:
1660          WriteChatOverHead Userindex, "¡¡¡Bien hecho, ya no tengo más recompensas para ti. Sigue así!!!", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite
1670              Exit Sub
1680      End Select
1690  If Not MeterItemEnInventario(Userindex, MiObj) Then
1700              Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
1710     End If

1720  Call WriteChatOverHead(Userindex, "¡¡¡Bien hecho " + TituloCaos(Userindex) + ", aquí tienes tu recompensa!!!", Str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
      'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
      'If UserList(UserIndex).Stats.Exp > MAXEXP Then
      '    UserList(UserIndex).Stats.Exp = MAXEXP
      'End If
      'Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpX100 & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
      'Call CheckUserLevel(UserIndex)


End Sub

Public Function TituloCaos(ByVal Userindex As Integer) As String
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 23/01/2007 Pablo (ToxicWaste)
      'Handles the titles of the members of the "Legión Oscura"
      '***************************************************
      'Rango 1: Esbirro (20)
      'Rango 2: Sanguinario (30 + 100k)
      'Rango 3: Condenado (40 + 200k)
      'Rango 4: Caballero de la oscuridad (50 + 375k)
      'Rango 5: Devorador de almas (100 + 500k)


10    Select Case UserList(Userindex).Faccion.RecompensasCaos
          Case 0
20            TituloCaos = "Esbirro"
30        Case 1
40            TituloCaos = "Sanguinario"
50        Case 2
60            TituloCaos = "Condenado"
70        Case 3
80            TituloCaos = "Caballero de la Oscuridad"
90        Case 4
100           TituloCaos = "Devorador de Almas"
110   End Select

End Function


Sub PerderItemsFaccionarios(ByVal Userindex As Integer)
      Dim i As Byte
      Dim MiObj As Obj
      Dim ItemIndex As Integer

10    For i = 1 To MAX_INVENTORY_SLOTS
20    ItemIndex = UserList(Userindex).Invent.Object(i).objindex
30    If ItemIndex > 0 Then
40        If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
50        QuitarUserInvItem Userindex, i, UserList(Userindex).Invent.Object(i).Amount
60        UpdateUserInv False, Userindex, i
70            If ObjData(ItemIndex).OBJType = otarmadura Or ObjData(ItemIndex).OBJType = otescudo Then
80            If ObjData(ItemIndex).Real = 1 Then UserList(Userindex).Faccion.RecibioArmaduraReal = 0
90            If ObjData(ItemIndex).Caos = 1 Then UserList(Userindex).Faccion.RecibioArmaduraCaos = 0
100     Else
110           UserList(Userindex).Faccion.RecibioArmaduraCaos = 0 Or UserList(Userindex).Faccion.RecibioArmaduraReal = 0
120     End If
130       End If
          
140       End If
          
150               Next i
End Sub



