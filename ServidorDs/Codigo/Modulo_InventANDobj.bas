Attribute VB_Name = "InvNpc"
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
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj, Optional NotPirata As Boolean = True) As WorldPos
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler

          Dim NuevaPos As WorldPos
20        NuevaPos.X = 0
30        NuevaPos.Y = 0
          
40        Tilelibre Pos, NuevaPos, Obj, NotPirata, True
50        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
60            Call MakeObj(Obj, Pos.map, NuevaPos.X, NuevaPos.Y)
70        End If
80        TirarItemAlPiso = NuevaPos

90        Exit Function
Errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef Npc As Npc, ByVal IsPretoriano As Boolean)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 28/11/2009
      'Give away npc's items.
      '28/11/2009: ZaMa - Implementado drops complejos
      '02/04/2010: ZaMa - Los pretos vuelven a tirar oro.
      '***************************************************
10    On Error Resume Next

          Dim LoopC As Integer
20        With Npc
              
              Dim i As Byte
              Dim MiObj As Obj
              Dim NroDrop As Integer
              Dim Random As Integer
              Dim objindex As Integer
              
              ' Tira todo el inventario
              
30                For i = 1 To MAX_INVENTORY_SLOTS
40                    If .Invent.Object(i).objindex > 0 Then
50                          MiObj.Amount = .Invent.Object(i).Amount
60                          MiObj.objindex = .Invent.Object(i).objindex
70                          Call TirarItemAlPiso(.Pos, MiObj)
                                                             
80                    End If
90                Next i


                  'Agrega a la lista de objetos - maTih.-
          
               
               
100            For LoopC = 1 To MAX_NPC_DROPS
110               If .Drop(LoopC).objindex > 0 Then
120                   If RandomNumber(1, 100) <= .Drop(LoopC).Probability Then
130                       MiObj.Amount = .Drop(LoopC).Amount
140                       MiObj.objindex = .Drop(LoopC).objindex
                          
150                       Call TirarItemAlPiso(.Pos, MiObj)
160                   End If
170               End If
180           Next LoopC
190       End With
          

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal objindex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next

          Dim i As Integer
20        If Npclist(NpcIndex).Invent.NroItems > 0 Then
30            For i = 1 To MAX_INVENTORY_SLOTS
40                If Npclist(NpcIndex).Invent.Object(i).objindex = objindex Then
50                    QuedanItems = True
60                    Exit Function
70                End If
80            Next
90        End If
100       QuedanItems = False
End Function

''
' Gets the amount of a certain item that an npc has.
'
' @param npcIndex Specifies reference to npcmerchant
' @param ObjIndex Specifies reference to object
' @return   The amount of the item that the npc has
' @remarks This function reads the Npc.dat file
Function EncontrarCant(ByVal NpcIndex As Integer, ByVal objindex As Integer) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: 03/09/08
      'Last Modification By: Marco Vanotti (Marco)
      ' - 03/09/08 EncontrarCant now returns 0 if the npc doesn't have it (Marco)
      '***************************************************
10    On Error Resume Next
      'Devuelve la cantidad original del obj de un npc

          Dim ln As String, npcfile As String
          Dim i As Integer
          
20        npcfile = DatPath & "NPCs.dat"
           
30        For i = 1 To MAX_INVENTORY_SLOTS
40            ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
50            If objindex = val(ReadField(1, ln, 45)) Then
60                EncontrarCant = val(ReadField(2, ln, 45))
70                Exit Function
80            End If
90        Next
                             
100       EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error Resume Next

          Dim i As Integer
          
20        With Npclist(NpcIndex)
30            .Invent.NroItems = 0
              
40            For i = 1 To MAX_INVENTORY_SLOTS
50               .Invent.Object(i).objindex = 0
60               .Invent.Object(i).Amount = 0
70            Next i
              
80            .InvReSpawn = 0
90        End With

End Sub

''
' Removes a certain amount of items from a slot of an npc's inventory
'
' @param npcIndex Specifies reference to npcmerchant
' @param Slot Specifies reference to npc's inventory's slot
' @param antidad Specifies amount of items that will be removed
Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 23/11/2009
      'Last Modification By: Marco Vanotti (Marco)
      ' - 03/09/08 Now this sub checks that te npc has an item before respawning it (Marco)
      '23/11/2009: ZaMa - Optimizacion de codigo.
      '***************************************************
          Dim objindex As Integer
          Dim iCant As Integer
          
10        With Npclist(NpcIndex)
20            objindex = .Invent.Object(Slot).objindex
          
              'Quita un Obj
30            If ObjData(.Invent.Object(Slot).objindex).Crucial = 0 Then
40                .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
                  
50                If .Invent.Object(Slot).Amount <= 0 Then
60                    .Invent.NroItems = .Invent.NroItems - 1
70                    .Invent.Object(Slot).objindex = 0
80                    .Invent.Object(Slot).Amount = 0
90                    If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
100                      Call CargarInvent(NpcIndex) 'Reponemos el inventario
110                   End If
120               End If
130           Else
140               .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - Cantidad
                  
150               If .Invent.Object(Slot).Amount <= 0 Then
160                   .Invent.NroItems = .Invent.NroItems - 1
170                   .Invent.Object(Slot).objindex = 0
180                   .Invent.Object(Slot).Amount = 0
                      
190                   If Not QuedanItems(NpcIndex, objindex) Then
                          'Check if the item is in the npc's dat.
200                       iCant = EncontrarCant(NpcIndex, objindex)
210                       If iCant Then
220                           .Invent.Object(Slot).objindex = objindex
230                           .Invent.Object(Slot).Amount = iCant
240                           .Invent.NroItems = .Invent.NroItems + 1
250                       End If
260                   End If
                      
270                   If .Invent.NroItems = 0 And .InvReSpawn <> 1 Then
280                      Call CargarInvent(NpcIndex) 'Reponemos el inventario
290                   End If
300               End If
310           End If
320       End With
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'Vuelve a cargar el inventario del npc NpcIndex
          Dim LoopC As Integer
          Dim ln As String
          Dim npcfile As String
          
10        npcfile = DatPath & "NPCs.dat"
          
20        With Npclist(NpcIndex)
30            .Invent.NroItems = val(GetVar(npcfile, "NPC" & .Numero, "NROITEMS"))
              
40            For LoopC = 1 To .Invent.NroItems
50                ln = GetVar(npcfile, "NPC" & .Numero, "Obj" & LoopC)
60                .Invent.Object(LoopC).objindex = val(ReadField(1, ln, 45))
70                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                  
80            Next LoopC
90        End With

End Sub


Public Sub TirarOroNpc(ByVal Cantidad As Long, ByRef Pos As WorldPos)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 13/02/2010
      '***************************************************
10    On Error GoTo Errhandler

20        If Cantidad > 0 Then
              Dim i As Byte
              Dim MiObj As Obj
              Dim RemainingGold As Long
              
30            RemainingGold = Cantidad
              
40            While (RemainingGold > 0)
                  
                  ' Tira pilon de 10k
50                If RemainingGold > MAX_INVENTORY_OBJS Then
60                    MiObj.Amount = MAX_INVENTORY_OBJS
70                    RemainingGold = RemainingGold - MAX_INVENTORY_OBJS
                      
                  ' Tira lo que quede
80                Else
90                    MiObj.Amount = RemainingGold
100                   RemainingGold = 0
110               End If

120               MiObj.objindex = iORO
                  
130               Call TirarItemAlPiso(Pos, MiObj)
140           Wend
150       End If

160       Exit Sub

Errhandler:
170       Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.Description)
End Sub
