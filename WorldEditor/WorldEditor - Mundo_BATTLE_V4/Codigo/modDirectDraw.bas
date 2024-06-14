Attribute VB_Name = "modTileEngine"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modTileEngine Nothing to do with DD
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit

Public ma(1) As Single

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.Y + CY

End Sub

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)

If UBound(HeadData()) > Head Then
    CharList(CharIndex).Head = HeadData(Head)
End If
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

bRefreshRadar = True ' GS

End Sub

Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

bRefreshRadar = True ' GS

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

bRefreshRadar = True ' GS

End Sub

Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If X - 8 < 1 Or X + 8 > 100 Or Y - 6 < 1 Or Y + 6 > 100 Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function

Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Long, ByVal Y As Long) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, Y + DeSeleccionOY) = DeSeleccionMap(X, Y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim Y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SobreX, Y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, Y + SobreY) = SeleccionMap(X, Y)
        Next
    Next
    Seleccionando = False
End Sub
Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + Y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, Y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim Y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            SeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub
' [/Loopzer]
Public Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'Last modified: 21/11/07 By Loopzer
'Last modifier: 24/11/08 by GS
'*************************************************

On Error Resume Next
Dim Y       As Integer              'Keeps track of where on map we are
Dim X       As Integer
Dim MinY    As Integer              'Start Y pos on current map
Dim MaxY    As Integer              'End Y pos on current map
Dim MinX    As Integer              'Start X pos on current map
Dim MaxX    As Integer              'End X pos on current map
Dim ScreenX As Integer              'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim Sobre   As Long
Dim iPPx    As Integer              'Usado en el Layer de Chars
Dim iPPy    As Integer              'Usado en el Layer de Chars
Dim Grh     As Grh                  'Temp Grh for show tile and blocked
Dim bCapa    As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
Dim iGrhIndex           As Long  'Usado en el Layer 1
Dim PixelOffsetXTemp    As Integer  'For centering grhs
Dim PixelOffsetYTemp    As Integer
Dim TempChar            As Char

Dim colorlist(3) As Long

colorlist(0) = D3DColorXRGB(255, 200, 0)
colorlist(1) = D3DColorXRGB(255, 200, 0)
colorlist(2) = D3DColorXRGB(255, 200, 0)
colorlist(3) = D3DColorXRGB(255, 200, 0)


MinY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
MaxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
MinX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
MaxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize
' 31/05/2006 - GS, control de Capas
If Val(frmMain.cCapas.Text) >= 1 And (frmMain.cCapas.Text) <= 4 Then
    bCapa = Val(frmMain.cCapas.Text)
Else
    bCapa = 1
End If
GenerarVista 'Loopzer
ScreenY = -8
For Y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)
        If InMapBounds(X, Y) Then
            If X > 100 Or Y < 1 Then Exit For ' 30/05/2006

            'Layer 1 **********************************
            If SobreX = X And SobreY = Y Then
                ' Pone Grh !
                Sobre = -1
                If frmMain.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, Y).Graphic(bCapa).Grh_Index
                    If frmConfigSup.MOSAICO.value = vbChecked Then
                        Dim aux As Long
                        Dim dy As Integer
                        Dim dX As Integer
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dX = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dX = 0
                        End If
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmMain.cGrh.Text) + _
                            (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, Y).Graphic(bCapa).Grh_Index <> aux Then
                                MapData(X, Y).Graphic(bCapa).Grh_Index = aux
                                Grh_Initialize MapData(X, Y).Graphic(bCapa), aux
                            End If
                        Else
                            aux = Val(frmMain.cGrh.Text) + _
                            (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, Y).Graphic(bCapa).Grh_Index <> aux Then
                                MapData(X, Y).Graphic(bCapa).Grh_Index = aux
                                Grh_Initialize MapData(X, Y).Graphic(bCapa), aux
                            End If
                        End If
                    Else
                        If MapData(X, Y).Graphic(bCapa).Grh_Index <> Val(frmMain.cGrh.Text) Then
                            MapData(X, Y).Graphic(bCapa).Grh_Index = Val(frmMain.cGrh.Text)
                            Grh_Initialize MapData(X, Y).Graphic(bCapa), Val(frmMain.cGrh.Text)
                        End If
                    End If
                End If
            Else
                Sobre = -1
            End If
            If VerCapa1 Then
                With MapData(X, Y).Graphic(1)

    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tex As Direct3DTexture8
    Dim srcwidth As Integer
    Dim Width As Integer
    Dim srcheight As Integer
    Dim Height As Integer
    Dim SrcBitmapWidth As Long
    Dim SrcBitmapHeight As Long
    Dim xb As Integer
    Dim yb As Integer
    'Dim iGrhIndex As long
    Dim srdesc As D3DSURFACE_DESC
    
    'Ready the texture
    'If grhindex = 0 Then Exit Sub
    If MapData(X, Y).Graphic(1).Grh_Index Then
    xb = (ScreenX - 1) * 32 + PixelOffsetX
    yb = (ScreenY - 1) * 32 + PixelOffsetY
   
    If MapData(X, Y).Graphic(1).Started = 1 Then
       
        MapData(X, Y).Graphic(1).frame_counter = MapData(X, Y).Graphic(1).frame_counter + ((timer_elapsed_time * 0.1) * Grh_list(MapData(X, Y).Graphic(1).Grh_Index).frame_count / MapData(X, Y).Graphic(1).frame_speed)
            If MapData(X, Y).Graphic(1).frame_counter > Grh_list(MapData(X, Y).Graphic(1).Grh_Index).frame_count Then
                MapData(X, Y).Graphic(1).frame_counter = (MapData(X, Y).Graphic(1).frame_counter Mod Grh_list(MapData(X, Y).Graphic(1).Grh_Index).frame_count) + 1
            End If
           
    End If

If MapData(X, Y).Graphic(1).frame_counter = 0 Then MapData(X, Y).Graphic(1).frame_counter = 1
If MapData(X, Y).Graphic(1).Grh_Index <= 0 Then Exit Sub
 
iGrhIndex = Grh_list(MapData(X, Y).Graphic(1).Grh_Index).frame_list(MapData(X, Y).Graphic(1).frame_counter)
    
    
    If iGrhIndex > 0 Then
    With Grh_list(iGrhIndex)
    
    Set tex = DXPool.GetTexture(.texture_index)
    'Call DXPool.Texture_Dimension_Get(.texture_index, texture_width, texture_height)
    tex.GetLevelDesc 0, srdesc
    
    'If .src_x = 0 And SrcHeight = 0 And Width = 0 And Height = 0 Then
        srcwidth = 32 'd3dtextures.texwidth
        Width = 32 'd3dtextures.texwidth
       
        Height = 32 'd3dtextures.texheight
        srcheight = 32 'd3dtextures.texheight
        SrcBitmapWidth = srdesc.Width
        SrcBitmapHeight = srdesc.Height
    'Set the RHWs (must always be 1)
   
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        
        'Find the left side of the rectangle
        VertexArray(0).X = xb
        VertexArray(0).tu = (.Src_X / SrcBitmapWidth)
 
        'Find the top side of the rectangle
        VertexArray(0).Y = yb
        VertexArray(0).tv = (.Src_Y / SrcBitmapHeight)
   
        'Find the right side of the rectangle
        VertexArray(1).X = xb + Width
        VertexArray(1).tu = (.Src_X + srcwidth) / SrcBitmapWidth
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 
    'Find the bottom of the rectangle
    VertexArray(2).Y = yb + Height
    VertexArray(2).tv = (.Src_Y + srcheight) / SrcBitmapHeight
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).Y = VertexArray(0).Y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).Y = VertexArray(2).Y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv
   
    VertexArray(0).Color = -1
    VertexArray(1).Color = -1
    VertexArray(2).Color = -1
    VertexArray(3).Color = -1
    
    If HayAgua(X, Y) Then

    Dim ignorarpoligonossuperiores As Byte
    Dim ignorarpoligonosinferiores As Byte
    ignorarpoligonosinferiores = 0
    ignorarpoligonossuperiores = 0
    If HayAgua(X, Y - 1) = False Then ignorarpoligonossuperiores = 1
    If HayAgua(X, Y + 1) = False Then ignorarpoligonosinferiores = 1
   
        If X Mod 2 = 0 Then
       
            If Y Mod 2 = 0 Then
                If ignorarpoligonossuperiores <> 1 Then
                    VertexArray(0).Y = VertexArray(0).Y - Val(ma(0))
                    VertexArray(1).Y = VertexArray(1).Y + Val(ma(0))
                End If
                If ignorarpoligonosinferiores <> 1 Then
                    VertexArray(2).Y = VertexArray(2).Y + Val(ma(1))
                    VertexArray(3).Y = VertexArray(3).Y - Val(ma(1))
                End If
               
            Else
                If ignorarpoligonossuperiores <> 1 Then
                    VertexArray(0).Y = VertexArray(0).Y + Val(ma(1))
                    VertexArray(1).Y = VertexArray(1).Y - Val(ma(1))
                End If
                If ignorarpoligonosinferiores <> 1 Then
                    VertexArray(2).Y = VertexArray(2).Y - Val(ma(0))
                    VertexArray(3).Y = VertexArray(3).Y + Val(ma(0))
                End If
               
            End If
           
        ElseIf X Mod 2 = 1 Then
       
            If Y Mod 2 = 0 Then
                If ignorarpoligonossuperiores <> 1 Then
                    VertexArray(0).Y = VertexArray(0).Y + Val(ma(0))
                    VertexArray(1).Y = VertexArray(1).Y - Val(ma(0))
                End If
                If ignorarpoligonosinferiores <> 1 Then
                    VertexArray(2).Y = VertexArray(2).Y - Val(ma(1))
                    VertexArray(3).Y = VertexArray(3).Y + Val(ma(1))
                End If
               
             Else
                If ignorarpoligonossuperiores <> 1 Then
                    VertexArray(0).Y = VertexArray(0).Y - Val(ma(1))
                    VertexArray(1).Y = VertexArray(1).Y + Val(ma(1))
                End If
               
                If ignorarpoligonosinferiores <> 1 Then
                    VertexArray(2).Y = VertexArray(2).Y + Val(ma(0))
                    VertexArray(3).Y = VertexArray(3).Y - Val(ma(0))
                End If
            End If
        End If
    End If
    
    ddevice.SetTexture 0, tex

    ddevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0) 'wiii
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))
    End With
    End If
    End If
                End With
            End If
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).Grh_Index <> 0 And VerCapa2 Then
                modGrh.Grh_Render MapData(X, Y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, True
            End If
            
            If Sobre >= 0 Then
                If MapData(X, Y).Graphic(bCapa).Grh_Index <> Sobre Then
                MapData(X, Y).Graphic(bCapa).Grh_Index = Sobre
                Grh_Initialize MapData(X, Y).Graphic(bCapa), Sobre
                End If
            End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y
ScreenY = -8
For Y = (MinY) To (MaxY)   '- 8+ 8
    ScreenX = -8
    For X = (MinX) To (MaxX)   '- 8 + 8
        If InMapBounds(X, Y) Then
            If X > 100 Or X < -3 Then Exit For ' 30/05/2006

            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
             'Object Layer **********************************
             If MapData(X, Y).OBJInfo.objindex <> 0 And VerObjetos Then
                modGrh.Grh_Render MapData(X, Y).ObjGrh, iPPx, iPPy, True
             End If
            
                  'Char layer **********************************
                 If MapData(X, Y).CharIndex <> 0 And VerNpcs Then
                 
                     TempChar = CharList(MapData(X, Y).CharIndex)
                 
                     PixelOffsetXTemp = PixelOffsetX
                     PixelOffsetYTemp = PixelOffsetY
                    
                   'Dibuja solamente players
                   If TempChar.Head.Head(TempChar.Heading).Grh_Index <> 0 Then
                     'Draw Body
                     modGrh.Grh_Render TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, True
                     'Draw Head
                     modGrh.Grh_Render TempChar.Head.Head(TempChar.Heading), iPPx, iPPy + TempChar.Body.HeadOffset.Y, True
                   Else: modGrh.Grh_Render TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, True
                   End If
                 End If
             'Layer 3 *****************************************
             If MapData(X, Y).Graphic(3).Grh_Index <> 0 And VerCapa3 Then
                'Draw
                modGrh.Grh_Render MapData(X, Y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, True
             End If
        
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y
'Tiles blokeadas, techos, triggers , seleccion
ScreenY = -8
For Y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)
        If X < 101 And X > 0 And Y < 101 And Y > 0 Then ' 30/05/2006
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            If MapData(X, Y).Graphic(4).Grh_Index <> 0 _
            And (frmMain.mnuVerCapa4.Checked = True) Then
                'Draw
                modGrh.Grh_Render MapData(X, Y).Graphic(4), iPPx, iPPy, True
            End If
            If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
                Grh.Grh_Index = 3
                Grh.frame_counter = 1
                Grh.Started = 0
                modGrh.Grh_Render Grh, iPPx, iPPy, True
            End If

            
            'Show blocked tiles
            If VerBlockeados And MapData(X, Y).Blocked = 1 Then
                Grh.Grh_Index = 4
                Grh.frame_counter = 1
                Grh.Started = 0
                modGrh.Grh_Render Grh, iPPx, iPPy, True
            End If
            If VerGrilla Then
                'Grilla 24/11/2008 by GS
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 32, RGB(255, 255, 255)
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 1, RGB(255, 255, 255)
            End If
            
            If VerTriggers Then
               If MapData(X, Y).Trigger <> 0 Then
                    Dim lColor As Long
                    Select Case MapData(X, Y).Trigger
                         Case 1
                             lColor = D3DColorXRGB(255, 0, 0)
                         Case 2
                             lColor = D3DColorXRGB(0, 255, 0)
                         Case 3
                             lColor = D3DColorXRGB(0, 0, 255)
                         Case 4
                             lColor = D3DColorXRGB(0, 255, 255)
                         Case 5
                             lColor = D3DColorXRGB(255, 64, 0)
                         Case 6
                             lColor = D3DColorXRGB(255, 128, 255)
                         Case Else
                             lColor = D3DColorXRGB(255, 255, 0)
                     End Select
                    
                    modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1, lColor
                    modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 32, lColor
                    modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 1, lColor
                    modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 32, lColor
                End If
            End If
                'Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(X, Y).Trigger), D3DColorXRGB(102, 150, 150))
            If Seleccionando Then
                'If ScreenX >= SeleccionIX And ScreenX <= SeleccionFX And ScreenY >= SeleccionIY And ScreenY <= SeleccionFY Then
                    If X >= SeleccionIX And Y >= SeleccionIY Then
                        If X <= SeleccionFX And Y <= SeleccionFY Then
                            modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 32, RGB(100, 255, 255)
                        End If
                    End If
            End If

        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

End Sub



Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
    If LenB(strText) <> 0 Then
        Call modDXEngine.DXEngine_TextRender(1, strText, lngXPos, lngYPos, D3DColorXRGB(255, 255, 255))
    End If
End Sub

Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 15/10/06 by GS
'*************************************************
    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    '[GS] 02/10/2006
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    InitTileEngine = True
    EngineRun = True
    DoEvents
End Function

Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean

        If InMapBounds(X, Y) Then

                With MapData(X, Y)

                        If ((.Graphic(1).Grh_Index >= 1505 And .Graphic(1).Grh_Index <= 1520) Or (.Graphic(1).Grh_Index >= 5665 And _
                            .Graphic(1).Grh_Index <= 5680) Or (.Graphic(1).Grh_Index >= 13547 And _
                                .Graphic(1).Grh_Index <= 13562)) And .Graphic(2).Grh_Index = 0 Then
                    
                                HayAgua = True
                        Else
                                HayAgua = False
                        End If

                End With

        End If
                  
End Function

