VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If Wgl = 0 Then

'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************
 
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
 
Option Explicit
 
Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item
 Private MouseDownSelectedItem As Long
 
Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer
 
Private ShowValue As Boolean
Private ValueFormated As Boolean
 
Private ImgContainer As Boolean
 
Private slotDragged As Integer
 
 Private dragAndDropOn As Boolean
 Public uMoveItem As Boolean 'USANDO DRAG
Public sMoveItem As Boolean 'USABA DRAG
 
Private InvSurface As DirectDrawSurface7            'DD Surface used to render everything
Private bufferInvSurface As DirectDrawSurface7
 
Private Inventory() As Inventory    'User's inventory
  Public Event dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

 
#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If
 
Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub
 
Private Sub Class_Terminate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      '
      '***************************************************
10        Set InvSurface = Nothing
20        Set bufferInvSurface = Nothing
End Sub
 
Public Sub ScrollInventory(ByVal Up As Boolean)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 14/01/2010
      'Scrolls the graphical inventory up / down
      '14/01/2010: ZaMa - Now redraws inventory only if needed
      '***************************************************
       
          'Only move if there is something to show up /down
10        If CanScroll(Up) Then
20            InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
             
              'Redraw the inventory if needed
30            DrawInventory
40        End If
       
End Sub
 
Public Function CanScroll(ByVal Up As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 14/01/2010
      'Determines wether inventory can be scrolled up / down
      '***************************************************
10        If Up Then
20            CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ _
                  ownTilePixelHeight)) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) _
                  < UBound(Inventory)
30        Else
40            CanScroll = InventoryOffset > 0
50        End If
         
End Function
 
Public Sub SelectGold()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Sets the gold as the currently selected item
      '***************************************************
          Dim prevSelectedItem As Long
         
          'Store preivously selected item
10        prevSelectedItem = InvSelectedItem
         
          'Select the gold
20        InvSelectedItem = FLAGORO
         
          'Redraw old item to deselect it
30        If prevSelectedItem <> FLAGORO Then
40            Call DrawInventorySlot(prevSelectedItem)
50        End If
End Sub
 
Public Sub DeselectItem()
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 30/11/2009
      'Deselect the currently selected item
      '***************************************************
10        If InvSelectedItem = 0 Then Exit Sub
         
          Dim ClearSlot As Byte
         
20        ClearSlot = InvSelectedItem
         
          'Select nothing
30        InvSelectedItem = 0
         
          'Redraw old item to deselect it
40        Call DrawInventorySlot(ClearSlot)
       
End Sub
 
Public Sub ChangeSlotItemAmount(ByVal Slot As Byte, ByVal Amount As Long)
      '***************************************************
      'Author: ZaMa
      'Created Date: 30/11/2009
      'Last Modify Date: 09/12/2009 (Budi)
      'Change the amount of the required item, and update the slot.
      '***************************************************
         
10        Inventory(Slot).Amount = Amount
         
          'Update Amount
20        Call DrawInventorySlot(Slot)
       
End Sub
 
Public Property Get SelectedItem() As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the selected item index
      '***************************************************
10        SelectedItem = InvSelectedItem
End Property
 
Public Property Get MaxHit(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the max hit of the selected item
      '***************************************************
10        MaxHit = Inventory(Slot).MaxHit
End Property
 
Public Property Get MinHit(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the min hit of the selected item
      '***************************************************
10        MinHit = Inventory(Slot).MinHit
End Property
 
Public Property Get MaxDef(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the defense of the selected item
      '***************************************************
10        MaxDef = Inventory(Slot).MaxDef
End Property
 
Public Property Get MinDef(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Budi
      'Last Modify Date: 02/1209
      'Retrieves the defense of the selected item
      '***************************************************
10        MinDef = Inventory(Slot).MinDef
End Property
 
Public Property Get GrhIndex(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the grh index of the selected item
      '***************************************************
      'If Inventario.SelectedItem = 0 And vbRightButton Then Exit Sub
10    If (Slot <= 0) Then Exit Property
20        GrhIndex = Inventory(Slot).GrhIndex
End Property
 
Public Property Get Valor(ByVal Slot As Byte) As Single
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the value of the selected item
      '***************************************************
10        Valor = Inventory(Slot).Valor
End Property
 
Public Property Get Amount(ByVal Slot As Byte) As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the selected item's amount
      '***************************************************
10        If Slot = FLAGORO Then
20            Amount = UserGLD
30        ElseIf Slot >= LBound(Inventory) And Slot <= UBound(Inventory) Then
40            Amount = Inventory(Slot).Amount
50        End If
End Property
 
Public Property Get ObjIndex(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the selected item's object index
      '***************************************************
10        ObjIndex = Inventory(Slot).ObjIndex
End Property
 
Public Property Get ObjType(ByVal Slot As Byte) As Integer
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the selected item's object type
      '***************************************************
10        ObjType = Inventory(Slot).ObjType
End Property
 
Public Property Get ItemName(ByVal Slot As Byte) As String
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves the selected item's name
      '***************************************************
10        ItemName = Inventory(Slot).Name
End Property
 Public Property Get Equipped(ByVal Slot As Byte) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Retrieves True if the item at the given pos is eqiupped
      '***************************************************
10        Equipped = Inventory(Slot).Equipped
End Property
 
Public Property Get MaxObjs() As Byte
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 09/16/09
      'Retrieves the capacity of the Inventory
      '***************************************************
10        MaxObjs = UBound(Inventory)
End Property
 
Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
      '***************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 09/16/09
      'Set the capacity of the Inventary
      '***************************************************
10        If UBound(Inventory) = MaxObjs Then Exit Sub
         
20        ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub
 
Public Sub ToggleShowValue(ByVal bShow As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 11/12
      'Indicates if the amount of the items in the inventory is shown
      '***************************************************
10        ShowValue = bShow
         
          Dim Slot As Integer
          ' Update Inventory
20        For Slot = 1 To Me.MaxObjs
30            DrawInventorySlot (Slot)
40        Next Slot
End Sub
 
Public Sub ClearAllSlots()
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 15/03/2011
      'Clear All Slots.
      '***************************************************
       
          Dim Slot As Long
10        For Slot = 1 To Me.MaxObjs
20            Call SetItem(Slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
30        Next Slot
       
End Sub
 
Public Sub SetItem(ByVal Slot As Byte, ByVal eOBJIndex As Integer, ByVal _
    eAmount As Long, ByVal eEquipped As Byte, ByVal eGrhIndex As Integer, ByVal _
    eOBJType As eOBJType, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, ByVal _
    MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName _
    As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 12/04/06
      'Sets all data for a given inventory slot
      '***************************************************
10        If Slot < 1 Or Slot > UBound(Inventory) Then Exit Sub
         
20        With Inventory(Slot)
30            .Amount = eAmount
40            .MaxDef = MaxDef
50            .MinDef = MinDef
60            .Equipped = eEquipped
70            .GrhIndex = eGrhIndex
80            .MaxHit = eMaxHit
90            .MinHit = eMinHit
100           .Name = eName
110           .ObjIndex = eOBJIndex
120           .ObjType = eOBJType
130           .Valor = eValor
140       End With
         
          'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
         
          'Render inventory slot (we don't need to render the whole inventory)
150       Call DrawInventorySlot(Slot)
End Sub
 
Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Selects the item clicked if it's valid and return's it's index
      '***************************************************
          Dim TempItem As Long
          Dim temp_x As Long
          Dim temp_y As Long
         
10        temp_x = X \ ownTilePixelWidth
20        temp_y = Y \ ownTilePixelHeight
         
30        TempItem = temp_x + (temp_y + InventoryOffset) * _
              (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
         
         'Make sure it's within limits
40        If TempItem <= UBound(Inventory) Then
              'Make sure slot isn't empty
50            If Inventory(TempItem).GrhIndex Then
60                ClickItem = TempItem
70            Else
80                ClickItem = 0
90            End If
100       End If
         
          'Make sure it's within limits
         ' If TempItem <= UBound(Inventory) Then End If
              'Make sure slot isn't empty
110       If TempItem <= UBound(Inventory()) Then
120           ClickItem = TempItem
130       End If
          
End Function
 
Public Sub DrawInventory()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Renders the inventory to the given PictureBox
      '***************************************************
          Dim LoopC As Long
          Dim SrcRect As RECT
          Dim TempRect As RECT
          Dim destRect As RECT
         
10        With TempRect
20            .Bottom = InventoryWindow.Height
30            .Right = InventoryWindow.Width
40        End With
         
50        With destRect
60            .Bottom = ownTilePixelHeight
70            .Right = ownTilePixelWidth
80        End With
         
          'Clear the inventory window
90        InventoryWindow.Cls
         
          'Clear the slot area
100       Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
         
110       For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ _
              ownTilePixelWidth) + 1 To UBound(Inventory)
          'For loopc = 1 To UBound(Inventory)
         
120           If Inventory(LoopC).GrhIndex Then
                  'Get source rect
130               With SrcRect
140                   .Left = GrhData(Inventory(LoopC).GrhIndex).sX
150                   .Top = GrhData(Inventory(LoopC).GrhIndex).sY
160                   .Right = .Left + ownTilePixelWidth
170                   .Bottom = .Top + ownTilePixelHeight
180               End With
                 
                  'Render the item grh and the amount
190               InvSurface.SetForeColor vbWhite
200               InvSurface.BltFast destRect.Left + startdX, destRect.Top + startdY, _
                      SurfaceDB.Surface(GrhData(Inventory(LoopC).GrhIndex).FileNum), _
                      SrcRect, DDBLTFAST_WAIT
210
                 
                  Dim sValue As String
               sValue = IIf(ValueFormated, Format(Inventory(LoopC).Amount, _
                      "#,###,###,###"), Inventory(LoopC).Amount)
             '     sValue = Inventory(slot).Amount
            InvSurface.DrawText destRect.Left + startdX / 2, destRect.Top + _
                      startdY / 2, sValue, False

                  'If equipped we render "E"
220               If Inventory(LoopC).Equipped Then
230                   InvSurface.SetForeColor vbYellow
240                   InvSurface.DrawText destRect.Left + 20, destRect.Top + 20, "+", _
                          False
250               End If
               
                  'Render a box around the selected item
260               If InvSelectedItem = LoopC Then
270                   If ImgContainer Then
280                         InvSurface.SetForeColor RGB(142, 107, 35)
290                   InvSurface.setDrawStyle DrawStyleConstants.vbDash
300                   InvSurface.DrawBox destRect.Left, destRect.Top, destRect.Left + _
                          ownTilePixelWidth, destRect.Top + ownTilePixelHeight
310                   End If
320               End If
330           End If
             
              'Render the item to the Inventory Window
340           With destRect
                  'Check if we have reached the right end of the inventory picturebox
350               If .Left = (InventoryWindow.ScaleWidth - _
                      (InventoryWindow.ScaleWidth Mod ownTilePixelWidth) - _
                      ownTilePixelWidth) Then
360                   .Left = 0
370                   .Right = ownTilePixelWidth
380                   .Bottom = .Bottom + ownTilePixelHeight
390                   .Top = .Top + ownTilePixelHeight
400               Else
410                   .Left = .Left + ownTilePixelWidth
420                   .Right = .Right + ownTilePixelWidth
430               End If
440           End With
450       Next LoopC
       
460       Call bufferInvSurface.BltFast(0, 0, InvSurface, TempRect, DDBLTFAST_WAIT)
470       Call InvSurface.BltToDC(InventoryWindow.hdc, TempRect, TempRect)
         
480       Call InventoryWindow.Refresh
End Sub
 
Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Implements the mouse move event of the inventory picture box
      'Displays a ToolTip of the item under the mouse
      '***************************************************
          Dim temp_x As Integer
          Dim temp_y As Integer
          Dim TempItem As Integer
         
          'Exit if it got outside the control's area
10        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > _
              InventoryWindow.Height Then
20            If uMoveItem = True Then sMoveItem = True
30            Exit Sub
40        End If
          
50        If Button = vbRightButton Then
60            uMoveItem = True
70        Else
80            temp_x = X \ TilePixelWidth
90            temp_y = Y \ TilePixelHeight
              
100           TempItem = temp_x + (temp_y + InventoryOffset) * _
                  (InventoryWindow.ScaleWidth \ TilePixelWidth) + 1
              
              
              
              'If frmComerciarUsu.Visible = False Then
110               If TempItem <= UBound(Inventory()) Then
                  
                  'If TempItem <= MAX_INVENTORY_SLOTS Then
120                   InventoryWindow.ToolTipText = Inventory(TempItem).Name
130               End If
              'End If
140       End If
          
End Sub
 
Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
          Dim temp_x As Integer
          Dim temp_y As Integer
          Dim TempItem As Integer
         
10        If dragAndDropOn Then
              'Exit if it got outside the control's area
20            If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > _
                  InventoryWindow.Height Then Exit Sub
             
30            If Button = vbRightButton Then
40                temp_x = X \ ownTilePixelWidth
50                temp_y = Y \ ownTilePixelHeight
             
60                TempItem = temp_x + (temp_y + InventoryOffset) * _
                      (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
                 
70                slotDragged = TempItem
80            End If
90        End If
End Sub

 
Public Sub DrawInventorySlot(ByVal Slot As Integer)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 12/04/06
      'Renders a inventory slot to the given PictureBox
      '***************************************************
          Dim destRect As RECT
          Dim SrcRectSlot As RECT
         
          Dim upperLeftSlot As Integer
10        upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ _
              ownTilePixelWidth) + 1
         
          'If not in renderable area we exit
20        If Slot < upperLeftSlot Then Exit Sub
         
30        With destRect
40            .Top = ((Slot - upperLeftSlot) \ (InventoryWindow.ScaleWidth \ _
                  ownTilePixelWidth)) * ownTilePixelHeight
50            .Left = ((Slot - 1) Mod (InventoryWindow.ScaleWidth \ _
                  ownTilePixelWidth)) * ownTilePixelWidth
60            .Bottom = .Top + ownTilePixelHeight
70            .Right = .Left + ownTilePixelWidth
80        End With
         
          'Clear the slot area
90        Call InvSurface.BltColorFill(destRect, InventoryWindow.BackColor)
         
100       If Inventory(Slot).GrhIndex Then
              'Get source rect
110           With SrcRectSlot
120               .Left = GrhData(Inventory(Slot).GrhIndex).sX
130               .Top = GrhData(Inventory(Slot).GrhIndex).sY
140               .Right = .Left + TilePixelWidth
150               .Bottom = .Top + TilePixelHeight
160           End With
             
              'Render the item grh and the amount
170           InvSurface.SetForeColor vbWhite
180           InvSurface.BltFast destRect.Left + startdX, destRect.Top + startdY, _
                  SurfaceDB.Surface(CLng(GrhData(Inventory(Slot).GrhIndex).FileNum)), _
                  SrcRectSlot, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
             
190           If ShowValue Then
                  Dim sValue As String
200               sValue = IIf(ValueFormated, Format(Inventory(Slot).Amount, _
                      "#,###,###,###"), Inventory(Slot).Amount)
             '     sValue = Inventory(slot).Amount
210               InvSurface.DrawText destRect.Left + startdX / 2, destRect.Top + _
                      startdY / 2, sValue, False
220           End If
             
              'If equipped we render "E"
230           If Inventory(Slot).Equipped Then
240               InvSurface.SetForeColor vbYellow
250               InvSurface.DrawText destRect.Left + 20, destRect.Top + 20, "+", _
                      False
260           End If
             
              'Render a box around the selected item
270           If InvSelectedItem = Slot Then
280               If ImgContainer Then
290                   InvSurface.SetForeColor RGB(142, 107, 35)
300                   InvSurface.setDrawStyle DrawStyleConstants.vbDash
310                   InvSurface.DrawBox destRect.Left, destRect.Top, destRect.Left + _
                          ownTilePixelWidth, destRect.Top + ownTilePixelHeight
320               End If
330           End If
340       End If
         
          'Render the item to the Inventory Window
350       InvSurface.BltToDC InventoryWindow.hdc, destRect, destRect
         
360       Call InventoryWindow.Refresh
End Sub
 
#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As _
        PictureBox, ByVal MaxObjs As Byte, ByRef PopUpMenu As Menu)
       
#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As _
              PictureBox, ByVal MaxObjs As Byte, Optional ByVal FontSize As Integer = 8, _
              Optional ByVal TileWidth As Integer = 32, Optional ByVal TileHeight As _
              Integer = 32, Optional ByVal startX As Integer = 0, Optional ByVal startY _
              As Integer = 0, Optional ByVal bImgContainer As Boolean = True, Optional _
              ByVal bShowText As Boolean = True, Optional FormatValue As Boolean = False, _
              Optional ByVal dragAndDrop As Boolean = False)
#End If
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 03/12/09
      'Sets the reference to the picture box on which to render the inventory
      '03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
      '***************************************************
       
          Dim ddck As DDCOLORKEY
       
10        Set InventoryWindow = InvPic
20        ValueFormated = FormatValue
         
          'Make sure auto-redraw is set to true
30        InventoryWindow.AutoRedraw = True
         
          'Set apropiate scale (pixel)
40        InventoryWindow.ScaleMode = 3
         
50        ReDim Inventory(1 To MaxObjs) As Inventory
         
          'initialize DX stuff
          Dim SurfaceDesc As DDSURFACEDESC2
         
          'Make sure DirectDraw was correctly initialized
60        If DirectDraw Is Nothing Then Exit Sub
                 
          'Set surface's description
70        With SurfaceDesc
80            .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
90            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
100           .lHeight = InventoryWindow.Height
110           .lWidth = InventoryWindow.Width
120       End With
         
              'Create inventory surface
130       Set InvSurface = DirectDraw.CreateSurface(SurfaceDesc)
140       Set bufferInvSurface = DirectDraw.CreateSurface(SurfaceDesc)
         
150       ownTilePixelHeight = TileHeight
160       ownTilePixelWidth = TileWidth
170       startdX = startX
180       startdY = startY
         
190       ImgContainer = bImgContainer
200       ShowValue = bShowText
         
210       dragAndDropOn = dragAndDrop
         
          'Prepare surfaces for text rendering
220       InvSurface.SetFontTransparency True
      'TODO : Fonts should be in a separate class / collection
          Dim font As StdFont
230       Set font = New StdFont
          Dim Ifnt As IFont
         
240       font.Name = "Tahoma"
250       font.bold = True
260       font.italic = False
270       font.Size = FontSize
280       font.Underline = False
290       font.Strikethrough = False
         
300       Set Ifnt = font
         
310       InvSurface.SetFont Ifnt
320       bufferInvSurface.SetFont Ifnt
      '\TODO
         
330       If ImgContainer Then InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
         
#If ConMenuesConextuales = 1 Then
340       Set ItemMenu = PopUpMenu
#End If
End Sub
Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 27/07/04
      'Implements the mouse up event of the inventory picture box
      'Check outs which item was clicked
      '***************************************************
          'Store previously selected item
          Dim prevSelItem As Long
          Dim TempItem As Long
          Dim temp_x As Long
          Dim temp_y As Long
         
          'Exit if it got outside the control's area
10        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > _
              InventoryWindow.Height Then Exit Sub
         
20        prevSelItem = InvSelectedItem
         
30        If Button = vbRightButton Then
40            If dragAndDropOn Then
50                temp_x = X \ ownTilePixelWidth
60                temp_y = Y \ ownTilePixelHeight
             
70                InvSelectedItem = temp_x + (temp_y + InventoryOffset) * _
                      (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
                 
80                Call DrawInventory
                 
90                If InvSelectedItem <> slotDragged Then
100                   RaiseEvent dragDone(slotDragged, InvSelectedItem)
110               End If
120           End If
130       Else
                 'Get the currently clickced item
140         If Not sMoveItem Then InvSelectedItem = ClickItem(CInt(X), CInt(Y)): modAnalisis.ClickEnInv
                  
150           If uMoveItem And InvSelectedItem <> MouseDownSelectedItem And _
                  MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 And Not sMoveItem And _
                  Not frmComerciar.Visible Then
160               Call Protocol.WriteDragInventory(InvSelectedItem, MouseDownSelectedItem, _
                      eMoveType.Inventory)
170               uMoveItem = False
180           Else
190               uMoveItem = False
200           End If
         
              'Update needed inventory slots
210           If prevSelItem <> InvSelectedItem Then
220               If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call _
                      DrawInventorySlot(prevSelItem)
230               If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
240           End If
250       End If
End Sub

#End If
