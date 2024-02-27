VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComerciar 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.TextBox cantidad 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2910
      TabIndex        =   2
      Text            =   "1"
      Top             =   3135
      Width           =   690
   End
   Begin VB.PictureBox picInvUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   3480
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   1
      Top             =   360
      Width           =   2880
   End
   Begin VB.PictureBox picInvNpc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   240
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   360
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2625
      TabIndex        =   6
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haz Click en algun item más información."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2865
      TabIndex        =   4
      Top             =   2895
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5565
      TabIndex        =   3
      Top             =   120
      Width           =   105
   End
   Begin VB.Image imgCross 
      Height          =   450
      Left            =   5880
      MouseIcon       =   "frmComerciar.frx":133CD
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3000
      Width           =   570
   End
   Begin VB.Image imgVender 
      Height          =   465
      Left            =   3960
      MouseIcon       =   "frmComerciar.frx":136D7
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3000
      Width           =   1380
   End
   Begin VB.Image imgComprar 
      Height          =   465
      Left            =   720
      MouseIcon       =   "frmComerciar.frx":13829
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3000
      Width           =   1500
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Private clsFormulario As clsFormMovementManager

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean
Private lIndex As Byte

Private cBotonVender As clsGraphicalButton
Private cBotonComprar As clsGraphicalButton
Private cBotonCruz As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub cantidad_Change()
10        If Val(cantidad.Text) < 1 Then
20            cantidad.Text = 1
30        End If
          
40        If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
50            cantidad.Text = MAX_INVENTORY_OBJS
60        End If
          Dim ItemSlot As Byte
70        ItemSlot = InvComNpc.SelectedItem
          
80        If ClickNpcInv Then
90        Label1(0).Caption = "Precio : " & _
              CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No _
              mostramos numeros reales
100       Else
110           If InvComUsu.SelectedItem <> 0 Then
120               Label1(0).Caption = "Precio : " & _
                      CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), _
                      Val(cantidad.Text))  'No mostramos numeros reales
130           End If
140       End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
10        If (KeyAscii <> 8) Then
20            If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
30                KeyAscii = 0
40            End If
50        End If
End Sub

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me

          
          'Cargamos la interfase
      '    Me.Picture = LoadPicture(DirGraficos & "ventanacomercio.jpg")
          

          
End Sub


''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount _
    As Long) As Long
      '*************************************************
      'Author: Marco Vanotti (MarKoxX)
      'Last modified: 19/08/2008
      'Last modify by: Franco Zeoli (Noich)
      '*************************************************
10        On Error GoTo error
          'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
20        CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
          
30        Exit Function
error:
40        MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As _
    Long) As Long
      '*************************************************
      'Author: Marco Vanotti (MarKoxX)
      'Last modified: 19/08/2008
      'Last modify by: Franco Zeoli (Noich)
      '*************************************************
10        On Error GoTo error
          'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
20        CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
          
30        Exit Function
error:
40        MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub imgComprar_Click()
    Static LastClick As Long
    
    If timeGetTime - LastClick <= 200 Then Exit Sub
    
    LastClick = timeGetTime
          ' Debe tener seleccionado un item para comprarlo.
10        If InvComNpc.SelectedItem = 0 Then Exit Sub
          
20        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
          
30        Call Audio.PlayWave(SND_CLICK)
          
40        LasActionBuy = True
50        If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, _
              Val(cantidad.Text)) Then
60            Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
70        Else
80            Call AddtoRichTextBox(frmMain.RecTxt, "Se necesita más oro.", 2, 51, _
                  223, 1, 1)
90            Exit Sub
100       End If
          
End Sub


Private Sub imgCross_Click()
10        Call WriteCommerceEnd
End Sub

Private Sub imgVender_Click()
          ' Debe tener seleccionado un item para comprarlo.
10        If InvComUsu.SelectedItem = 0 Then Exit Sub

20        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
          
30        Call Audio.PlayWave(SND_CLICK)
          
40        LasActionBuy = False

50        Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
'    'LastPressed.ToggleToNormal
End Sub

Private Sub picInvNpc_Click()
             Dim ItemSlot As Byte
          
10        ItemSlot = InvComNpc.SelectedItem
20        Call Audio.PlayWave(SND_CLICK)
30        If ItemSlot = 0 Then Exit Sub
          
40        ClickNpcInv = True
50        InvComUsu.DeselectItem
          
60        Label1(0).Caption = NPCInventory(ItemSlot).Name
70            If NPCInventory(ItemSlot).Copas > 0 Then
80        Label1(1).Caption = "DsP : " & NPCInventory(ItemSlot).Copas
90        cantidad.Enabled = False
100       Else
110        If NPCInventory(ItemSlot).Eldhir > 0 Then
120       Label1(1).Caption = "Eldhires : " & NPCInventory(ItemSlot).Eldhir
130       cantidad.Enabled = False
140       Else
150       Label1(1).Caption = "Valor : " & _
              CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No _
              mostramos numeros reales
160       End If
170       End If
          
180       If NPCInventory(ItemSlot).Amount <> 0 Then
          
190           Select Case NPCInventory(ItemSlot).ObjType
                  Case eOBJType.otWeapon
200                   Label1(2).Caption = "Hit: " & NPCInventory(ItemSlot).MinHit & _
                          "/" & NPCInventory(ItemSlot).MaxHit
210                   Label1(2).Visible = True
220                   Label1(3).Visible = True
230               Case eOBJType.otArmadura, eOBJType.otcasco, eOBJType.otescudo
240                   Label1(2).Caption = "Def: " & NPCInventory(ItemSlot).MinDef & _
                          "/" & NPCInventory(ItemSlot).MaxDef
250                   Label1(2).Visible = True
260                   Label1(3).Visible = True
270               Case Else
280                   Label1(2).Visible = False
290                   Label1(3).Visible = False
300           End Select
310       Else
320           Label1(2).Visible = False
330           Label1(3).Visible = False
340       End If
End Sub

Private Sub picInvUser_Click()
      Dim ItemSlot As Byte
          
10        ItemSlot = InvComUsu.SelectedItem
          
20        If ItemSlot = 0 Then Exit Sub
          
30        ClickNpcInv = False
40        InvComNpc.DeselectItem
          
50        Label1(0).Caption = Inventario.ItemName(ItemSlot)
60        Label1(1).Caption = "Precio : " & _
              CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text)) 'No _
              mostramos numeros reales
          
70        If Inventario.Amount(ItemSlot) <> 0 Then
          
80            Select Case Inventario.ObjType(ItemSlot)
                  Case eOBJType.otWeapon
90                    Label1(2).Caption = "Hit: " & Inventario.MinHit(ItemSlot) & "/" _
                          & Inventario.MaxHit(ItemSlot)
100                   Label1(3).Caption = ""
110                   Label1(2).Visible = True
120                   Label1(3).Visible = True
130               Case eOBJType.otArmadura, eOBJType.otcasco, eOBJType.otescudo
140                   Label1(2).Caption = "Def: " & Inventario.MinDef(ItemSlot) & "/" _
                          & Inventario.MaxDef(ItemSlot)
150                   Label1(3).Caption = ""
160                   Label1(2).Visible = True
170                   Label1(3).Visible = True
180               Case Else
190                   Label1(2).Visible = False
200                   Label1(3).Visible = False
210           End Select
220       Else
230           Label1(2).Visible = False
240           Label1(3).Visible = False
250       End If
End Sub
Private Sub PicInvNpc_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      Dim Position As Integer
      Dim i As Long
      Dim file_path As String
      Dim data() As Byte
      Dim bmpInfo As BITMAPINFO
      Dim handle As Integer
      Dim bmpData As StdPicture
      Dim Last_I As Long
10    If (Button = vbRightButton) Then

20        If InvComNpc.GrhIndex(InvComNpc.SelectedItem) > 0 Then

30            Last_I = InvComNpc.SelectedItem
40            If Last_I > 0 And Last_I <= MAX_NPC_INVENTORY_SLOTS Then
                          
50                Position = BuscarI(InvComNpc.GrhIndex(InvComNpc.SelectedItem))
                  
60                If Position = 0 Then
70                    i = GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum
80                    Call Get_Bitmapp(DirGraficos, _
                          CStr(GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum) _
                          & ".BMP", bmpInfo, data)
90                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
100                   ImageList1.ListImages.Add , CStr("g" & _
                          InvComNpc.GrhIndex(InvComNpc.SelectedItem)), Picture:=bmpData
110                   Position = ImageList1.ListImages.Count
120                   Set bmpData = Nothing
130               End If
                  
                  
                '  InvComNpc.uMoveItem = True
                  
140               Set picInvNpc.MouseIcon = _
                      ImageList1.ListImages(Position).ExtractIcon
150               picInvNpc.MousePointer = vbCustom

160               Exit Sub
170           End If
180       End If
190   End If
End Sub
 
Private Sub picInvNpc_MouseUp(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
       
      ' @Wildem
       
      'Pongo el puntero por default primero.
10    picInvNpc.MousePointer = vbDefault
       
20    If X > 0 And X < picInvNpc.ScaleWidth And Y > 0 And Y < picInvNpc.ScaleHeight _
          Then
       
          'mmm, dejo por si alguien quiere agregarle algo (?
       
30    Else
          ' Debe tener seleccionado un item para comprarlo.
40        If InvComNpc.SelectedItem = 0 Then Exit Sub
         
50        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
         
60        Call Audio.PlayWave(SND_CLICK)
         
70        LasActionBuy = True
         
80        If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, _
              Val(cantidad.Text)) Then
90            Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
100       Else
110           Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente oro.", 2, _
                  51, 223, 1, 1)
120           Exit Sub
130       End If
         
140       picInvNpc.MousePointer = vbDefault
150   End If
160   InvComNpc.DrawInventory
170   InvComUsu.DrawInventory
180    InvComNpc.sMoveItem = False
190    InvComNpc.uMoveItem = False
End Sub
 
Private Function BuscarI(gh As Integer) As Integer
      Dim i As Integer
       
10    For i = 1 To frmComerciar.ImageList1.ListImages.Count
20        If frmComerciar.ImageList1.ListImages(i).key = "g" & CStr(gh) Then
30            BuscarI = i
40            Exit For
50        End If
60    Next i
       
End Function
 
Private Sub PicInvUser_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      Dim Position As Integer
      Dim i As Long
      Dim file_path As String
      Dim data() As Byte
      Dim bmpInfo As BITMAPINFO
      Dim handle As Integer
      Dim bmpData As StdPicture
      Dim Last_I As Long
10    If (Button = vbRightButton) Then

20        If InvComUsu.GrhIndex(InvComUsu.SelectedItem) > 0 Then

30            Last_I = InvComUsu.SelectedItem
40            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                          
50                Position = BuscarI(InvComUsu.GrhIndex(InvComUsu.SelectedItem))
                  
60                If Position = 0 Then
70                    i = GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum
80                    Call Get_Bitmapp(DirGraficos, _
                          CStr(GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum) _
                          & ".BMP", bmpInfo, data)
90                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
100                   ImageList1.ListImages.Add , CStr("g" & _
                          InvComUsu.GrhIndex(InvComUsu.SelectedItem)), Picture:=bmpData
110                   Position = ImageList1.ListImages.Count
120                   Set bmpData = Nothing
130               End If
                  
                  
                '  InvComUsu.uMoveItem = True
                  
140               Set picInvUser.MouseIcon = _
                      ImageList1.ListImages(Position).ExtractIcon
150               picInvUser.MousePointer = vbCustom

160               Exit Sub
170           End If
180       End If
190   End If
End Sub

 
Private Sub picInvUser_MouseUp(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
       
      'Pongo el puntero por default primero.
10    picInvUser.MousePointer = vbDefault
       
20    If Not X > 0 And X < picInvUser.ScaleWidth And Y > 0 And Y < _
          picInvUser.ScaleHeight Then
          ' Debe tener seleccionado un item para comprarlo.
30        If InvComUsu.SelectedItem = 0 Then Exit Sub
       
40        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
         
50        Call Audio.PlayWave(SND_CLICK)
         
60        LasActionBuy = False
       
70        Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
         
80        picInvUser.MousePointer = vbDefault
90    End If
100   InvComNpc.DrawInventory
110   InvComUsu.DrawInventory
120   InvComUsu.sMoveItem = False
130   InvComUsu.uMoveItem = False
       
End Sub
