VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
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
      Left            =   3525
      TabIndex        =   4
      Top             =   120
      Width           =   75
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
      Index           =   1
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   75
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
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    Dim ItemSlot As Byte
    ItemSlot = InvComNpc.SelectedItem
    
    If ClickNpcInv Then
    Label1(0).Caption = "Precio : " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
    Else
        If InvComUsu.SelectedItem <> 0 Then
            Label1(0).Caption = "Precio : " & CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text))  'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    
    'Cargamos la interfase
'    Me.Picture = LoadPicture(DirGraficos & "ventanacomercio.jpg")
    

    
End Sub


''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub imgComprar_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Se necesita más oro.", 2, 51, 223, 1, 1)
        Exit Sub
    End If
    
End Sub


Private Sub imgCross_Click()
    Call WriteCommerceEnd
End Sub

Private Sub imgVender_Click()
    ' Debe tener seleccionado un item para comprarlo.
    If InvComUsu.SelectedItem = 0 Then Exit Sub

    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    
    Call Audio.PlayWave(SND_CLICK)
    
    LasActionBuy = False

    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    'LastPressed.ToggleToNormal
End Sub

Private Sub picInvNpc_Click()
       Dim ItemSlot As Byte
    
    ItemSlot = InvComNpc.SelectedItem
    Call Audio.PlayWave(SND_CLICK)
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = True
    InvComUsu.DeselectItem
    
    Label1(0).Caption = NPCInventory(ItemSlot).Name
        If NPCInventory(ItemSlot).Copas > 0 Then
    Label1(1).Caption = "DsP : " & NPCInventory(ItemSlot).Copas
    cantidad.Enabled = False
    Else
     If NPCInventory(ItemSlot).Eldhir > 0 Then
    Label1(1).Caption = "Eldhires : " & NPCInventory(ItemSlot).Eldhir
    cantidad.Enabled = False
    Else
    Label1(1).Caption = "Valor : " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
    End If
    End If
    
    If NPCInventory(ItemSlot).amount <> 0 Then
    
        Select Case NPCInventory(ItemSlot).OBJType
            Case eObjType.otWeapon
                Label1(2).Caption = "Hit: " & NPCInventory(ItemSlot).MinHit & "/" & NPCInventory(ItemSlot).MaxHit
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Def: " & NPCInventory(ItemSlot).MinDef & "/" & NPCInventory(ItemSlot).MaxDef
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub

Private Sub picInvUser_Click()
Dim ItemSlot As Byte
    
    ItemSlot = InvComUsu.SelectedItem
    
    If ItemSlot = 0 Then Exit Sub
    
    ClickNpcInv = False
    InvComNpc.DeselectItem
    
    Label1(0).Caption = Inventario.ItemName(ItemSlot)
    Label1(1).Caption = "Precio : " & CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text)) 'No mostramos numeros reales
    
    If Inventario.amount(ItemSlot) <> 0 Then
    
        Select Case Inventario.OBJType(ItemSlot)
            Case eObjType.otWeapon
                Label1(2).Caption = "Hit: " & Inventario.MinHit(ItemSlot) & "/" & Inventario.MaxHit(ItemSlot)
                Label1(3).Caption = ""
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Def: " & Inventario.MinDef(ItemSlot) & "/" & Inventario.MaxDef(ItemSlot)
                Label1(3).Caption = ""
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub
Private Sub PicInvNpc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Position As Integer
Dim i As Integer
Dim file_path As String
Dim data() As Byte
Dim bmpInfo As BITMAPINFO
Dim handle As Integer
Dim bmpData As StdPicture
Dim Last_I As Long
If (Button = vbRightButton) Then

    If InvComNpc.GrhIndex(InvComNpc.SelectedItem) > 0 Then

        Last_I = InvComNpc.SelectedItem
        If Last_I > 0 And Last_I <= MAX_NPC_INVENTORY_SLOTS Then
                    
            Position = BuscarI(InvComNpc.GrhIndex(InvComNpc.SelectedItem))
            
            If Position = 0 Then
                i = GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum
                Call Get_Bitmapp(DirGraficos, CStr(GrhData(InvComNpc.GrhIndex(InvComNpc.SelectedItem)).FileNum) & ".BMP", bmpInfo, data)
                Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                ImageList1.ListImages.Add , CStr("g" & InvComNpc.GrhIndex(InvComNpc.SelectedItem)), Picture:=bmpData
                Position = ImageList1.ListImages.Count
                Set bmpData = Nothing
            End If
            
            
          '  InvComNpc.uMoveItem = True
            
            Set picInvNpc.MouseIcon = ImageList1.ListImages(Position).ExtractIcon
            picInvNpc.MousePointer = vbCustom

            Exit Sub
        End If
    End If
End If
End Sub
 
Private Sub picInvNpc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
' @Wildem
 
'Pongo el puntero por default primero.
picInvNpc.MousePointer = vbDefault
 
If x > 0 And x < picInvNpc.ScaleWidth And y > 0 And y < picInvNpc.ScaleHeight Then
 
    'mmm, dejo por si alguien quiere agregarle algo (?
 
Else
    ' Debe tener seleccionado un item para comprarlo.
    If InvComNpc.SelectedItem = 0 Then Exit Sub
   
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
   
    Call Audio.PlayWave(SND_CLICK)
   
    LasActionBuy = True
   
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficiente oro.", 2, 51, 223, 1, 1)
        Exit Sub
    End If
   
    picInvNpc.MousePointer = vbDefault
End If
InvComNpc.DrawInventory
InvComUsu.DrawInventory
 InvComNpc.sMoveItem = False
 InvComNpc.uMoveItem = False
End Sub
 
Private Function BuscarI(gh As Integer) As Integer
Dim i As Integer
 
For i = 1 To frmComerciar.ImageList1.ListImages.Count
    If frmComerciar.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
 
End Function
 
Private Sub PicInvUser_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Position As Integer
Dim i As Integer
Dim file_path As String
Dim data() As Byte
Dim bmpInfo As BITMAPINFO
Dim handle As Integer
Dim bmpData As StdPicture
Dim Last_I As Long
If (Button = vbRightButton) Then

    If InvComUsu.GrhIndex(InvComUsu.SelectedItem) > 0 Then

        Last_I = InvComUsu.SelectedItem
        If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
                    
            Position = BuscarI(InvComUsu.GrhIndex(InvComUsu.SelectedItem))
            
            If Position = 0 Then
                i = GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum
                Call Get_Bitmapp(DirGraficos, CStr(GrhData(InvComUsu.GrhIndex(InvComUsu.SelectedItem)).FileNum) & ".BMP", bmpInfo, data)
                Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                ImageList1.ListImages.Add , CStr("g" & InvComUsu.GrhIndex(InvComUsu.SelectedItem)), Picture:=bmpData
                Position = ImageList1.ListImages.Count
                Set bmpData = Nothing
            End If
            
            
          '  InvComUsu.uMoveItem = True
            
            Set picInvUser.MouseIcon = ImageList1.ListImages(Position).ExtractIcon
            picInvUser.MousePointer = vbCustom

            Exit Sub
        End If
    End If
End If
End Sub

 
Private Sub picInvUser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
'Pongo el puntero por default primero.
picInvUser.MousePointer = vbDefault
 
If Not x > 0 And x < picInvUser.ScaleWidth And y > 0 And y < picInvUser.ScaleHeight Then
    ' Debe tener seleccionado un item para comprarlo.
    If InvComUsu.SelectedItem = 0 Then Exit Sub
 
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
   
    Call Audio.PlayWave(SND_CLICK)
   
    LasActionBuy = False
 
    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
   
    picInvUser.MousePointer = vbDefault
End If
InvComNpc.DrawInventory
InvComUsu.DrawInventory
InvComUsu.sMoveItem = False
InvComUsu.uMoveItem = False
 
End Sub
