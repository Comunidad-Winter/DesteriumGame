VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":0000
   ScaleHeight     =   378
   ScaleMode       =   0  'User
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1890
      ScaleHeight     =   10.129
      ScaleMode       =   0  'User
      ScaleWidth      =   996.129
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3225
      Width           =   2895
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   240
      ScaleHeight     =   2520
      ScaleWidth      =   3870
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   3870
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
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
      Left            =   5535
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "1"
      Top             =   3015
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      _Version        =   393216
   End
   Begin VB.TextBox CantidadOro 
      Alignment       =   2  'Center
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
      Height          =   270
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "1"
      Top             =   7320
      Width           =   1035
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
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   3600
      Width           =   90
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
      Index           =   2
      Left            =   885
      TabIndex        =   11
      Top             =   5100
      Visible         =   0   'False
      Width           =   90
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
      Left            =   885
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   5550
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   5565
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   4920
      Tag             =   "0"
      Top             =   5160
      Width           =   1095
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
      Index           =   3
      Left            =   885
      TabIndex        =   9
      Top             =   3990
      Width           =   105
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
      Index           =   4
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
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
      Index           =   5
      Left            =   4920
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   90
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
      Index           =   6
      Left            =   4920
      TabIndex        =   6
      Top             =   585
      Width           =   105
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
      Index           =   7
      Left            =   4920
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Width           =   135
   End
   Begin VB.Image imgDepositarOro 
      Height          =   1050
      Left            =   120
      Tag             =   "0"
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Image imgRetirarOro 
      Height          =   1005
      Left            =   2280
      Tag             =   "0"
      Top             =   6600
      Width           =   1065
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   1
      Left            =   3480
      Top             =   7200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image CmdMoverBov 
      Height          =   375
      Index           =   0
      Left            =   3480
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmBancoObj"
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

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private clsFormulario As clsFormMovementManager

Private cBotonRetirarOro As clsGraphicalButton
Private cBotonDepositarOro As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private Last_I      As Long
Public LastPressed As clsGraphicalButton


Dim Button As Integer

Public Attack As Boolean
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Private ClickNpcInv As Boolean
Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean

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
    Label1(0).Caption = "Precio: " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
    Else
        If InvComUsu.SelectedItem <> 0 Then
            Label1(0).Caption = "Precio: " & CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text))  'No mostramos numeros reales
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

Private Sub CantidadOro_Change()
    If Val(CantidadOro.Text) < 1 Then
        cantidad.Text = 1
    End If
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
 
    'Cargamos la interfase
    'Me.Picture = LoadPicture(App.path & "\Recursos\Boveda.jpg")
   
    Call LoadButtons
   
End Sub

Private Sub LoadButtons()

    Dim GrhPath As String
    
    GrhPath = DirGraficos
    'CmdMoverBov(1).Picture = LoadPicture(App.path & "\Recursos\FlechaSubirObjeto.jpg") ' www.gs-zone.org
    'CmdMoverBov(0).Picture = LoadPicture(App.path & "\Recursos\FlechaBajarObjeto.jpg") ' www.gs-zone.org
    
    Set cBotonRetirarOro = New clsGraphicalButton
    Set cBotonDepositarOro = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton


    'Call cBotonDepositarOro.Initialize(imgDepositarOro, "", GrhPath & "BotonDepositaOroApretado.jpg", GrhPath & "BotonDepositaOroApretado.jpg", Me)
    'Call cBotonRetirarOro.Initialize(imgRetirarOro, "", GrhPath & "BotonRetirarOroApretado.jpg", GrhPath & "BotonRetirarOroApretado.jpg", Me)
    'Call cBotonCerrar.Initialize(imgCerrar, "", GrhPath & "xPrendida.bmp", GrhPath & "xPrendida.bmp", Me)
    
    Image1(0).MouseIcon = picMouseIcon
    Image1(1).MouseIcon = picMouseIcon
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call LastPressed.ToggleToNormal
End Sub

Private Sub Image1_Click(Index As Integer)
    
    Call Audio.PlayWave(SND_CLICK)
    
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    
    Select Case Index
        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
            
       Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    End Select

End Sub


Private Sub imgDepositarOro_Click()
    Call WriteBankDepositGold(Val(CantidadOro.Text))
End Sub

Private Sub imgRetirarOro_Click()
    Call WriteBankExtractGold(Val(CantidadOro.Text))
End Sub

Private Sub PicBancoInv_Click()
    If InvBanco(0).SelectedItem <> 0 Then
    
        With UserBancoInventory(InvBanco(0).SelectedItem)
            Label1(6).Caption = .name
            
            Select Case .OBJType
                Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 43, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64
                    Label1(4).Caption = "" & .MaxHit
                    Label1(5).Caption = "" & .MinHit
                    Label1(7).Caption = "" & .MinDef & " / " & .MaxDef
                    Label1(4).Visible = True
                    Label1(5).Visible = True
                    Label1(7).Visible = True
                    
                 Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 43, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64
                    Label1(7).Caption = "" & .MinDef & " / " & .MaxDef
                    Label1(7).Visible = True
                    
                Case Else
                    Label1(4).Visible = False
                    Label1(5).Visible = False
                    Label1(7).Visible = False
                    
            End Select
            
        End With
        
    Else
        Label1(6).Caption = ""
        Label1(4).Visible = False
        Label1(5).Visible = False
        Label1(7).Visible = False
    End If

End Sub
Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Position As Integer
Dim i As Integer
Dim file_path As String
Dim data() As Byte
Dim bmpInfo As BITMAPINFO
Dim handle As Integer
Dim bmpData As StdPicture

If (Button = vbRightButton) Then

 If InvBanco(1).GrhIndex(InvBanco(1).SelectedItem) > 0 Then
        Last_I = InvBanco(1).SelectedItem
        If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
       
           
            Position = Search_GhID(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem))
            
            If Position = 0 Then
                i = GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum
                Call Get_Bitmapp(DirGraficos, CStr(GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum) & ".BMP", bmpInfo, data)
                Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                frmBancoObj.ImageList1.ListImages.Add , CStr("g" & InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)), Picture:=bmpData
                Position = frmBancoObj.ImageList1.ListImages.Count
                Set bmpData = Nothing
            End If
            
           
                Set PicInv.MouseIcon = frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
            frmBancoObj.PicInv.MousePointer = vbCustom
 
            Exit Sub
        End If
  End If
 
End If

End Sub
Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
'Pongo el puntero por default primero.
PicInv.MousePointer = vbDefault
 
If x > 0 And x < PicInv.ScaleWidth And y > 0 And y < PicInv.ScaleHeight Then

    If InvBanco(1).SelectedItem = 0 Then Exit Sub
   
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
   
    Call Audio.PlayWave(SND_CLICK)

Else
    Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    PicInv.MousePointer = vbDefault
    
    InvBanco(1).DrawInventory
InvBanco(1).DrawInventory
 InvBanco(1).sMoveItem = False
 InvBanco(1).uMoveItem = False
    
End If
End Sub
Private Function Search_GhID(gh As Integer) As Integer

Dim i As Integer

For i = 1 To frmBancoObj.ImageList1.ListImages.Count
    If frmBancoObj.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        Search_GhID = i
        Exit For
    End If
Next i

End Function
Private Sub PicInv_Click()

    If InvBanco(1).SelectedItem <> 0 Then
        With Inventario
            Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
            
            Select Case .OBJType(InvBanco(1).SelectedItem)
                Case eObjType.otUseOnce, eObjType.otWeapon, eObjType.otArmadura

                    Label1(1).Caption = "" & .MaxHit(InvBanco(1).SelectedItem)
                    Label1(2).Caption = "" & .MinHit(InvBanco(1).SelectedItem)
                    Label1(3).Caption = "" & .MaxDef(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                    Label1(3).Visible = True
                    
                Case eObjType.otUseOnce, eObjType.otWeapon, eObjType.otArmadura
                    Label1(3).Caption = "" & .MaxDef(InvBanco(1).SelectedItem)
                    Label1(3).Visible = True
                    
                Case Else
                    Label1(1).Visible = False
                    Label1(2).Visible = False
                    Label1(3).Visible = False
                    
            End Select
            
        End With
        
    Else
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub


Private Sub imgCerrar_Click()
    Call WriteBankEnd
    NoPuedeMover = False
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


Private Function BuscarI(gh As Integer) As Integer
Dim i As Long
 
For i = 1 To frmBancoObj.ImageList1.ListImages.Count
    If frmBancoObj.ImageList1.ListImages(i).Key = "g" & CStr(gh) Then
        BuscarI = i
        Exit For
    End If
Next i
 
End Function
Private Sub PicBancoInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Position As Integer
Dim i As Integer
Dim file_path As String
Dim data() As Byte
Dim bmpInfo As BITMAPINFO
Dim handle As Integer
Dim bmpData As StdPicture

If (Button = vbRightButton) Then

 If InvBanco(0).GrhIndex(InvBanco(0).SelectedItem) > 0 Then
        Last_I = InvBanco(0).SelectedItem
        If Last_I > 0 And Last_I <= MAX_BANCOINVENTORY_SLOTS Then
       
           
            Position = Search_GhID(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem))
            
            If Position = 0 Then
                i = GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum
                Call Get_Bitmapp(DirGraficos, CStr(GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum) & ".BMP", bmpInfo, data)
                Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
                frmBancoObj.ImageList1.ListImages.Add , CStr("g" & InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)), Picture:=bmpData
                Position = frmBancoObj.ImageList1.ListImages.Count
                Set bmpData = Nothing
            End If
            
           
                Set PicBancoInv.MouseIcon = frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
            frmBancoObj.PicBancoInv.MousePointer = vbCustom
 
            Exit Sub
        End If
  End If
 
End If

End Sub
 
Private Sub PicBancoInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
'Pongo el puntero por default primero.
PicBancoInv.MousePointer = vbDefault
 
If x > 0 And x < PicBancoInv.ScaleWidth And y > 0 And y < PicBancoInv.ScaleHeight Then
    'Acá va la parte donde podemos
    'acomodar los items adentro de la boveda.
    
        If InvBanco(0).SelectedItem = 0 Then Exit Sub
   
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
   
    Call Audio.PlayWave(SND_CLICK)
    
Else
    Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
    PicBancoInv.MousePointer = vbDefault
End If

InvBanco(0).DrawInventory
InvBanco(0).DrawInventory
 InvBanco(0).sMoveItem = False
 InvBanco(0).uMoveItem = False
 
End Sub
