VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5715
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
   ScaleHeight     =   381
   ScaleMode       =   0  'User
   ScaleWidth      =   411
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   7
      Text            =   "1"
      Top             =   7320
      Width           =   1035
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
      TabIndex        =   6
      Text            =   "1"
      Top             =   3015
      Width           =   495
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   2400
      ScaleWidth      =   3870
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   3870
   End
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   1890
      ScaleHeight     =   10.659
      ScaleMode       =   0  'User
      ScaleWidth      =   996.129
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   393216
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
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
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
      Index           =   6
      Left            =   4920
      TabIndex        =   11
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
      Index           =   5
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
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
      Index           =   4
      Left            =   4920
      TabIndex        =   9
      Top             =   1800
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
      Index           =   3
      Left            =   885
      TabIndex        =   8
      Top             =   4200
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
      TabIndex        =   5
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
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   4920
      Tag             =   "0"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   5565
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   5550
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   375
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
      TabIndex        =   2
      Top             =   4800
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
      Index           =   2
      Left            =   885
      TabIndex        =   1
      Top             =   5400
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
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3600
      Width           =   90
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
10        If Val(cantidad.Text) < 1 Then
20            cantidad.Text = 1
30        End If
          
40        If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
50            cantidad.Text = MAX_INVENTORY_OBJS
60        End If
          Dim ItemSlot As Byte
70        ItemSlot = InvComNpc.SelectedItem
          
80        If ClickNpcInv Then
90        Label1(0).Caption = "Precio: " & _
              CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text)) 'No _
              mostramos numeros reales
100       Else
110           If InvComUsu.SelectedItem <> 0 Then
120               Label1(0).Caption = "Precio: " & _
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

Private Sub CantidadOro_Change()
10        If Val(CantidadOro.Text) < 1 Then
20            cantidad.Text = 1
30        End If
End Sub

Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
10        If (KeyAscii <> 8) Then
20            If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
30                KeyAscii = 0
40            End If
50        End If
End Sub

Private Sub Form_Load()
       
          'Cargamos la interfase
          'Me.Picture = LoadPicture(App.path & "\Recursos\Boveda.jpg")
         
10        Call LoadButtons
         
End Sub

Private Sub LoadButtons()

          Dim GrhPath As String
          
10        GrhPath = DirGraficos
          'CmdMoverBov(1).Picture = LoadPicture(App.path & "\Recursos\FlechaSubirObjeto.jpg") ' www.gs-zone.org
          'CmdMoverBov(0).Picture = LoadPicture(App.path & "\Recursos\FlechaBajarObjeto.jpg") ' www.gs-zone.org
          
20        Set cBotonRetirarOro = New clsGraphicalButton
30        Set cBotonDepositarOro = New clsGraphicalButton
40        Set cBotonCerrar = New clsGraphicalButton
          
50        Set LastPressed = New clsGraphicalButton


          'Call cBotonDepositarOro.Initialize(imgDepositarOro, "", GrhPath & "BotonDepositaOroApretado.jpg", GrhPath & "BotonDepositaOroApretado.jpg", Me)
          'Call cBotonRetirarOro.Initialize(imgRetirarOro, "", GrhPath & "BotonRetirarOroApretado.jpg", GrhPath & "BotonRetirarOroApretado.jpg", Me)
          'Call cBotonCerrar.Initialize(imgCerrar, "", GrhPath & "xPrendida.bmp", GrhPath & "xPrendida.bmp", Me)
          
60        Image1(0).MouseIcon = picMouseIcon
70        Image1(1).MouseIcon = picMouseIcon
          
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        Call LastPressed.ToggleToNormal
End Sub

Private Sub Image1_Click(Index As Integer)
          
10        Call Audio.PlayWave(SND_CLICK)
          
20        If InvBanco(Index).SelectedItem = 0 Then Exit Sub
          
30        If Not IsNumeric(cantidad.Text) Then Exit Sub
          
40        Select Case Index
              Case 0
50                LastIndex1 = InvBanco(0).SelectedItem
60                LasActionBuy = True
70                Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
                  
80           Case 1
90                LastIndex2 = InvBanco(1).SelectedItem
100               LasActionBuy = False
110               Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
120       End Select

End Sub


Private Sub imgDepositarOro_Click()
10        Call WriteBankDepositGold(Val(CantidadOro.Text))
End Sub

Private Sub imgRetirarOro_Click()
10        Call WriteBankExtractGold(Val(CantidadOro.Text))
End Sub

Private Sub PicBancoInv_Click()
10        If InvBanco(0).SelectedItem <> 0 Then
          
20            With UserBancoInventory(InvBanco(0).SelectedItem)
30                Label1(6).Caption = .Name
                  
40                Select Case .ObjType
                      Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, _
                          18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, _
                          34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, _
                          50, 51, 52, 53, 43, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64
50                        Label1(4).Caption = "" & .MaxHit
60                        Label1(5).Caption = "" & .MinHit
70                        Label1(7).Caption = "" & .MinDef & " / " & .MaxDef
80                        Label1(4).Visible = True
90                        Label1(5).Visible = True
100                       Label1(7).Visible = True
                          
110                    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, _
                           18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, _
                           34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, _
                           50, 51, 52, 53, 43, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64
120                       Label1(7).Caption = "" & .MinDef & " / " & .MaxDef
130                       Label1(7).Visible = True
                          
140                   Case Else
150                       Label1(4).Visible = False
160                       Label1(5).Visible = False
170                       Label1(7).Visible = False
                          
180               End Select
                  
190           End With
              
200       Else
210           Label1(6).Caption = ""
220           Label1(4).Visible = False
230           Label1(5).Visible = False
240           Label1(7).Visible = False
250       End If

End Sub
Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
      Dim Position As Integer
      Dim i As Long
      Dim file_path As String
      Dim data() As Byte
      Dim bmpInfo As BITMAPINFO
      Dim handle As Integer
      Dim bmpData As StdPicture

10    If (Button = vbRightButton) Then

20     If InvBanco(1).GrhIndex(InvBanco(1).SelectedItem) > 0 Then
30            Last_I = InvBanco(1).SelectedItem
40            If Last_I > 0 And Last_I <= MAX_INVENTORY_SLOTS Then
             
                 
50                Position = _
                      Search_GhID(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem))
                  
60                If Position = 0 Then
70                    i = _
                          GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum
80                    Call Get_Bitmapp(DirGraficos, _
                          CStr(GrhData(InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)).FileNum) _
                          & ".BMP", bmpInfo, data)
90                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
100                   frmBancoObj.ImageList1.ListImages.Add , CStr("g" & _
                          InvBanco(1).GrhIndex(InvBanco(1).SelectedItem)), _
                          Picture:=bmpData
110                   Position = frmBancoObj.ImageList1.ListImages.Count
120                   Set bmpData = Nothing
130               End If
                  
                 
140                   Set PicInv.MouseIcon = _
                          frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
150               frmBancoObj.PicInv.MousePointer = vbCustom
       
160               Exit Sub
170           End If
180     End If
       
190   End If

End Sub

Private Sub PicInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
       
      'Pongo el puntero por default primero.
10    PicInv.MousePointer = vbDefault
       
20    If X > 0 And X < PicInv.ScaleWidth And Y > 0 And Y < PicInv.ScaleHeight Then

30        If InvBanco(1).SelectedItem = 0 Then Exit Sub
         
40        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
         
50        Call Audio.PlayWave(SND_CLICK)

60    Else
70        Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
80        PicInv.MousePointer = vbDefault
          
90        InvBanco(1).DrawInventory
100   InvBanco(1).DrawInventory
110    InvBanco(1).sMoveItem = False
120    InvBanco(1).uMoveItem = False
          
130   End If
End Sub
Private Function Search_GhID(gh As Integer) As Integer

      Dim i As Integer

10    For i = 1 To frmBancoObj.ImageList1.ListImages.Count
20        If frmBancoObj.ImageList1.ListImages(i).key = "g" & CStr(gh) Then
30            Search_GhID = i
40            Exit For
50        End If
60    Next i

End Function
Private Sub PicInv_Click()

10        If InvBanco(1).SelectedItem <> 0 Then
20            With Inventario
30                Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
                  
40                Select Case .ObjType(InvBanco(1).SelectedItem)
                      Case eOBJType.otUseOnce, eOBJType.otWeapon, eOBJType.otArmadura

50                        Label1(1).Caption = "" & .MaxHit(InvBanco(1).SelectedItem)
60                        Label1(2).Caption = "" & .MinHit(InvBanco(1).SelectedItem)
70                        Label1(3).Caption = "" & .MaxDef(InvBanco(1).SelectedItem)
80                        Label1(1).Visible = True
90                        Label1(2).Visible = True
100                       Label1(3).Visible = True
                          
110                   Case eOBJType.otUseOnce, eOBJType.otWeapon, eOBJType.otArmadura
120                       Label1(3).Caption = "" & .MaxDef(InvBanco(1).SelectedItem)
130                       Label1(3).Visible = True
                          
140                   Case Else
150                       Label1(1).Visible = False
160                       Label1(2).Visible = False
170                       Label1(3).Visible = False
                          
180               End Select
                  
190           End With
              
200       Else
210           Label1(0).Caption = ""
220           Label1(1).Visible = False
230           Label1(2).Visible = False
240           Label1(3).Visible = False
250       End If
End Sub


Private Sub imgCerrar_Click()
10        Call WriteBankEnd
20        NoPuedeMover = False
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


Private Function BuscarI(gh As Integer) As Integer
      Dim i As Long
       
10    For i = 1 To frmBancoObj.ImageList1.ListImages.Count
20        If frmBancoObj.ImageList1.ListImages(i).key = "g" & CStr(gh) Then
30            BuscarI = i
40            Exit For
50        End If
60    Next i
       
End Function
Private Sub PicBancoInv_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      Dim Position As Integer
      Dim i As Long
      Dim file_path As String
      Dim data() As Byte
      Dim bmpInfo As BITMAPINFO
      Dim handle As Integer
      Dim bmpData As StdPicture

10    If (Button = vbRightButton) Then

20     If InvBanco(0).GrhIndex(InvBanco(0).SelectedItem) > 0 Then
30            Last_I = InvBanco(0).SelectedItem
40            If Last_I > 0 And Last_I <= MAX_BANCOINVENTORY_SLOTS Then
             
                 
50                Position = _
                      Search_GhID(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem))
                  
60                If Position = 0 Then
70                    i = _
                          GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum
80                    Call Get_Bitmapp(DirGraficos, _
                          CStr(GrhData(InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)).FileNum) _
                          & ".BMP", bmpInfo, data)
90                    Set bmpData = ArrayToPicture(data(), 0, UBound(data) + 1) ' GSZAO ' GSZAO
100                   frmBancoObj.ImageList1.ListImages.Add , CStr("g" & _
                          InvBanco(0).GrhIndex(InvBanco(0).SelectedItem)), _
                          Picture:=bmpData
110                   Position = frmBancoObj.ImageList1.ListImages.Count
120                   Set bmpData = Nothing
130               End If
                  
                 
140                   Set PicBancoInv.MouseIcon = _
                          frmBancoObj.ImageList1.ListImages(Position).ExtractIcon
150               frmBancoObj.PicBancoInv.MousePointer = vbCustom
       
160               Exit Sub
170           End If
180     End If
       
190   End If

End Sub
Private Sub PicBancoInv_MouseUp(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
       
      'Pongo el puntero por default primero.
10    PicBancoInv.MousePointer = vbDefault
       
20    If X > 0 And X < PicBancoInv.ScaleWidth And Y > 0 And Y < _
          PicBancoInv.ScaleHeight Then
          'Acá va la parte donde podemos
          'acomodar los items adentro de la boveda.
          
30            If InvBanco(0).SelectedItem = 0 Then Exit Sub
         
40        If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
         
50        Call Audio.PlayWave(SND_CLICK)
          
60    Else
70        Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
80        PicBancoInv.MousePointer = vbDefault
90    End If

100   InvBanco(0).DrawInventory
110   InvBanco(0).DrawInventory
120    InvBanco(0).sMoveItem = False
130    InvBanco(0).uMoveItem = False
       
End Sub
