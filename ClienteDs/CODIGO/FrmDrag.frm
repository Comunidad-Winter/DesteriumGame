VERSION 5.00
Begin VB.Form FrmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1500
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmDrag.frx":0000
   ScaleHeight     =   1500
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   478
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1680
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X As Single
Private Y As Single
Private tX As Byte
Private tY As Byte
Private MouseX As Long
Private MouseY As Long


Private Sub Image1_Click()
10     If LenB(FrmCantidad.Text1) > 0 Then
20            If Not IsNumeric(FrmCantidad.Text1) Then Exit Sub  'Should never happen
30            If Inventario.sMoveItem Then
              
40            If Text1 > Inventario.Amount(Inventario.SelectedItem) Then
50                ShowConsoleMsg "No tienes esa cantidad!", 65, 190, 156, False, False
60            Unload Me
70            Exit Sub
80            End If
              
90              X = frmMain.MouseX
100             Y = frmMain.MouseY
110
                ConvertCPtoTP X, Y, tX, tY
120           WriteDragToPos tX, tY, Inventario.SelectedItem, Text1
130           FrmCantidad.Text1.Text = ""
              
140           Else
150           Call WriteDrop(Inventario.SelectedItem, FrmCantidad.Text1.Text)
160           FrmCantidad.Text1.Text = ""
170       End If
180       End If
190       Unload Me
End Sub

Private Sub Image2_Click()
10    If Inventario.SelectedItem = 0 Then Exit Sub
          ' If LenB(frmcantidad.Text1) > 0 Then
             ' If Not IsNumeric(frmcantidad.Text1) Then Exit Sub  'Should never happen
20            If Inventario.sMoveItem Then 'drag and drop
30    X = frmMain.MouseX
40    Y = frmMain.MouseY
50      ConvertCPtoTP X, Y, tX, tY
60            WriteDragToPos tX, tY, Inventario.SelectedItem, _
                  Inventario.Amount(Inventario.SelectedItem)
70            FrmCantidad.Text1.Text = ""
80    Unload Me
          
90    Else 'tirar al piso :D
100       If Inventario.SelectedItem <> FLAGORO Then
110           Call WriteDrop(Inventario.SelectedItem, _
                  Inventario.Amount(Inventario.SelectedItem))
120           Unload Me
130       Else
140           If UserGLD > 10000 Then
150               Call WriteDrop(Inventario.SelectedItem, 10000)
160               Unload Me
170           Else
180               Call WriteDrop(Inventario.SelectedItem, UserGLD)
190               Unload Me
200           End If
210       End If
220       End If
230       FrmCantidad.Text1.Text = ""
          
240       Unload Me
End Sub

Private Sub Text1_Change()
10    On Error GoTo ErrHandler
20        If Val(FrmCantidad.Text1) < 0 Then
30            FrmCantidad.Text1 = "1"
40        End If
          
50        If Val(Text1.Text) < 1 Then
60            Text1 = 1
70        End If
          
80        If Val(FrmCantidad.Text1) > 100000 Then
90            FrmCantidad.Text1 = "100000"
100       End If
          
110       Exit Sub
          
ErrHandler:
          'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
120       FrmCantidad.Text1 = "1"
End Sub


