VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCantidad.frx":0000
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2625
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   120
      Tag             =   "1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmCantidad"
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

Private Sub Command1_Click()
 If LenB(frmCantidad.text1) > 0 Then
        If Not IsNumeric(frmCantidad.text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then
        If text1 > Inventario.amount(Inventario.SelectedItem) Then
        ShowConsoleMsg "No tienes esa cantidad!"
        Unload Me
        Exit Sub
        End If
        
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, text1
        frmCantidad.text1.Text = ""
        
        Else
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.text1.Text)
        frmCantidad.text1.Text = ""
    End If
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
If Inventario.SelectedItem = 0 Then Exit Sub
    ' If LenB(frmcantidad.Text1) > 0 Then
       ' If Not IsNumeric(frmcantidad.Text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then 'drag and drop
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem)
        frmCantidad.text1.Text = ""
Unload Me
    
Else 'tirar al piso :D
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem))
        Unload Me
    Else
        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000)
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If
    End If
    frmCantidad.text1.Text = ""
    
    Unload Me
End Sub

Private Sub Text1_Change()
On Error GoTo ErrHandler
    If Val(frmCantidad.text1) < 0 Then
        frmCantidad.text1 = "1"
    End If
    
    If Val(frmCantidad.text1) > 100000 Then
        frmCantidad.text1 = "100000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    frmCantidad.text1 = "1"
End Sub


