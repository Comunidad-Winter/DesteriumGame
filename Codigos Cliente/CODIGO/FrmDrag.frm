VERSION 5.00
Begin VB.Form FrmDrag 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   1680
   ClientTop       =   4455
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Todo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MousePointer    =   99  'Custom
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1035
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&ceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      MousePointer    =   99  'Custom
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la cantidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   585
      TabIndex        =   0
      Top             =   165
      Width           =   2415
   End
End
Attribute VB_Name = "FrmDrag"
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
 If LenB(FrmDrag.text1) > 0 Then
        If Not IsNumeric(FrmDrag.text1) Then Exit Sub  'Should never happen
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
        FrmDrag.text1.Text = ""
        
        Else
        Call WriteDrop(Inventario.SelectedItem, FrmDrag.text1.Text)
        FrmDrag.text1.Text = ""
    End If
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
If Inventario.SelectedItem = 0 Then Exit Sub
    ' If LenB(FrmDrag.Text1) > 0 Then
       ' If Not IsNumeric(FrmDrag.Text1) Then Exit Sub  'Should never happen
        If Inventario.sMoveItem Then 'drag and drop
X = frmMain.MouseX
Y = frmMain.MouseY
  ConvertCPtoTP X, Y, tX, tY
        WriteDragToPos tX, tY, Inventario.SelectedItem, Inventario.amount(Inventario.SelectedItem)
        FrmDrag.text1.Text = ""
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
    FrmDrag.text1.Text = ""
    
    Unload Me
End Sub

Private Sub Text1_Change()
On Error GoTo ErrHandler
    If Val(FrmDrag.text1) < 1 Then
        FrmDrag.text1 = "1"
    End If
    
    If Val(FrmDrag.text1) > 100000 Then
        FrmDrag.text1 = "100000"
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    FrmDrag.text1 = "1"
End Sub

