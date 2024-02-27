VERSION 5.00
Begin VB.Form frmBuscar 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6660
      TabIndex        =   13
      Text            =   "9999"
      Top             =   1140
      Width           =   750
   End
   Begin VB.CommandButton Limpiarlistas 
      Caption         =   "Limpiar Listas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   7335
   End
   Begin VB.ListBox ListCrearNpcs 
      Height          =   3180
      ItemData        =   "frmBuscar.frx":0000
      Left            =   960
      List            =   "frmBuscar.frx":0002
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.ListBox ListCrearObj 
      Height          =   3180
      ItemData        =   "frmBuscar.frx":0004
      Left            =   120
      List            =   "frmBuscar.frx":0006
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frmBuscar.frx":0008
      Left            =   1800
      List            =   "frmBuscar.frx":000A
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar NPCs"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar Objetos."
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
      Left            =   2400
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox NPCs 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Ingrese Nombre de NPC(ej: Lilith)"
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox Objetos 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Ingrese nombre de objeto(ej: espada)"
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   5700
      TabIndex        =   14
      Top             =   1140
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBuscar.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   135
      TabIndex        =   12
      Top             =   5895
      Width           =   7320
   End
   Begin VB.Label CrearObjetos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Objetos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label CrearNPCs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NPCS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Crear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1905
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu mnuCrearO 
      Caption         =   "Crear Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearObj 
         Caption         =   "¿Crear Objeto?"
      End
   End
   Begin VB.Menu mnuCrearN 
      Caption         =   "Crear NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNPC 
         Caption         =   "¿Crear NPC?"
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    ListCrearNpcs.Clear
    List1.Clear
    ListCrearObj.Clear

    If Not Objetos.Text = vbNullString Then
        Call WriteSearchObj(Objetos.Text)

    End If

End Sub

Private Sub Command2_Click()

    ListCrearNpcs.Clear
    List1.Clear
    ListCrearObj.Clear

    If Not NPCs.Text = vbNullString Then
        Call WriteSearchNpc(NPCs.Text)

    End If

End Sub

Private Sub Limpiarlistas_Click()

    ListCrearNpcs.Clear
    List1.Clear
    ListCrearObj.Clear

End Sub

Private Sub ListCrearNpcs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopUpMenu mnuCrearN

    End If

End Sub

Private Sub ListCrearObj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopUpMenu mnuCrearO

    End If

End Sub

Private Sub mnuCrearObj_Click()

    If txtCant.Text = "0" Then
        MsgBox "Cantidad Incorrecta."
        Exit Sub

    End If

    If ListCrearObj.ListIndex >= 0 Then

        If ListCrearObj.Visible Then
           
            If txtCant.Text = "" Then txtCant.Text = "1"
            
            WriteCreateItem ListCrearObj.Text, txtCant.Text = "1"

        End If

    Else
        MsgBox "Error, seleccione un OBJETO."

    End If

End Sub

Private Sub mnuCrearNPC_Click()

    If ListCrearNpcs.ListIndex >= 0 Then

        If ListCrearNpcs.Visible Then
            Call ParseUserCommand("/ACC " & ListCrearNpcs.Text)

        End If

    Else
        MsgBox "Error, seleccione un NPC."

    End If

End Sub

