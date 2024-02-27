VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   Picture         =   "frmGuildMember.frx":0000
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMiembros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   3075
      TabIndex        =   3
      Top             =   675
      Width           =   2610
   End
   Begin VB.ListBox lstClanes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   195
      TabIndex        =   2
      Top             =   690
      Width           =   2610
   End
   Begin VB.TextBox txtSearch 
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
      Height          =   225
      Left            =   225
      TabIndex        =   1
      Top             =   3630
      Width           =   2550
   End
   Begin VB.Label ImgCerrar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar ventana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      Top             =   5040
      Width           =   2520
   End
   Begin VB.Label imgNoticias 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ver noticias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   2520
   End
   Begin VB.Label imgDetalles 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ver detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   2520
   End
   Begin VB.Label lblCantMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   4635
      TabIndex        =   0
      Top             =   3510
      Width           =   360
   End
   Begin VB.Image imgCerrar1 
      Height          =   495
      Left            =   3000
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgNoticias1 
      Height          =   495
      Left            =   150
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image imgDetalles11 
      Height          =   375
      Left            =   150
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmGuildMember"
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

Private cBotonNoticias As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()

          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me

          'Me.Picture = LoadPicture(DirGraficos & "VentanaMiembroClan.jpg")
          
         ' 'Call Loadbuttons
          
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos

20        Set cBotonNoticias = New clsGraphicalButton
30        Set cBotonDetalles = New clsGraphicalButton
40        Set cBotonCerrar = New clsGraphicalButton
          
50        Set LastPressed = New clsGraphicalButton
          
          
60        Call cBotonDetalles.Initialize(imgDetalles, GrhPath & _
              "BotonDetallesMiembroClan.jpg", GrhPath & _
              "BotonDetallesRolloverMiembroClan.jpg", GrhPath & _
              "BotonDetallesClickMiembroClan.jpg", Me)

70        Call cBotonNoticias.Initialize(imgNoticias, GrhPath & _
              "BotonNoticiasMiembroClan.jpg", GrhPath & _
              "BotonNoticiasRolloverMiembroClan.jpg", GrhPath & _
              "BotonNoticiasClickMiembroClan.jpg", Me)

80        Call cBotonCerrar.Initialize(ImgCerrar, GrhPath & _
              "BotonCerrarMimebroClan.jpg", GrhPath & _
              "BotonCerrarRolloverMimebroClan.jpg", GrhPath & _
              "BotonCerrarClickMimebroClan.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
    'LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
10        Unload Me
End Sub

Private Sub imgDetalles_Click()
10        If lstClanes.ListIndex = -1 Then Exit Sub
          
20        frmGuildBrief.EsLeader = False

30        Call WriteGuildRequestDetails(lstClanes.List(lstClanes.ListIndex))
End Sub

Private Sub imgNoticias_Click()
10        Call WriteShowGuildNews
End Sub

Private Sub txtSearch_Change()
10        Call FiltrarListaClanes(txtSearch.Text)
End Sub

Private Sub txtSearch_GotFocus()
10        With txtSearch
20            .SelStart = 0
30            .SelLength = Len(.Text)
40        End With
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

          Dim lIndex As Long
          
10        If UBound(GuildNames) <> 0 Then
20            With lstClanes
                  'Limpio la lista
30                .Clear
                  
40                .Visible = False
                  
                  ' Recorro los arrays
50                For lIndex = 0 To UBound(GuildNames)
                      ' Si coincide con los patrones
60                    If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                          ' Lo agrego a la lista
70                        .AddItem GuildNames(lIndex)
80                    End If
90                Next lIndex
                  
100               .Visible = True
110           End With
120       End If

End Sub

