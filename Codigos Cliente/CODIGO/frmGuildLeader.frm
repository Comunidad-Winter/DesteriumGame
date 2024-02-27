VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildLeader.frx":0000
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ImgElecciones 
      Caption         =   "Abrir elección"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtguildnews 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   5475
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      ItemData        =   "frmGuildLeader.frx":81A92
      Left            =   240
      List            =   "frmGuildLeader.frx":81A94
      TabIndex        =   2
      Top             =   4305
      Width           =   2595
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":81A96
      Left            =   3120
      List            =   "frmGuildLeader.frx":81A98
      TabIndex        =   1
      Top             =   480
      Width           =   2595
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      ItemData        =   "frmGuildLeader.frx":81A9A
      Left            =   180
      List            =   "frmGuildLeader.frx":81A9C
      TabIndex        =   0
      Top             =   480
      Width           =   2595
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "El clan cuenta con x miembros"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   3840
      Tag             =   "1"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgPropuestasPaz 
      Height          =   375
      Left            =   3000
      Tag             =   "1"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Image imgEditarURL 
      Height          =   375
      Left            =   3000
      Tag             =   "1"
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Image imgEditarCodex 
      Height          =   375
      Left            =   3120
      Tag             =   "1"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Image imgActualizar 
      Height          =   390
      Left            =   2400
      Tag             =   "1"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image imgDetallesSolicitudes 
      Height          =   375
      Left            =   960
      Tag             =   "1"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Image imgDetallesMiembros 
      Height          =   375
      Left            =   3840
      Tag             =   "1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgDetallesClan 
      Height          =   375
      Left            =   840
      Tag             =   "1"
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmGuildLeader"
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

Private Const MAX_NEWS_LENGTH As Integer = 512
Private clsFormulario As clsFormMovementManager

Private cBotonElecciones As clsGraphicalButton
Private cBotonActualizar As clsGraphicalButton
Private cBotonDetallesClan As clsGraphicalButton
Private cBotonDetallesMiembros As clsGraphicalButton
Private cBotonDetallesSolicitudes As clsGraphicalButton
Private cBotonEditarCodex As clsGraphicalButton
Private cBotonEditarURL As clsGraphicalButton
Private cBotonPropuestasPaz As clsGraphicalButton
Private cBotonPropuestasAlianzas As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
'    Me.Picture = LoadPicture(App.path & "\graficos\VentanaAdministrarClan.jpg")
    
 
End Sub


Private Sub imgActualizar_Click()
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)
End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub imgDetallesClan_Click()
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub imgDetallesMiembros_Click()
    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))
End Sub

Private Sub imgDetallesSolicitudes_Click()
    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))
End Sub

Private Sub imgEditarCodex_Click()
    Call frmGuildDetails.Show(vbModal, frmGuildLeader)
End Sub

Private Sub imgEditarURL_Click()
    Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub

Private Sub imgElecciones_Click()
    Call WriteGuildOpenElections
    Unload Me
End Sub


Private Sub ImgEleccion_Click()

End Sub

Private Sub imgPropuestasPaz_Click()
    Call WriteGuildPeacePropList
End Sub


Private Sub txtguildnews_Change()
    If Len(txtguildnews.Text) > MAX_NEWS_LENGTH Then _
        txtguildnews.Text = Left$(txtguildnews.Text, MAX_NEWS_LENGTH)
End Sub


