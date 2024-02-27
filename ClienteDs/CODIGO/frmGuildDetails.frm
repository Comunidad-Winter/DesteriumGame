VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6885
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
   Picture         =   "frmGuildDetails.frx":0000
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   6120
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   570
      TabIndex        =   1
      Top             =   3525
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   570
      TabIndex        =   2
      Top             =   3885
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   570
      TabIndex        =   3
      Top             =   4245
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   570
      TabIndex        =   4
      Top             =   4605
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   570
      TabIndex        =   5
      Top             =   4965
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   570
      TabIndex        =   6
      Top             =   5325
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   570
      TabIndex        =   7
      Top             =   5685
      Width           =   5625
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   570
      TabIndex        =   8
      Top             =   6045
      Width           =   5625
   End
   Begin VB.Image imgConfirmar 
      Height          =   360
      Left            =   5520
      Tag             =   "1"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgSalir 
      Height          =   360
      Left            =   120
      Tag             =   "1"
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frmGuildDetails"
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

Private cBotonConfirmar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Const MAX_DESC_LENGTH As Integer = 520
Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
          
       '   Me.Picture = LoadPicture(App.path & "\Recursos\VentanaCodex.jpg")
          

End Sub


Private Sub imgConfirmar_Click()
          Dim fdesc As String
          Dim Codex() As String
          Dim k As Byte
          Dim Cont As Byte

10        fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)


20        Cont = 0
30        For k = 0 To txtCodex1.UBound
40            If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
50        Next k
          
60        If Cont < 4 Then
70            MsgBox "Debes definir al menos cuatro mandamientos."
80            Exit Sub
90        End If
                      
100       ReDim Codex(txtCodex1.UBound) As String
110       For k = 0 To txtCodex1.UBound
120           Codex(k) = txtCodex1(k)
130       Next k

140       If CreandoClan Then
150           Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
160       Else
170           Call WriteClanCodexUpdate(fdesc, Codex)
180       End If

190       CreandoClan = False
200       Unload Me
End Sub

Private Sub imgSalir_Click()
10        Unload Me
End Sub

Private Sub txtCodex1_Change(Index As Integer)
10        If Len(txtCodex1.Item(Index).Text) > MAX_CODEX_LENGTH Then _
              txtCodex1.Item(Index).Text = Left$(txtCodex1.Item(Index).Text, _
              MAX_CODEX_LENGTH)
End Sub


Private Sub txtDesc_Change()
10        If Len(txtDesc.Text) > MAX_DESC_LENGTH Then txtDesc.Text = _
              Left$(txtDesc.Text, MAX_DESC_LENGTH)
End Sub
