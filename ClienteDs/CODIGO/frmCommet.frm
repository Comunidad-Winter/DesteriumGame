VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz o alianza"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5055
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
   Picture         =   "frmCommet.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
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
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   4575
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   840
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Image imgEnviar 
      Height          =   480
      Left            =   3120
      Top             =   2640
      Width           =   1200
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Desterium AO 0.11.6
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
'Desterium AO is based on Baronsoft's VB6 Online RPG
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
Private Const MAX_PROPOSAL_LENGTH As Integer = 520

Private cBotonEnviar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public Nombre As String

Public t As Tipo

Public Enum Tipo
    ALIANZA = 1
    PAZ = 2
    RECHAZOPJ = 3
End Enum

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me

30        Call LoadBackGround
40        Call LoadButtons
End Sub

Private Sub LoadButtons()
          Dim GrhPath As String
          
10        GrhPath = DirGraficos

20        Set cBotonEnviar = New clsGraphicalButton
30        Set cBotonCerrar = New clsGraphicalButton
          
40        Set LastPressed = New clsGraphicalButton
          
          
         ' Call cBotonEnviar.Initialize(imgEnviar, GrhPath & "BotonEnviarSolicitud.jpg", GrhPath & "BotonEnviarRolloverSolicitud.jpg", GrhPath & "BotonEnviarClickSolicitud.jpg", Me)

          'Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarSolicitud.jpg", GrhPath & "BotonCerrarRolloverSolicitud.jpg", GrhPath & "BotonCerrarClickSolicitud.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
10        Unload Me
End Sub

Private Sub imgEnviar_Click()

10        If text1 = "" Then
20            If t = PAZ Or t = ALIANZA Then
30                MsgBox _
                      "Debes redactar un mensaje solicitando la paz o alianza al líder de " _
                      & Nombre
40            Else
50                MsgBox _
                      "Debes indicar el motivo por el cual rechazas la membresía de " & _
                      Nombre
60            End If
              
70            Exit Sub
80        End If
          
90        If t = PAZ Then
100           Call WriteGuildOfferPeace(Nombre, Replace(text1, vbCrLf, "º"))
              
110       ElseIf t = ALIANZA Then
120           Call WriteGuildOfferAlliance(Nombre, Replace(text1, vbCrLf, "º"))
              
130       ElseIf t = RECHAZOPJ Then
140           Call WriteGuildRejectNewMember(Nombre, Replace(Replace(text1.Text, ",", _
                  " "), vbCrLf, " "))
              'Sacamos el char de la lista de aspirantes
              Dim i As Long
              
150           For i = 0 To frmGuildLeader.solicitudes.ListCount - 1
160               If frmGuildLeader.solicitudes.List(i) = Nombre Then
170                   frmGuildLeader.solicitudes.RemoveItem i
180                   Exit For
190               End If
200           Next i
              
210           Me.Hide
220           Unload frmCharInfo
230       End If
          
240       Unload Me

End Sub

Private Sub Text1_Change()
10        If Len(text1.Text) > MAX_PROPOSAL_LENGTH Then text1.Text = Left$(text1.Text, _
              MAX_PROPOSAL_LENGTH)
End Sub

Private Sub LoadBackGround()

10        Select Case t
              Case Tipo.ALIANZA
20                Me.Picture = LoadPicture(DirGraficos & _
                      "VentanaPropuestaAlianza.jpg")
                  
30            Case Tipo.PAZ
40                Me.Picture = LoadPicture(DirGraficos & "VentanaPropuestaPaz.jpg")
                  
50            Case Tipo.RECHAZOPJ
60                Me.Picture = LoadPicture(DirGraficos & "VentanaMotivoRechazo.jpg")
                  
70        End Select
          
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
10        LastPressed.ToggleToNormal
End Sub
