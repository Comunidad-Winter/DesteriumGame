VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3735
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
   Picture         =   "frmGuildAdm.frx":0000
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
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
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   4200
      Width           =   2385
   End
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      ItemData        =   "frmGuildAdm.frx":10899
      Left            =   675
      List            =   "frmGuildAdm.frx":1089B
      TabIndex        =   0
      Top             =   1245
      Width           =   2445
   End
   Begin VB.Image imgDetalles 
      Height          =   495
      Left            =   120
      Tag             =   "1"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   2280
      Tag             =   "1"
      Top             =   4680
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildAdm"
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

Private cBotonCerrar As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()
          ' Handles Form movement (drag and drop).
10        Set clsFormulario = New clsFormMovementManager
20        clsFormulario.Initialize Me
              
       '   Me.Picture = LoadPicture(App.path & "\Recursos\VentanaListaClanes.jpg")
          

          
End Sub



Private Sub imgCerrar_Click()
10        Unload Me
20        frmMain.SetFocus
End Sub

Private Sub imgDetalles_Click()

10        If guildslist.ListIndex = -1 Then
20            MsgBox "Debes seleccionar un clan para ver sus detalles."
30            Exit Sub
40        End If
          
50        If guildslist.List(guildslist.ListIndex) = vbNullString Then
60            MsgBox "Debes seleccionar un clan para ver sus detalles."
70            Exit Sub
80        End If
          
90        frmGuildBrief.EsLeader = False
        
100       Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
End Sub

Private Sub txtBuscar_Change()
10    Call FiltrarListaClanes(txtBuscar.Text)
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

          Dim lIndex As Long
          
10        If UBound(GuildNames) <> 0 Then
20            With guildslist
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

