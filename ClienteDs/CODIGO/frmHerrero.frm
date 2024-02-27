VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHerrero.frx":0000
   ScaleHeight     =   4020
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   3480
   End
   Begin VB.ListBox lstArmaduras 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   480
      TabIndex        =   1
      Top             =   1550
      Width           =   3495
   End
   Begin VB.Image Command1 
      Height          =   375
      Left            =   240
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Command2 
      Height          =   375
      Left            =   2280
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Command4 
      Height          =   375
      Left            =   2760
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image Command3 
      Height          =   375
      Left            =   240
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmHerrero"
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

Private Sub Command1_Click()
10    lstArmaduras.Visible = False
20    lstArmas.Visible = True
End Sub

Private Sub Command2_Click()
10    lstArmaduras.Visible = True
20    lstArmas.Visible = False
End Sub

Private Sub Command3_Click()
10    On Error Resume Next

20        If lstArmas.Visible Then
30            Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1))
              
40            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = _
                  ArmasHerrero(lstArmas.ListIndex + 1)
50        Else
60            Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1))
              
70            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = _
                  ArmadurasHerrero(lstArmaduras.ListIndex + 1)
80        End If

90        Unload Me
End Sub

Private Sub Command4_Click()
10    Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

