VERSION 5.00
Begin VB.Form frmCarp 
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   4035
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmCarp.frx":0000
   ScaleHeight     =   4035
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
      Height          =   2160
      Left            =   360
      TabIndex        =   0
      Top             =   1220
      Width           =   3840
   End
   Begin VB.Image Command4 
      Height          =   375
      Left            =   2640
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Command3 
      Height          =   375
      Left            =   240
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmCarp"
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

Private Sub Command3_Click()
10        On Error Resume Next

20        Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1))
30        If frmMain.macrotrabajo.Enabled Then MacroBltIndex = _
              ObjCarpintero(lstArmas.ListIndex + 1)
          
40        Unload Me
End Sub

Private Sub Command4_Click()
10        Unload Me
End Sub
