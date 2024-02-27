VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   0  'None
   Caption         =   "Misiones"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmQuests.frx":0000
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
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
      Height          =   3330
      Left            =   2145
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   270
      Width           =   2175
   End
   Begin VB.ListBox lstQuests 
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
      ForeColor       =   &H80000004&
      Height          =   2565
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   1755
   End
   Begin VB.Image CmdOptions 
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image CmdOptions 
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private Sub Image1_Click()
10    Unload Me
End Sub

Private Sub CMdOptions_Click(Index As Integer)
10    Call Audio.PlayWave(SND_CLICK)
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el click de los CommandButtons cmdOptions.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
20        Select Case Index
              Case 0 'Botón ABANDONAR MISIÓN
                  'Chequeamos si hay items.
30                If lstQuests.ListCount = 0 Then
40                    MsgBox "¡No tienes ninguna misión!", vbOKOnly + vbExclamation
50                    Exit Sub
60                End If
                  
                  'Chequeamos si tiene algun item seleccionado.
70                If lstQuests.ListIndex < 0 Then
80                    MsgBox "¡Primero debes seleccionar una misión!", vbOKOnly + _
                          vbExclamation
90                    Exit Sub
100               End If
                  
110               Select Case MsgBox("¿Estás seguro que deseas abandonar la misión?", _
                      vbYesNo + vbExclamation)
                      Case vbYes  'Botón SÍ.
                          'Enviamos el paquete para abandonar la quest
120                       Call WriteQuestAbandon(lstQuests.ListIndex + 1)
                          
130                   Case vbNo   'Botón NO.
                          'Como seleccionó que no, no hace nada.
140                       Exit Sub
150               End Select
                  
160           Case 1 'Botón VOLVER
170               Unload Me
180       End Select
End Sub


Private Sub lstQuests_Click()
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
      'Maneja el click del ListBox lstQuests.
      'Last modified: 31/01/2010 by Amraphen
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
10        If lstQuests.ListIndex < 0 Then Exit Sub
          
20        Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)
End Sub


