VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10140
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   960
      Left            =   3705
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   6630
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1693
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":3BFF8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   960
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   9480
      Width           =   9600
   End
   Begin VB.Image barra 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   3510
      Picture         =   "frmCargando.frx":3C076
      Top             =   8580
      Width           =   4770
   End
End
Attribute VB_Name = "frmCargando"
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
'This program is free software;you can redistribute it and/or modify
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
'La Plata - Pcia, Buenos Aiçres - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Dim F As Integer


Private Sub Form_Load()

10    barra.Width = 1
End Sub

Private Sub LOGO_KeyPress(KeyAscii As Integer)
10    Debug.Print 2
End Sub

Private Sub Status_KeyPress(KeyAscii As Integer)
10    Debug.Print 1
End Sub
Function Analizar()
10                On Error Resume Next
                 
                  Dim iX As Integer
                  Dim tX As Integer
                  Dim DifX As Integer
                 
      'LINK1            'Variable que contiene el numero de actualización correcto del servidor
20     '   iX = Inet1.OpenURL("http://ds.porvo.online/Update.txt") 'Host
30      '  tX = LeerInt(App.path & "\INIT\Update.DS")
       
40        If iX <> 0 Then
50            If tX <> iX Then
60                Call _
                      MsgBox("¡¡IMPORTANTE!! Se te cerrará el cliente y abrirá el Update del juego. Recuerda que cada actualización sera subida a la página para que puedas bajar el cliente completo y funcional.")
70                Call ShellExecute(Me.hWnd, "open", App.path & _
                      "/UpdateDesterium.exe", "", "", 1)
80                End
90            End If
100       End If
End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
10        F = FreeFile
20        Open Ruta For Input As F
30        LeerInt = Input$(LOF(F), #F)
40        Close #F
End Function
 
Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
10        F = FreeFile
20        Open Ruta For Output As F
30        Print #F, data
40        Close #F
End Sub
