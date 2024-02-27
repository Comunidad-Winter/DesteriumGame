VERSION 5.00
Begin VB.Form PanelUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "©Fusion Argentum"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton RECUPERAR 
      Caption         =   "Recuperar Usuario"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton BORRAR 
      Caption         =   "Borrar Usuario"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "PanelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
10    Unload Me

      'MsgBox "Atención con esta acción va a eliminar el personaje, no podra volver a usarlo"
End Sub

Private Sub RECUPERAR_Click()
       
10    EstadoLogin = E_MODO.RecuperarPJ
       
#If UsarWrench = 1 Then
20        If frmMain.Socket1.Connected Then
30            frmMain.Socket1.Disconnect
40            frmMain.Socket1.Cleanup
50            DoEvents
60        End If
70        frmMain.Socket1.HostName = CurServerIp
80        frmMain.Socket1.RemotePort = CurServerPort
90        frmMain.Socket1.Connect
#Else
100       If frmMain.Winsock1.State <> sckClosed Then
110           frmMain.Winsock1.Close
120           DoEvents
130       End If
140       frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
150    Unload Me
End Sub
 
Private Sub BORRAR_Click()
       
10    EstadoLogin = E_MODO.BorrarPJ
       
#If UsarWrench = 1 Then
20        If frmMain.Socket1.Connected Then
30            frmMain.Socket1.Disconnect
40            frmMain.Socket1.Cleanup
50            DoEvents
60        End If
70        frmMain.Socket1.HostName = CurServerIp
80        frmMain.Socket1.RemotePort = CurServerPort
90        frmMain.Socket1.Connect
#Else
100       If frmMain.Winsock1.State <> sckClosed Then
110           frmMain.Winsock1.Close
120           DoEvents
130       End If
140       frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
150    Unload Me
End Sub

