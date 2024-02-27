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
Unload Me

'MsgBox "Atención con esta acción va a eliminar el personaje, no podra volver a usarlo"
End Sub

Private Sub RECUPERAR_Click()
 
EstadoLogin = E_MODO.RecuperarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = "190.112.0.123"
    frmMain.Socket1.RemotePort = 7818
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
 Unload Me
End Sub
 
Private Sub BORRAR_Click()
 
EstadoLogin = E_MODO.BorrarPJ
 
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    frmMain.Socket1.HostName = "190.112.0.123"
    frmMain.Socket1.RemotePort = 7818
    frmMain.Socket1.Connect
#Else
    If frmMain.Winsock1.State <> sckClosed Then
        frmMain.Winsock1.Close
        DoEvents
    End If
    frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If
 Unload Me
End Sub

