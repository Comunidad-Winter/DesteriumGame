VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadisticasIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsEstadisticasIPC.cls
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private sMensaje As Long, hVentana As Long, hVentanaMia As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2

'*************************************************
Public Enum EstaNotificaciones
    CANTIDAD_ONLINE = 1
    RECORD_USUARIOS = 2
    UPTIME_SERVER = 3
    CANTIDAD_MAPAS = 4
    EVENTO_NUEVO_CLAN = 5
    
    HANDLE_WND_SERVER = 100
End Enum

'*************************************************

Private Function BuscaVentana(Wnd As Long, Str As String) As Long
      Dim W As Long
      Dim L As Long
      Dim t As String


10    W = GetWindow(Wnd, GW_HWNDFIRST)

20    While W <> 0
30       L = GetWindowTextLength(W)

40       If L > 0 Then
50          t = Space(L + 1)
60          L = GetWindowText(W, t, L + 1)

70          If Left$(t, Len(Str)) = Str Then
80             BuscaVentana = W
90             Exit Function
100         End If
110      End If

120      W = GetWindow(W, GW_HWNDNEXT)
130   Wend

140   BuscaVentana = 0

End Function

Public Function Informar(ByVal QueCosa As EstaNotificaciones, ByVal Parametro As Long) As Long
10    Call BuscaWndEstadisticas
20    If hVentana <> 0 Then
30        Informar = SendMessageLong(hVentana, sMensaje, QueCosa, Parametro)
40    End If

End Function

Public Function EstadisticasAndando() As Boolean

10    Call BuscaWndEstadisticas
      'Ret = SendNotifyMessage(hVentana, sMensaje, 0, 0)
20    EstadisticasAndando = (hVentana <> 0)

End Function

Public Sub Inicializa(ByVal hWnd As Long)
10    hVentanaMia = hWnd
20    sMensaje = RegisterWindowMessage("EstadisticasAO")

End Sub

Private Sub BuscaWndEstadisticas()
10    hVentana = BuscaVentana(hVentanaMia, "Servidor de estadisticas AO")

End Sub
