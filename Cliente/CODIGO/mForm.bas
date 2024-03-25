Attribute VB_Name = "mForm"

Option Explicit

'------------------------------------------------------------------------------
' APIS para incluir las ventanas en un PictureBox
'------------------------------------------------------------------------------
'
' Para hacer ventanas hijas
Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long

'
' Para mostrar una ventana seg�n el handle (hwnd)
' ShowWindow() Commands
Private Enum eShowWindow

    HIDE_eSW = 0&
    SHOWNORMAL_eSW = 1&
    NORMAL_eSW = 1&
    SHOWMINIMIZED_eSW = 2&
    SHOWMAXIMIZED_eSW = 3&
    MAXIMIZE_eSW = 3&
    SHOWNOACTIVATE_eSW = 4&
    SHOW_eSW = 5&
    MINIMIZE_eSW = 6&
    SHOWMINNOACTIVE_eSW = 7&
    SHOWNA_eSW = 8&
    RESTORE_eSW = 9&
    SHOWDEFAULT_eSW = 10&
    MAX_eSW = 10&

End Enum

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nCmdShow As eShowWindow) As Long

'
' Para posicionar una ventana seg�n su hWnd
Public Declare Function MoveWindow _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal bRepaint As Long) As Long

'
' Para cambiar el tama�o de una ventana y asignar los valores m�ximos y m�nimos del tama�o
Private Type POINTAPI

    X As Long
    Y As Long

End Type

Private Type RECTAPI

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Type WINDOWPLACEMENT

    Length As Long
    Flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECTAPI

End Type

Private Declare Function GetWindowPlacement _
                Lib "user32" (ByVal hWnd As Long, _
                              ByRef lpwndpl As WINDOWPLACEMENT) As Long

' Mostrar el formulario indicado, dentro de picDock
Public Sub dockForm(ByVal formhWnd As Long, _
                    ByVal picDock As PictureBox, _
                    Optional ByVal ajustar As Boolean = True)
    ' Hacer el formulario indicado, un hijo del picDock
    ' Si Ajustar es True, se ajustar� al tama�o del contenedor,
    ' si Ajustar es False, se quedar� con el tama�o actual.
    Call SetParent(formhWnd, picDock.hWnd)
    posDockForm formhWnd, picDock, ajustar
    Call ShowWindow(formhWnd, NORMAL_eSW)
    
    FrmMain.SetFocus

End Sub

' Posicionar el formulario indicado dentro de picDock
Private Sub posDockForm(ByVal formhWnd As Long, _
                        ByRef picDock As PictureBox, _
                        Optional ByVal ajustar As Boolean = True)

    ' Posicionar el formulario indicado en las coordenadas del picDock
    ' Si Ajustar es True, se ajustar� al tama�o del contenedor,
    ' si Ajustar es False, se quedar� con el tama�o actual.
    Dim nWidth As Long, nHeight As Long

    Dim wndPl  As WINDOWPLACEMENT

    '
    
    If ajustar Then
        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
    Else
        ' el tama�o del formulario que se va a posicionar
        Call GetWindowPlacement(formhWnd, wndPl)
        
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top

        End With

    End If

    Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)

End Sub
