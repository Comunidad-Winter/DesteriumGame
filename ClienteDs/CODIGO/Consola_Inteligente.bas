Attribute VB_Name = "Consola_Inteligente"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal _
    wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As _
    Long
    'Geodar
Const EM_SETEVENTMASK = &H445
Const EN_LINK = &H70B
Const ENM_LINK = &H4000000
Const EM_AUTOURLDETECT = &H45B
Const EM_GETEVENTMASK = &H43B
Const GWL_WNDPROC = (-4) 'Geodar
Const WM_NOTIFY = &H4E 'Geodar
Const WM_LBUTTONDOWN = &H201
Const EM_GETTEXTRANGE = &H44B
 
Dim lOldProc As Long
Dim hWndRTB As Long
Dim hWndParent As Long
'Geodar
Private Type NMHDR 'Geodar
    hWndFrom As Long
    idFrom As Long
    code As Long
End Type 'Geodar
 
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'Geodar
Private Type ENLINK
    hdr As NMHDR
    msg As Long
    wParam As Long 'Geodar
    lParam As Long
    chrg As CHARRANGE 'Geodar
End Type
'Geodar
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
 
 
Public Sub Detectar(ByVal hWndTextbox As Long, ByVal hWndOwner As Long)
        'Don't want to subclass twice!
10      If lOldProc = 0 Then
          'Subclass!
20        lOldProc = SetWindowLong(hWndOwner, GWL_WNDPROC, AddressOf WndProc)
30        SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or _
              SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
40        SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
50        hWndParent = hWndOwner 'Geodar
60        hWndRTB = hWndTextbox
70      End If
End Sub
 
Public Sub NoDetectar()
10      If lOldProc Then
20        SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
          'Reset the window procedure (stop the subclassing)
30        SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
          'Set this to 0 so we can subclass again in future
40        lOldProc = 0
50      End If
End Sub
 
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As _
    Long, ByVal lParam As Long) As Long
          Dim uHead As NMHDR 'Geodar
          Dim eLink As ENLINK
          Dim eText As TEXTRANGE
          Dim sText As String
          Dim lLen As Long 'Geodar
         
          'Which message?
10        Select Case uMsg 'Geodar
          Case WM_NOTIFY
              'Copy the notification header into our structure from the pointer
20            CopyMemory uHead, ByVal lParam, Len(uHead)
             
30            If (uHead.hWndFrom = hWndRTB) And (uHead.code = EN_LINK) Then
40                CopyMemory eLink, ByVal lParam, Len(eLink)
                 
                  'What kind of message?
50                Select Case eLink.msg 'Geodar
                 
                  Case WM_LBUTTONDOWN
60                  eText.chrg.cpMin = eLink.chrg.cpMin
70                  eText.chrg.cpMax = eLink.chrg.cpMax 'Geodar
80                  eText.lpstrText = Space$(1024)
90                  lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
100                 sText = Left$(eText.lpstrText, lLen)
110                 ShellExecute hWndParent, vbNullString, sText, vbNullString, _
                        vbNullString, SW_SHOW
                     
120               End Select
                 
130           End If 'Geodar
             
140       End Select
150       WndProc = CallWindowProc(lOldProc, hWnd, uMsg, wParam, lParam) 'Geodar
End Function
 
