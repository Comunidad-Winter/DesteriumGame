Attribute VB_Name = "mConsole"
Option Explicit
 
Const CONSOLE_LINES As Integer = 5
 
Private Type consoleLine

    Alpha As Byte
    DamageType As EDType
    
    mString     As String
    Time        As Long
    Duration    As Long
    Shadow As Boolean

End Type
 
Private RenderConsole(CONSOLE_LINES - 1) As consoleLine
 
Public Declare Function AddFontResource _
               Lib "gdi32.dll" _
               Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Public Declare Function RemoveFontResource _
               Lib "gdi32.dll" _
               Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
 
Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Declare Function EnumWindows _
        Lib "user32" (ByVal wndenmprc As Long, _
                      ByVal lParam As Long) As Long
 
Declare Function GetWindowText _
        Lib "user32" _
        Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                ByVal lpString As String, _
                                ByVal cch As Long) As Long
 
Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, _
                              ByVal wMsg As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long
  
Private Declare Function GetWindowTextLength _
                Lib "user32" _
                Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function GetWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wCmd As Long) As Long
 
Const WM_SYSCOMMAND = &H112

Const SC_CLOSE = &HF060&

Private Const GW_HWNDNEXT = 2

Private Caption       As String

Public CaptionTemp    As String

Public TempModuleName As String

' These subs temporarily install and uninstall a font
' for use while your program is running. You MUST
' call this code to install the font before loading
' any form that uses it, or you'll generate an error.
Public Function AddFont(ByVal sFontFile As String) As Boolean
    ' e.g.: AddFontResource("c:\myApp\myFont.ttf")
    AddFont = (AddFontResource(sFontFile) <> 0)

End Function

' To remove the font
Public Sub RemoveFont(ByVal sFontFile As String)

    Dim lResult As Long

    'e.g.: RemoveFontResource "c:\myApp\myFont.ttf"
    lResult = RemoveFontResource(sFontFile)

End Sub

Public Sub RenderText_Console_Add(ByRef mText As String, _
                                  ByVal DamageType As EDType, _
                                  ByVal Duration As Long, _
                                  ByVal Slot As Byte)
    
    Dim SlotSearched           As Byte

    Dim LoopC                  As Long

    Dim tmp(CONSOLE_LINES - 1) As consoleLine
        
    If Slot = 0 Then
    
        For LoopC = 0 To (CONSOLE_LINES - 1)
            tmp(LoopC) = RenderConsole(LoopC)
        Next LoopC
        
        For LoopC = 1 To (CONSOLE_LINES - 1)
            RenderConsole(LoopC - 1) = tmp(LoopC)
        Next LoopC
        
        SlotSearched = (CONSOLE_LINES - 1) ' Último Slot
    Else
        SlotSearched = Slot

    End If
    
    With RenderConsole(SlotSearched)
        .Alpha = 255
        .mString = mText
        .Time = FrameTime
        .Duration = Duration
        .DamageType = DamageType
        .Shadow = True

    End With
 
End Sub
 
Public Sub RenderText_Console()
    
    Dim RenderY As Integer

    Dim LoopC   As Long

    Dim LoopX   As Long

    Dim AddY    As Long
    
    #If ModoBig = 0 Then
        AddY = 11
    #Else
        AddY = 22
    #End If
    
    For LoopC = 0 To (CONSOLE_LINES - 1)

        With RenderConsole((CONSOLE_LINES - 1) - LoopC)

            If .mString <> vbNullString Then
                
                RenderY = ((CONSOLE_LINES - 1) - LoopC + 1) * AddY
                     
                Call Draw_Text(f_Tahoma, 15, 12, 12 + RenderY, To_Depth(8), 0, ModifyColour(.Alpha, .DamageType), FONT_ALIGNMENT_BASELINE, .mString, .Shadow)
                
                If (FrameTime - .Time) >= .Duration Then
                    .Shadow = False
                    .Alpha = .Alpha - 1
                        
                    If .Alpha = 0 Then
                        .mString = vbNullString
                        .Alpha = 0
                        .Time = 0
                        .Duration = 0
                        .DamageType = 0

                    End If

                End If

            End If

        End With

    Next LoopC
    
End Sub

Public Sub RenderText_Clear()

    Dim LoopC As Integer
    
    For LoopC = 0 To (CONSOLE_LINES - 1)

        With RenderConsole(LoopC)
            .mString = vbNullString
            .Alpha = 0
            .Time = 0
            .Duration = 0
            .DamageType = 0

        End With

    Next LoopC

End Sub

Private Function GetHandleFromPartialCaption(ByRef lWnd As Long, _
                                             ByVal sCaption As String) As String

    '<EhHeader>
    On Error GoTo GetHandleFromPartialCaption_Err

    '</EhHeader>
    Dim lhWndP As Long

    Dim sStr   As String

    lhWndP = FindWindow(vbNullString, vbNullString)  'PARENT WINDOW

    Do While lhWndP <> 0
        sStr = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
        GetWindowText lhWndP, sStr, Len(sStr)
        sStr = Left$(sStr, Len(sStr) - 1)

        If InStr(1, UCase$(sStr), UCase$(sCaption)) > 0 Then
            GetHandleFromPartialCaption = sStr
            lWnd = lhWndP
            Exit Do

        End If

        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
    '<EhFooter>
    Exit Function

GetHandleFromPartialCaption_Err:
    'LogError err.Description & vbCrLf & _
     "in ARGENTUM.mConsole.GetHandleFromPartialCaption " & _
     "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Private Function TextFijo(ByVal Text As String) As Boolean

    Select Case Text
    
        Case "INJECTED ANTI-CHEAT"
            TextFijo = True

        Case "MACROKEY HIDDEN WND"
            TextFijo = True

        Case "BakkesModInjectorCpp"
            TextFijo = True

    End Select

End Function

Public Function SearchDesterium()

    '<EhHeader>
    On Error GoTo SearchDesterium_Err

    '</EhHeader>
    Dim lhWndP       As Long

    Dim Searching(6) As String

    Dim Temp         As String
    
    Searching(1) = "MACRO"
    Searching(2) = "CHEAT"
    Searching(3) = "XENOS"
    Searching(4) = "INJECTOR"
    Searching(5) = "INYECTOR"
    Searching(6) = "SÍMBOLO"
          
    Dim A As Long
    
    For A = 1 To 6
        Temp = GetHandleFromPartialCaption(lhWndP, Searching(A))
         
        If Temp <> vbNullString And Not TextFijo(UCase$(Temp)) Then
            CaptionTemp = CaptionTemp & Temp & ", "

        End If
         
    Next A
    
    If Len(CaptionTemp) <> 0 Then
        CaptionTemp = Left$(CaptionTemp, Len(CaptionTemp) - 2)

    End If

    '<EhFooter>
    Exit Function

SearchDesterium_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.mConsole.SearchDesterium " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

