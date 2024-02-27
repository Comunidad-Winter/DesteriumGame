VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg
    Texto As String
    Segundos As Byte
End Type

Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

Private usedFont As StdFont

Public Property Let CantidadDialogos(ByVal v As Byte)
10        If v > 0 Then
20            ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
30            p_CANTIDADDIALOGOS = v
40        End If
End Property

''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/02/07
      'Removes all dialogs
      '***************************************************
          Dim i As Long
          
10        For i = 1 To p_CANTIDADDIALOGOS
20            p_Dialogos(i).Texto = vbNullString
30        Next i
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 04/01/07
      'Retrieves the number of dialogs to be displayed on screen
      '***************************************************
10        CantidadDialogos = p_CANTIDADDIALOGOS
End Property

Public Property Let Activo(ByVal v As Boolean)
10        p_Activo = v
End Property

Public Property Get Activo() As Boolean
10        Activo = p_Activo
End Property

Private Sub Class_Initialize()
10        p_CANTIDADDIALOGOS = 5
20        ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
30        p_Activo = True
End Sub

Public Sub Draw()
      Dim i As Byte
      Dim Y As Long

10    If Not p_Activo Then Exit Sub

20    Y = 605
      '500 600 700 800 900
30        For i = 1 To p_CANTIDADDIALOGOS
40            If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
                  #If Wgl = 0 Then
                        If Left$(p_Dialogos(i).Texto, 1) = Chr(3) Then
60                          Call Mod_TileEngine.RenderText(260, Y + (i * 10), _
                                mid(p_Dialogos(i).Texto, 2), vbGreen, usedFont)
70                      Else
80                          Call Mod_TileEngine.RenderText(260, Y + (i * 10), _
                                p_Dialogos(i).Texto, vbGreen, usedFont)
90                      End If
                  #End If
100           End If
110       Next i
End Sub

Public Sub PassTimer()
      Dim i As Byte

10    For i = 1 To p_CANTIDADDIALOGOS
20        If p_Dialogos(i).Segundos > 0 Then
30            p_Dialogos(i).Segundos = p_Dialogos(i).Segundos - 1
40        Else
50            p_Dialogos(i).Texto = vbNullString
60        End If
70    Next i
          
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/04/2009
      'Splits dialogs into lines fitting properly the render area and inserts them
      '07/04/2009: Now cuts the string properly in spite of not fitting in the screen. This avoids an infite loop.
      '***************************************************
          Dim str As String
          Dim tmp As Integer
          
10        str = s
          
          ' If it's too long to fit, split it
20        Do While frmMain.TextWidth(str) > 500
30            tmp = InStrRev(str, " ")
              
40            If tmp = 0 Then Exit Do
              
50            str = Left$(str, tmp - 1)
60        Loop
          
          'Show message and continue with the rest
70        Call RealPushBackText(str)
          
80        If LenB(str) <> LenB(s) Then
90            Call PushBackText(Right$(s, Len(s) - Len(str) - 1))
100       End If
End Sub

Private Sub RealPushBackText(ByVal s As String)
      Dim i As Byte
      Dim Vacio As Boolean

10    If p_CANTIDADDIALOGOS > 0 Then
20        i = p_CANTIDADDIALOGOS
30        Vacio = True
40        While i > 0 And Vacio
50            Vacio = p_Dialogos(i).Texto = vbNullString
60            If Vacio Then i = i - 1
70        Wend
80        If i = p_CANTIDADDIALOGOS Then
              'hay q scrollear, estamos llenos
90            i = 1
100           While i < p_CANTIDADDIALOGOS
110               p_Dialogos(i) = p_Dialogos(i + 1)
120               i = i + 1
130           Wend
140           p_Dialogos(i).Texto = s
150           p_Dialogos(i).Segundos = p_SEGUNDOS
160       Else
170           p_Dialogos(i + 1).Texto = s
180           p_Dialogos(i + 1).Segundos = p_SEGUNDOS
190       End If
200   End If

End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 04/27/06
      'Retrieves the font being used to render dialogs.
      '**************************************************************
10        font = usedFont
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 04/27/06
      'Sets the font being used to render dialogs.
      '**************************************************************
10        Set usedFont = nValue
End Property
