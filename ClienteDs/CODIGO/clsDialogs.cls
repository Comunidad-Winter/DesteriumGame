VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If Wgl = 0 Then
Option Explicit
Option Base 0

''
' Maximum number of dialogs that can exist.
Private Const MAX_DIALOGS As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Private Const MS_PER_CHAR As Byte = 100

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Private Const MS_ADD_EXTRA As Integer = 5000

''
' The dialog structure
'
' @param    textLine    Array of lines of a formated chat.
' @param    x           X coord for rendering.
' @param    y           Y coord for rendering.
' @param    startTime   The time (in ms) at which the dialog was created.
' @param    lifeTime    Time (in ms) this dialog should last.
' @param    charIndex   The charIndex that created this dialog.
' @param    color       The color to be used when rendering the text.
' @param    renderable  Set to True if the chat should be rendered this frame, False otherwise
'                           (used to skip dialogs from people outside render area).
Private Type DialoG
    textLine()  As String
    X           As Integer
    Y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    Color       As Long
    renderable  As Boolean
End Type

''
' Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1) As DialoG

''
' The number of dialogs being used at the moment.
Private dialogCount As Byte

''
' the size of the font being used to render
Private usedFont As StdFont

''
' A millisecond precision timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' Initializes the instance of the class.

Private Sub Class_Initialize()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/28/07
      'Initializes the dialog instance
      '***************************************************
10        dialogCount = 0
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Clean up
'***************************************************
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
      'Returns the index of the dialog in the list, or the negation
      'of the position were it should be if not found (for binary insertion)
      '**************************************************************
          Dim min As Long
          Dim max As Long
          Dim mid As Long
          
10        min = 0
20        max = dialogCount - 1
          
30        Do While min <= max
40            mid = (min + max) \ 2
              
50            If dialogs(mid).CharIndex < CharIndex Then
60                min = mid + 1
70            ElseIf dialogs(mid).CharIndex > CharIndex Then
80                max = mid - 1
90            Else
                  'We found it
100               BinarySearch = mid
110               Exit Function
120           End If
130       Loop
          
          'Not found, return the negation of the position where it should be
          '(all higher values are to the right of the list and lower values are to the left)
140       BinarySearch = Not min
End Function

''
' Creates a new dialog. If the user already had a dialog, then it's replaced,
' otherwise a new slot is used.
'
' @param    chat        The text of the chat being spoken.
' @param    charIndex   The character saying the chat.
' @param    color       The color being used in the chat when rendered.
'
' @remarks  If there are no more free slots, and the char didn't have a previous dialog,
'           the chat is ignored.

Public Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByVal _
    Color As Long)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modification: 08/29/07 by Rapsodius
      'Changed life time of the dialogs to a variable amount,
      'depending on the number of chars
      '**************************************************************
          Dim Slot As Integer
          Dim i As Long
          
10        Slot = BinarySearch(CharIndex)
          
20        If Slot < 0 Then
30            If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
              
              'We need to add it. Get insertion index and move list backwards.
40            Slot = Not Slot
              
50            For i = dialogCount To Slot + 1 Step -1
60                dialogs(i) = dialogs(i - 1)
70            Next i
              
80            dialogCount = dialogCount + 1
90        End If
          
100       With dialogs(Slot)
110           .CharIndex = CharIndex
120           .Color = Color
130           .renderable = False
140           .startTime = GetTickCount()
150           .textLine = FormatChat(chat)
160           .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(chat))
170       End With

          '****************************************
          'SAVE LAST DIALOG!!
180       Call FotoD_SaveLastDialog(CharIndex, chat)
          'SAVE LAST DIALOG!!
          '****************************************
End Sub

''
' Splits a text into several lines to make it comply with the MAX_LENGTH unless it's impossible (a single word longer than MAX_LENGTH).
'
' @param    chat    The text to be formated.
'
' @return   The array of lines into which the text is splitted.
'
' @see      MAX_LENGTH

Private Function FormatChat(ByRef chat As String) As String()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
      'Formats a dialog into different text lines.
      '**************************************************************
          Dim word As String
          Dim curPos As Long
          Dim Length As Long
          Dim acumLength As Long
          Dim lineLength As Long
          Dim wordLength As Long
          Dim curLine As Long
          Dim chatLines() As String
          
          'Initialize variables
10        curLine = 0
20        curPos = 1
30        Length = Len(chat)
40        acumLength = 0
50        lineLength = -1
60        ReDim chatLines(FieldCount(chat, 32)) As String
          
          'Start formating
70        Do While acumLength < Length
80            word = ReadField(curPos, chat, 32)
              
90            wordLength = Len(word)
              
              ' Is the first word of the first line? (it's the only that can start at -1)
100           If lineLength = -1 Then
110               chatLines(curLine) = word
                  
120               lineLength = wordLength
130               acumLength = wordLength
140           Else
                  ' Is the word too long to fit in this line?
150               If lineLength + wordLength + 1 > MAX_LENGTH Then
                      'Put it in the next line
160                   curLine = curLine + 1
170                   chatLines(curLine) = word
                      
180                   lineLength = wordLength
190               Else
                      'Add it to this line
200                   chatLines(curLine) = chatLines(curLine) & " " & word
                      
210                   lineLength = lineLength + wordLength + 1
220               End If
                  
230               acumLength = acumLength + wordLength + 1
240           End If
              
              'Increase to search for next word
250           curPos = curPos + 1
260       Loop
          
          ' If it's only one line, center text
270       If curLine = 0 And Length < MAX_LENGTH Then
280           chatLines(curLine) = String((MAX_LENGTH - Length) \ 2 + 1, " ") & _
                  chatLines(curLine)
290       End If
          
          'Resize array to fit
300       ReDim Preserve chatLines(curLine) As String
          
310       FormatChat = chatLines
End Function

''
' Updates the position where a given dialog will be rendered.
'
' @param    x           The X coord where to start rendering.
' @param    y           The Y coord where to start rendering.
' @param    charIndex   The character whose dialog's position is being updated.

Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal _
    CharIndex As Integer)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
      'Updates a dialog's position.
      '**************************************************************
          Dim Slot As Integer
          
10        Slot = BinarySearch(CharIndex)
          
20        If Slot >= 0 Then
30            With dialogs(Slot)
40                .X = X - (MAX_LENGTH \ 2) * 3
50                .Y = Y - (UBound(.textLine()) * 3)
60                .renderable = True
70            End With
80        End If
End Sub

''
' Displays all renderable dialogs on screen.

Public Sub Render()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modification: 08/29/07 by Rapsodius
      'Changed the way it gets rid of older dialogs.
      '**************************************************************
          Dim i As Long
          Dim j As Long
          Dim offset As Integer
          Dim removedDialogs As Long
          
10        For i = 0 To dialogCount - 1
              'Decrease index to prevent jumping over a dialog
              'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
20            With dialogs(i - removedDialogs)
30                If (GetTickCount() - .startTime) >= .lifeTime Then
40                    Call RemoveDialog(.CharIndex)
                      
50                    removedDialogs = removedDialogs + 1
60                ElseIf .renderable Then
70                    offset = -(usedFont.Size + 2) * UBound(.textLine())
                      
                      'Render it
80                    For j = 0 To UBound(.textLine())
90                        Call RenderText(.X, .Y + offset, .textLine(j), .Color, _
                              usedFont)
100                       offset = offset + usedFont.Size + 5
110                   Next j
                      
                      'Won't render it unless it's updated in next frmae
120                   .renderable = False
130               End If
140           End With
150       Next i
End Sub

''
' Removes the dialog from the requested char.
'
' @param    charIndex   The character whose dialog will be removed.

Public Sub RemoveDialog(ByVal CharIndex As Integer)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
      'Removes the dialog from the requested char.
      '**************************************************************
          Dim Slot As Integer
          Dim i As Long
          '****************************************
          'REMOVE LAST DIALOG!!
10        Call FotoD_RemoveLastDialog(CharIndex)
          'REMOVE LAST DIALOG!!
          '****************************************
20        Slot = BinarySearch(CharIndex)
          
30        If Slot < 0 Then Exit Sub
          
40        For i = Slot To MAX_DIALOGS - 2
50            dialogs(i) = dialogs(i + 1)
60        Next i
          
70        dialogCount = dialogCount - 1
End Sub

''
' Removes all dialogs from all characters.

Public Sub RemoveAllDialogs()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
      'Removes all dialogs from all chars.
      '**************************************************************
10        dialogCount = 0
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 07/28/07
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
      'Last Modify Date: 07/28/07
      'Sets the font being used to render dialogs.
      '**************************************************************
10        Set usedFont = nValue
End Property

#End If
