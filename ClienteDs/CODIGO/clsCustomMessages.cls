VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomMessages.cls - User-defined messages with shortcut.
' Allows the creation and serialization of user defined messages to be used with shortcuts.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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
Option Base 0

''
' Defines the number of messages that can be created.
Private Const MESSAGE_COUNT As Byte = 10

''
' Defines the maximum length allowed for any custom message.
Private Const MAX_MESSAGE_LENGTH As Byte = 120

''
' The error code thrown when the message being set is invalid.
Private Const INVALID_MESSAGE As Long = vbObjectError + 123

''
' Relative path from the application in which custom message are stored.
Private Const MESSAGE_FILE_PATH As String = "/init/messages.txt"


''
'Array of defined messages.
Private CustomMessages(MESSAGE_COUNT - 1) As String


''
' Class constructor. Initializes a new clsCustomMessage

Private Sub Class_Initialize()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/26/07
      'Initializes the messages by loading them from the config file
      '***************************************************
10    On Error GoTo ErrHandler
          Dim handle As Integer
          Dim i As Long
          
          'Load custom messages
20        handle = FreeFile()
          
30        Open App.path & MESSAGE_FILE_PATH For Input As handle
          
40        Do While (i < MESSAGE_COUNT) And (Not EOF(handle))
50            Line Input #handle, CustomMessages(i)
              
              'Purge invalid messages
60            If Not MessageIsValid(CustomMessages(i)) Then
70                CustomMessages(i) = vbNullString
80            End If
              
90            i = i + 1
100       Loop
          
110       Close handle
          
          'If the file does not exist (only possible error),
          'then we have nothing to worry. All messages will be empty
ErrHandler:
End Sub

''
' Class desconstructor. Serializes messages to

Private Sub Class_Terminate()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/26/07
      'Stores all messages back into config file for later use.
      '***************************************************
          Dim handle As Integer
          Dim i As Long
          Dim directories() As String
          Dim path As String
          
          'Make sure the path exists
10        directories = Split(MESSAGE_FILE_PATH, "/")
20        path = App.path
          
30        For i = 0 To UBound(directories()) - 1
40            path = path & "/" & directories(i)
              
50            If LenB(Dir(path, vbDirectory)) = 0 Then
60                Call MkDir(path)
70            End If
80        Next i
          
          'Serialize custom messages
90        handle = FreeFile()
          
100       Open App.path & MESSAGE_FILE_PATH For Output As handle
          
110       For i = 0 To MESSAGE_COUNT - 1
120           Print #handle, CustomMessages(i)
130       Next i
          
140       Close handle
End Sub

''
' Retrieves the requested message.
'
' @param    index   The message index to retrieve.
' @return   The message requested.

Public Property Get Message(ByVal Index As Integer) As String
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/26/07
      'Retrieves the requested message.
      '***************************************************
10        If Index < 0 Or Index >= MESSAGE_COUNT Then
20            Exit Property
30        End If
          
40        Message = CustomMessages(Index)
End Property

''
' Sets the requested message.
'
' @param    index   The message index to be set.
' @param    nValue  The message to be stored.

Public Property Let Message(ByVal Index As Integer, ByVal nValue As String)
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/26/07
      'Sets the given message index with the requested message.
      '***************************************************
10        If Index < 0 Or Index >= MESSAGE_COUNT Then
20            Exit Property
30        End If
          
40        If Not MessageIsValid(nValue) Then
50            Call Err.Raise(INVALID_MESSAGE)
60            Exit Property
70        End If
          
80        CustomMessages(Index) = nValue
End Property

''
' Retrieves the error code of invalid messages.
'
' @return   The error code triggered when an invalid message is found.

Public Property Get InvalidMessageErrCode() As Long
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 07/26/07
      'Retrieves the INVALID_MESSAGE error code.
      '***************************************************
10        InvalidMessageErrCode = INVALID_MESSAGE
End Property

''
' Checks if the given message is valid or not.
'
' @param    msg   The message to be checked.
' @return   True if the message is valid, False otherwise.

Private Function MessageIsValid(ByRef msg As String) As Boolean
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 08/07/07
      'Last Modified by: Rapsodius
      '
      'Erased line "msg=StrConv(msg,vbUnicode)", because
      'it was causing problems and detected as many messages
      'as wrong with nothing bad on them
      '***************************************************
          Dim i As Long
          Dim Char As String * 1
          Dim ascii As Byte
          
          'Get rid of spaces
10        msg = Trim$(msg)
              
          'Empty messages are valid
20        If LenB(msg) = 0 Then
30            MessageIsValid = True
40            Exit Function
50        End If
          
60        For i = 1 To Len(msg)
70            Char = mid$(msg, i, 1)
80            ascii = Asc(Char)
              
              'check for invalid chars...
90            If ascii < 31 Then Exit Function
              
100           If ascii > 126 And Char <> "á" And Char <> "Á" And Char <> "ä" And Char _
                  <> "Ä" And Char <> "â" And Char <> "Â" And Char <> "à" And Char <> "À" _
                  And Char <> "é" And Char <> "É" And Char <> "ë" And Char <> "Ë" And _
                  Char <> "ê" And Char <> "Ê" And Char <> "è" And Char <> "È" And Char <> _
                  "í" And Char <> "Í" And Char <> "ï" And Char <> "Ï" And Char <> "î" And _
                  Char <> "Î" And Char <> "ì" And Char <> "Ì" And Char <> "ó" And Char <> _
                  "Ó" And Char <> "ö" And Char <> "Ö" And Char <> "ô" And Char <> "Ô" And _
                  Char <> "ò" And Char <> "Ò" And Char <> "ú" And Char <> "Ú" And Char <> _
                  "ü" And Char <> "Ü" And Char <> "û" And Char <> "Û" And Char <> "ù" And _
                  Char <> "Ù" And Char <> "ñ" And Char <> "Ñ" And Char <> "¡" And Char <> _
                  "¿" Then
110               Exit Function
120           End If
130       Next i
          
          'Discard extremely long messages
140       If Len(msg) > MAX_MESSAGE_LENGTH Then Exit Function
          
150       MessageIsValid = True
End Function




