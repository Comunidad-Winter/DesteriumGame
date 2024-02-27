Attribute VB_Name = "PrevInstance"
'**************************************************************
' PrevInstance.bas - Checks for previous instances of the client running
' by using a named mutex.
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

''
'Prevents multiple instances of the game running on the same computer.
'
' @author Fredy Horacio Treboux (liquid) @and Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070104

Option Explicit

'Declaration of the Win32 API function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef _
    lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal _
    lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As _
    Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As _
    Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long

''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/04/07
      'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
      '***************************************************
          Dim sa As SECURITY_ATTRIBUTES
          
10        With sa
20            .bInheritHandle = 0
30            .lpSecurityDescriptor = 0
40            .nLength = LenB(sa)
50        End With
          
60        mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
          
70        CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function

''
' Checks if there's another instance of the app running, returns True if there is or False otherwise.

Public Function FindPreviousInstance() As Boolean
      '***************************************************
      'Author: Fredy Horacio Treboux (liquid)
      'Last Modification: 01/04/07
      '
      '***************************************************
          'We try to create a mutex, the name could be anything, but must contain no backslashes.
10        If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
              'There's no other instance running
20            FindPreviousInstance = False
30        Else
              'There's another instance running
40            FindPreviousInstance = True
50        End If
End Function

''
' Closes the client, allowing other instances to be open.

Public Sub ReleaseInstance()
      '***************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modification: 01/04/07
      '
      '***************************************************
10        Call ReleaseMutex(mutexHID)
20        Call CloseHandle(mutexHID)
End Sub
