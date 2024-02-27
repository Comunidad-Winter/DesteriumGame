VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08/03/07
' Created and Implemented by Rapsodius
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
'Customizes User Keys.
'
'@author Rapsodius
'@version 1.1.0
'@date 20080413

'03/08/2007 - Rapsodius
'   - First Release
'
'13/04/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Removed lots of unwanted buggy features keeping it clean. Revised documentation and coding standards.


Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 29

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\init\BindKeys.bin"

''
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyToggleSound Key for toggling sound
'@param mKeyToggleFxs Key for togglig the 3D Fxs
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyToggleCombatMode Key for toggling combat mode
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyToggleSound
    mKeyToggleFxs
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
    mKeyToggleCombatMode
    mKeyRETOS
End Enum

''
'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid

Public Function ReadableName(ByVal KeyCode As Integer) As String
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
10        Select Case KeyCode
              Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
20                ReadableName = Chr$(KeyCode)
              
30            Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
40                ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & _
                      " (teclado numérico)"
              
50            Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
60                ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
              
70            Case KeyCodeConstants.vbKeyMultiply
80                ReadableName = "* (teclado numerico)"
              
90            Case KeyCodeConstants.vbKeyAdd
100               ReadableName = "+ (teclado numerico)"
              
110           Case KeyCodeConstants.vbKeySeparator
                  'Note: Separator appears in MSDN as the Enter key for
                  'the numpad, but MS recognizes it as the normal
                  'Enter key
120               ReadableName = "Enter (teclado numérico)"
              
130           Case KeyCodeConstants.vbKeySubtract
140               ReadableName = "- (teclado numerico)"
              
150           Case KeyCodeConstants.vbKeyDecimal
160               ReadableName = ". (teclado numerico)"
              
170           Case KeyCodeConstants.vbKeyDivide
180               ReadableName = "/ (teclado numerico)"
              
190           Case KeyCodeConstants.vbKeyShift
200               ReadableName = "Shift"
              
210           Case KeyCodeConstants.vbKeyControl
220               ReadableName = "Control"
              
230           Case KeyCodeConstants.vbKeyMenu 'Alt
240               ReadableName = "Alt"
              
250           Case KeyCodeConstants.vbKeyPause
260               ReadableName = "Pausa"
              
270           Case KeyCodeConstants.vbKeyCapital
280               ReadableName = "Blq Mayus"
              
290           Case KeyCodeConstants.vbKeyEscape
300               ReadableName = "Esc"
              
310           Case KeyCodeConstants.vbKeyPageUp
320               ReadableName = "Ret Pag"
              
330           Case KeyCodeConstants.vbKeyPageDown
340               ReadableName = "Av Pag"
              
350           Case KeyCodeConstants.vbKeyEnd
360               ReadableName = "Fin"
              
370           Case KeyCodeConstants.vbKeyHome
380               ReadableName = "Inicio"
              
390           Case KeyCodeConstants.vbKeyLeft
400               ReadableName = "Izquierda"
              
410           Case KeyCodeConstants.vbKeyUp
420               ReadableName = "Arriba"
              
430           Case KeyCodeConstants.vbKeyRight
440               ReadableName = "Derecha"
              
450           Case KeyCodeConstants.vbKeyDown
460               ReadableName = "Abajo"
              
470           Case KeyCodeConstants.vbKeyInsert
480               ReadableName = "Insertar"
              
490           Case KeyCodeConstants.vbKeyNumlock
500               ReadableName = "Blq Num"
              
510           Case KeyCodeConstants.vbKeyScrollLock
520               ReadableName = "Blq Despl"
              
530           Case KeyCodeConstants.vbKeySpace
540               ReadableName = "Barra Espaciadora"
              
550           Case KeyCodeConstants.vbKeyBack
560               ReadableName = "Borrar"
              
570           Case KeyCodeConstants.vbKeyReturn
580               ReadableName = "Enter"
              
590           Case KeyCodeConstants.vbKeyDelete
600               ReadableName = "Suprimir"
              
610           Case Else
                  'In all other cases, key is considered not valid.
620               ReadableName = Constants.vbNullString
630       End Select
End Function

''
'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.

Public Sub LoadCustomKeys()
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/21/07
      '
      '***************************************************
          Dim i As Long
          Dim FileNum As Integer
10        FileNum = FreeFile()
          
          'Inits keyboard config with default configurations
20        Call LoadDefaults
              
30        Open App.path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As _
              FileNum
          
40        For i = 1 To NUM_CUSTOM_KEYS
50            Get FileNum, , mappedKeys(i)
60            If LenB(ReadableName(mappedKeys(i))) = 0 Then
70                Close FileNum
80                Call Err.Raise(ERROR_NO_KEY_FILE)
90                Exit Sub
100           End If
110       Next i
          
120       Close FileNum
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
          Dim FileNum As Integer
          Dim i As Long
10        FileNum = FreeFile()
          
20        Open App.path & KEY_CONFIG_FILE For Binary Access Write As FileNum
          
30        For i = 1 To NUM_CUSTOM_KEYS
40            Put FileNum, , mappedKeys(i)
50        Next i
          
60        Close #FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
10        If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
20        BindedKey = mappedKeys(Index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/05/07
      '
      '***************************************************
10        If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
          
20        If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
30            Exit Property                      'Rejects the new assignment
40        End If
          
50        If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
60            Exit Property
70        End If
          
80        mappedKeys(Index) = NewVal
End Property

''
'Loads default keys

Public Sub LoadDefaults()
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
10        mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
20        mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
30        mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
40        mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
50        mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
60        mappedKeys(eKeyType.mKeyToggleSound) = KeyCodeConstants.vbKeyS
70        mappedKeys(eKeyType.mKeyToggleFxs) = KeyCodeConstants.vbKeyF
80        mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
90        mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
100       mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
110       mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
120       mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
130       mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
140       mappedKeys(eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyMultiply
150       mappedKeys(eKeyType.mKeyToggleResuscitationSafe) = KeyCodeConstants.vbKeyEnd
160       mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
170       mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
180       mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
190       mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
200       mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
210       mappedKeys(eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
220       mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF12
230       mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
240       mappedKeys(eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
250       mappedKeys(eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
260       mappedKeys(eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
270       mappedKeys(eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyF4
280       mappedKeys(eKeyType.mKeyToggleCombatMode) = KeyCodeConstants.vbKeyC
290       mappedKeys(eKeyType.mKeyRETOS) = KeyCodeConstants.vbKeyF5
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
          Dim Counter As Long
          
10        KeyAssigned = False
          
20        For Counter = 1 To NUM_CUSTOM_KEYS
30            If mappedKeys(Counter) = KeyCode Then
40                KeyAssigned = True
50                Exit For
60            End If
70        Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get Count() As Byte
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
10        Count = NUM_CUSTOM_KEYS
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
On Error GoTo ErrHand:
10        Call LoadCustomKeys
20        Exit Sub
          
ErrHand:
30        If Err.number = ERROR_NO_KEY_FILE Then
40            Call LoadDefaults
50            Call SaveCustomKeys
60            Resume Next
70        End If
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
      '***************************************************
      'Author: Rapsodius
      'Last Modification: 08/04/07
      '
      '***************************************************
10        Call SaveCustomKeys
End Sub
