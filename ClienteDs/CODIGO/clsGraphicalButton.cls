VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical Button
'Designed and implemented by ZaMa
'zaxtor@gmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage As Picture  'img holder
Private NormalButtonImage As Picture 'img holder
Private ClickButtonImage As Picture  'img holder
Private DisabledButtonImage As Picture  'img holder

Private SoundClick As Boolean

Private FormActual As Form ' Form owner

Private OverButton As Boolean 'Indicates wether the mouse is over the button or not
Private Disabled As Boolean 'Indicates wether it's enabled or not

Public Sub Initialize(ByRef imgButton As Image, ByVal imgNormalPath As String, _
    ByVal imgRollOverPath As String, ByVal imgClickPath As String, ByRef Formulario _
    As Form, Optional ByVal imgDisabledPath As String, Optional ByVal StartDisabled _
    As Boolean = False, Optional ByVal MakeClickSound As Boolean = True, Optional _
    ByVal ShowHandPointer As Boolean = True)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 22/12/09
      'Sets the reference to the image container
      '***************************************************
10        Set ButtonHolder = imgButton
20        Set FormActual = Formulario
          
30        Set RollOverButtonImage = LoadPicture(imgRollOverPath)
40        Set NormalButtonImage = LoadPicture(imgNormalPath)
50        Set ClickButtonImage = LoadPicture(imgClickPath)
60        Set DisabledButtonImage = LoadPicture(imgDisabledPath)
          
70        SoundClick = MakeClickSound
          
80        With ButtonHolder
90            If ShowHandPointer Then
100               .MousePointer = vbCustom
110               .MouseIcon = picMouseIcon
120           End If
              
130           If StartDisabled Then
140               Disabled = True
150               .Picture = DisabledButtonImage
160           Else
170               .Picture = NormalButtonImage
180           End If
190       End With
          
End Sub

Private Sub ButtonHolder_Click()
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 02/01/2010
      'Implements the mouse click event of the image button.
      'Plays the Click Sound.
      '***************************************************
10        If Disabled Then Exit Sub
20        If Not SoundClick Then Exit Sub
          
30        Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 29/12/09
      'Implements the mouse down event of the image button.
      'Toggle to Click image button.
      '***************************************************
10        If Disabled Then Exit Sub
          
20        ButtonHolder.Picture = ClickButtonImage
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 22/12/09
      'Implements the mouse move event of the image button.
      'Toggle to pressed image button.
      '***************************************************
10        If Disabled Then Exit Sub
          
20        If Not OverButton Then
          
              'This helps when two buttons are close
30            FormActual.LastPressed.ToggleToNormal
              
40            ButtonHolder.Picture = RollOverButtonImage
50            OverButton = True
              
60            Set FormActual.LastPressed = Me
70        End If
End Sub

Public Sub ToggleToNormal()
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 22/12/09
      'Toggle to normal image button.
      '***************************************************
10        If Disabled Then Exit Sub
          
20        If Not OverButton Then Exit Sub
          
30        ButtonHolder.Picture = NormalButtonImage
40        OverButton = False
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: 29/12/09
      'Implements the mouse up event of the image button.
      'Toggle to RollOver image button.
      '***************************************************
10        If Disabled Then Exit Sub
          
20        ButtonHolder.Picture = RollOverButtonImage
End Sub

Public Sub EnableButton(Optional ByVal ToogleEnable As Boolean = True)
10        If Disabled = Not ToogleEnable Then Exit Sub
          
20        OverButton = False
30        Disabled = Not ToogleEnable
40        If Disabled Then
            '  ButtonHolder.Picture = DisabledButtonImage
50        Else
            '  ButtonHolder.Picture = NormalButtonImage
60        End If
          
End Sub

Public Function IsEnabled() As Boolean
10        IsEnabled = Not Disabled
End Function
