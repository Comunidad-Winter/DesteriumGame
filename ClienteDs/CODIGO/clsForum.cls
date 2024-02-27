VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Forum
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

Private Const MAX_MENSAJES_FORO As Byte = 30
Private Const MAX_ANUNCIOS_FORO As Byte = 5

' Post counter
Private CantPosts(0 To 2) As Byte
Private CantStickys(0 To 2) As Byte

Public Sub AddPost(ByVal ForumType As eForumType, ByRef Title As String, ByRef _
    Author As String, ByRef MsgPost As String, ByVal Sticky As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: -
      '
      '***************************************************

10        With Foros(ForumType)
20            If Sticky Then
30                If CantStickys(ForumType) < MAX_ANUNCIOS_FORO Then _
                      CantStickys(ForumType) = CantStickys(ForumType) + 1
              
40                Call MoveArray(ForumType, Sticky)
                  
50                .StickyTitle(1) = Title
60                .StickyAuthor(1) = Author
70                .StickyPost(1) = MsgPost
80            Else
90                If CantPosts(ForumType) < MAX_MENSAJES_FORO Then _
                      CantPosts(ForumType) = CantPosts(ForumType) + 1
                  
100               Call MoveArray(ForumType, Sticky)
                  
110               .GeneralTitle(1) = Title
120               .GeneralAuthor(1) = Author
130               .GeneralPost(1) = MsgPost
140           End If
150       End With

End Sub

Public Sub ClearForums()
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: -
      '
      '***************************************************

          Dim Index As Long
          
10        If CantPosts(eForumType.ieGeneral) > 0 Then
20            For Index = 1 To CantPosts(eForumType.ieGeneral)
30                Foros(eForumType.ieGeneral).GeneralAuthor(Index) = vbNullString
40                Foros(eForumType.ieGeneral).GeneralTitle(Index) = vbNullString
50                Foros(eForumType.ieGeneral).GeneralPost(Index) = vbNullString
60            Next Index
              
70            CantPosts(eForumType.ieGeneral) = 0
80        End If
          
90        If CantPosts(eForumType.ieREAL) > 0 Then
100           For Index = 1 To CantPosts(eForumType.ieREAL)
110               Foros(eForumType.ieREAL).GeneralAuthor(Index) = vbNullString
120               Foros(eForumType.ieREAL).GeneralTitle(Index) = vbNullString
130               Foros(eForumType.ieREAL).GeneralPost(Index) = vbNullString
140           Next Index
              
150           CantPosts(eForumType.ieREAL) = 0
160       End If
          
170       If CantPosts(eForumType.ieCAOS) > 0 Then
180           For Index = 1 To CantPosts(eForumType.ieCAOS)
190               Foros(eForumType.ieCAOS).GeneralAuthor(Index) = vbNullString
200               Foros(eForumType.ieCAOS).GeneralTitle(Index) = vbNullString
210               Foros(eForumType.ieCAOS).GeneralPost(Index) = vbNullString
220           Next Index
              
230           CantPosts(eForumType.ieCAOS) = 0
240       End If
          
250       If CantStickys(eForumType.ieGeneral) > 0 Then
260           For Index = 1 To CantStickys(eForumType.ieGeneral)
270               Foros(eForumType.ieGeneral).StickyAuthor(Index) = vbNullString
280               Foros(eForumType.ieGeneral).StickyTitle(Index) = vbNullString
290               Foros(eForumType.ieGeneral).StickyPost(Index) = vbNullString
300           Next Index
              
310           CantStickys(eForumType.ieGeneral) = 0
320       End If
          
330       If CantStickys(eForumType.ieREAL) > 0 Then
340           For Index = 1 To CantStickys(eForumType.ieREAL)
350               Foros(eForumType.ieREAL).StickyAuthor(Index) = vbNullString
360               Foros(eForumType.ieREAL).StickyTitle(Index) = vbNullString
370               Foros(eForumType.ieREAL).StickyPost(Index) = vbNullString
380           Next Index
              
390           CantStickys(eForumType.ieREAL) = 0
400       End If
          
410       If CantStickys(eForumType.ieCAOS) > 0 Then
420           For Index = 1 To CantStickys(eForumType.ieCAOS)
430               Foros(eForumType.ieCAOS).StickyAuthor(Index) = vbNullString
440               Foros(eForumType.ieCAOS).StickyTitle(Index) = vbNullString
450               Foros(eForumType.ieCAOS).StickyPost(Index) = vbNullString
460           Next Index
              
470           CantStickys(eForumType.ieCAOS) = 0
480       End If

End Sub

Public Function GetNroPost(ByVal ForumType As eForumType) As Byte
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: -
      '
      '***************************************************

10        GetNroPost = CantPosts(ForumType)
End Function

Public Function GetNroSticky(ByVal ForumType As eForumType) As Byte
      '***************************************************
      'Author: ZaMa
      'Last Modify Date: -
      '
      '***************************************************

10        GetNroSticky = CantStickys(ForumType)
End Function

Private Sub MoveArray(ByVal ForumType As eForumType, ByVal Sticky As Boolean)
      Dim i As Long

10    With Foros(ForumType)
20        If Sticky Then
30            For i = CantStickys(ForumType) To 2 Step -1
40                .StickyTitle(i) = .StickyTitle(i - 1)
50                .StickyPost(i) = .StickyPost(i - 1)
60                .StickyAuthor(i) = .StickyAuthor(i - 1)
70            Next i
80        Else
90            For i = CantPosts(ForumType) To 2 Step -1
100               .GeneralTitle(i) = .GeneralTitle(i - 1)
110               .GeneralPost(i) = .GeneralPost(i - 1)
120               .GeneralAuthor(i) = .GeneralAuthor(i - 1)
130           Next i
140       End If
150   End With
End Sub
