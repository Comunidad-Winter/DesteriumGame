Attribute VB_Name = "modForum"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public Const MAX_MENSAJES_FORO As Byte = 30
Public Const MAX_ANUNCIOS_FORO As Byte = 5

Public Const FORO_REAL_ID As String = "REAL"
Public Const FORO_CAOS_ID As String = "CAOS"

Public Type tPost
    sTitulo As String
    sPost As String
    Autor As String
End Type

Public Type tForo
    vsPost(1 To MAX_MENSAJES_FORO) As tPost
    vsAnuncio(1 To MAX_ANUNCIOS_FORO) As tPost
    CantPosts As Byte
    CantAnuncios As Byte
    Id As String
End Type

Private NumForos As Integer
Private Foros() As tForo


Public Sub AddForum(ByVal sForoID As String)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Adds a forum to the list and fills it.
      '***************************************************
          Dim ForumPath As String
          Dim PostPath As String
          Dim PostIndex As Integer
          Dim FileIndex As Integer
          
10        NumForos = NumForos + 1
20        ReDim Preserve Foros(1 To NumForos) As tForo
          
30        ForumPath = App.Path & "\foros\" & sForoID & ".for"
          
40        With Foros(NumForos)
          
50            .Id = sForoID
              
60            If FileExist(ForumPath, vbNormal) Then
70                .CantPosts = val(GetVar(ForumPath, "INFO", "CantMSG"))
80                .CantAnuncios = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
                  
                  ' Cargo posts
90                For PostIndex = 1 To .CantPosts
100                   FileIndex = FreeFile
110                   PostPath = App.Path & "\foros\" & sForoID & PostIndex & ".for"

120                   Open PostPath For Input Shared As #FileIndex
                      
                      ' Titulo
130                   Input #FileIndex, .vsPost(PostIndex).sTitulo
                      ' Autor
140                   Input #FileIndex, .vsPost(PostIndex).Autor
                      ' Mensaje
150                   Input #FileIndex, .vsPost(PostIndex).sPost
                      
160                   Close #FileIndex
170               Next PostIndex
                  
                  ' Cargo anuncios
180               For PostIndex = 1 To .CantAnuncios
190                   FileIndex = FreeFile
200                   PostPath = App.Path & "\foros\" & sForoID & PostIndex & "a.for"

210                   Open PostPath For Input Shared As #FileIndex
                      
                      ' Titulo
220                   Input #FileIndex, .vsAnuncio(PostIndex).sTitulo
                      ' Autor
230                   Input #FileIndex, .vsAnuncio(PostIndex).Autor
                      ' Mensaje
240                   Input #FileIndex, .vsAnuncio(PostIndex).sPost
                      
250                   Close #FileIndex
260               Next PostIndex
270           End If
              
280       End With
          
End Sub

Public Function GetForumIndex(ByRef sForoID As String) As Integer
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Returns the forum index.
      '***************************************************
          
          Dim ForumIndex As Integer
          
10        For ForumIndex = 1 To NumForos
20            If Foros(ForumIndex).Id = sForoID Then
30                GetForumIndex = ForumIndex
40                Exit Function
50            End If
60        Next ForumIndex
          
End Function

Public Sub AddPost(ByVal ForumIndex As Integer, ByRef Post As String, ByRef Autor As String, _
                   ByRef Titulo As String, ByVal bAnuncio As Boolean)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Saves a new post into the forum.
      '***************************************************

10        With Foros(ForumIndex)
              
20            If bAnuncio Then
30                If .CantAnuncios < MAX_ANUNCIOS_FORO Then _
                      .CantAnuncios = .CantAnuncios + 1
                  
40                Call MoveArray(ForumIndex, bAnuncio)
                  
                  ' Agrego el anuncio
50                With .vsAnuncio(1)
60                    .sTitulo = Titulo
70                    .Autor = Autor
80                    .sPost = Post
90                End With
                  
100           Else
110               If .CantPosts < MAX_MENSAJES_FORO Then _
                      .CantPosts = .CantPosts + 1
                      
120               Call MoveArray(ForumIndex, bAnuncio)
                  
                  ' Agrego el post
130               With .vsPost(1)
140                   .sTitulo = Titulo
150                   .Autor = Autor
160                   .sPost = Post
170               End With
              
180           End If
190       End With
End Sub

Public Sub SaveForums()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Saves all forums into disk.
      '***************************************************
          Dim ForumIndex As Integer

10        For ForumIndex = 1 To NumForos
20            Call SaveForum(ForumIndex)
30        Next ForumIndex
End Sub


Private Sub SaveForum(ByVal ForumIndex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Saves a forum into disk.
      '***************************************************

          Dim PostIndex As Integer
          Dim FileIndex As Integer
          Dim PostPath As String
          
10        Call CleanForum(ForumIndex)
          
20        With Foros(ForumIndex)
              
              ' Guardo info del foro
30            Call WriteVar(App.Path & "\Foros\" & .Id & ".for", "INFO", "CantMSG", .CantPosts)
40            Call WriteVar(App.Path & "\Foros\" & .Id & ".for", "INFO", "CantAnuncios", .CantAnuncios)
              
              ' Guardo posts
50            For PostIndex = 1 To .CantPosts
                  
60                PostPath = App.Path & "\Foros\" & .Id & PostIndex & ".for"
70                FileIndex = FreeFile()
80                Open PostPath For Output As FileIndex
                  
90                With .vsPost(PostIndex)
100                   Print #FileIndex, .sTitulo
110                   Print #FileIndex, .Autor
120                   Print #FileIndex, .sPost
130               End With
                  
140               Close #FileIndex
                  
150           Next PostIndex
              
              ' Guardo Anuncios
160           For PostIndex = 1 To .CantAnuncios
                  
170               PostPath = App.Path & "\Foros\" & .Id & PostIndex & "a.for"
180               FileIndex = FreeFile()
190               Open PostPath For Output As FileIndex
                  
200               With .vsAnuncio(PostIndex)
210                   Print #FileIndex, .sTitulo
220                   Print #FileIndex, .Autor
230                   Print #FileIndex, .sPost
240               End With
                  
250               Close #FileIndex

260           Next PostIndex
              
270       End With
          
End Sub

Public Sub CleanForum(ByVal ForumIndex As Integer)
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Cleans a forum from disk.
      '***************************************************
          Dim PostIndex As Integer
          Dim NumPost As Integer
          Dim ForumPath As String

10        With Foros(ForumIndex)
          
              ' Elimino todo
20            ForumPath = App.Path & "\Foros\" & .Id & ".for"
30            If FileExist(ForumPath, vbNormal) Then
          
40                NumPost = val(GetVar(ForumPath, "INFO", "CantMSG"))
                  
                  ' Elimino los post viejos
50                For PostIndex = 1 To NumPost
60                    Kill App.Path & "\Foros\" & .Id & PostIndex & ".for"
70                Next PostIndex
                  
                  
80                NumPost = val(GetVar(ForumPath, "INFO", "CantAnuncios"))
                  
                  ' Elimino los post viejos
90                For PostIndex = 1 To NumPost
100                   Kill App.Path & "\Foros\" & .Id & PostIndex & "a.for"
110               Next PostIndex
                  
                  
                  ' Elimino el foro
120               Kill App.Path & "\Foros\" & .Id & ".for"
          
130           End If
140       End With

End Sub

Public Function SendPosts(ByVal Userindex As Integer, ByRef ForoID As String) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Sends all the posts of a required forum
      '***************************************************
          
          Dim ForumIndex As Integer
          Dim PostIndex As Integer
          Dim bEsGm As Boolean
          
10        ForumIndex = GetForumIndex(ForoID)

20        If ForumIndex > 0 Then

30            With Foros(ForumIndex)
                  
                  ' Send General posts
40                For PostIndex = 1 To .CantPosts
50                    With .vsPost(PostIndex)
60                        Call WriteAddForumMsg(Userindex, eForumMsgType.ieGeneral, .sTitulo, .Autor, .sPost)
70                    End With
80                Next PostIndex
                  
                  ' Send Sticky posts
90                For PostIndex = 1 To .CantAnuncios
100                   With .vsAnuncio(PostIndex)
110                       Call WriteAddForumMsg(Userindex, eForumMsgType.ieGENERAL_STICKY, .sTitulo, .Autor, .sPost)
120                   End With
130               Next PostIndex
                  
140           End With
              
150           bEsGm = EsGM(Userindex)
              
              ' Caos?
160           If esCaos(Userindex) Or bEsGm Then
                  
170               ForumIndex = GetForumIndex(FORO_CAOS_ID)
                  
180               With Foros(ForumIndex)
                      
                      ' Send General Caos posts
190                   For PostIndex = 1 To .CantPosts
                      
200                       With .vsPost(PostIndex)
210                           Call WriteAddForumMsg(Userindex, eForumMsgType.ieCAOS, .sTitulo, .Autor, .sPost)
220                       End With
                          
230                   Next PostIndex
                      
                      ' Send Sticky posts
240                   For PostIndex = 1 To .CantAnuncios
250                       With .vsAnuncio(PostIndex)
260                           Call WriteAddForumMsg(Userindex, eForumMsgType.ieCAOS_STICKY, .sTitulo, .Autor, .sPost)
270                       End With
280                   Next PostIndex
                      
290               End With
300           End If
                  
              ' Caos?
310           If esArmada(Userindex) Or bEsGm Then
                  
320               ForumIndex = GetForumIndex(FORO_REAL_ID)
                  
330               With Foros(ForumIndex)
                      
                      ' Send General Real posts
340                   For PostIndex = 1 To .CantPosts
                      
350                       With .vsPost(PostIndex)
360                           Call WriteAddForumMsg(Userindex, eForumMsgType.ieREAL, .sTitulo, .Autor, .sPost)
370                       End With
                          
380                   Next PostIndex
                      
                      ' Send Sticky posts
390                   For PostIndex = 1 To .CantAnuncios
400                       With .vsAnuncio(PostIndex)
410                           Call WriteAddForumMsg(Userindex, eForumMsgType.ieREAL_STICKY, .sTitulo, .Autor, .sPost)
420                       End With
430                   Next PostIndex
                      
440               End With
450           End If
              
460           SendPosts = True
470       End If
          
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Returns true if the post is sticky.
      '***************************************************
10        Select Case ForumType
              Case eForumMsgType.ieCAOS_STICKY
20                EsAnuncio = True
                  
30            Case eForumMsgType.ieGENERAL_STICKY
40                EsAnuncio = True
                  
50            Case eForumMsgType.ieREAL_STICKY
60                EsAnuncio = True
                  
70        End Select
          
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
      '***************************************************
      'Author: ZaMa
      'Last Modification: 01/03/2010
      'Returns the forum alignment.
      '***************************************************
10        Select Case yForumType
              Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
20                ForumAlignment = eForumType.ieCAOS
                  
30            Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
40                ForumAlignment = eForumType.ieGeneral
                  
50            Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
60                ForumAlignment = eForumType.ieREAL
                  
70        End Select
          
End Function

Public Sub ResetForums()
      '***************************************************
      'Author: ZaMa
      'Last Modification: 22/02/2010
      'Resets forum info
      '***************************************************
10        ReDim Foros(1 To 1) As tForo
20        NumForos = 0
End Sub

Private Sub MoveArray(ByVal ForumIndex As Integer, ByVal Sticky As Boolean)
      Dim i As Long

10    With Foros(ForumIndex)
20        If Sticky Then
30            For i = .CantAnuncios To 2 Step -1
40                .vsAnuncio(i).sTitulo = .vsAnuncio(i - 1).sTitulo
50                .vsAnuncio(i).sPost = .vsAnuncio(i - 1).sPost
60                .vsAnuncio(i).Autor = .vsAnuncio(i - 1).Autor
70            Next i
80        Else
90            For i = .CantPosts To 2 Step -1
100               .vsPost(i).sTitulo = .vsPost(i - 1).sTitulo
110               .vsPost(i).sPost = .vsPost(i - 1).sPost
120               .vsPost(i).Autor = .vsPost(i - 1).Autor
130           Next i
140       End If
150   End With
End Sub
