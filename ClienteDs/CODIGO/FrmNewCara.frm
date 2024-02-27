VERSION 5.00
Begin VB.Form FrmNewCara 
   BorderStyle     =   0  'None
   Caption         =   "Cambio de cara"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form4"
   Picture         =   "FrmNewCara.frx":0000
   ScaleHeight     =   2265
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicHead 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2640
      Picture         =   "FrmNewCara.frx":CC26
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   600
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2160
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image ImgCerrar 
      Height          =   255
      Left            =   5040
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   2760
      Picture         =   "FrmNewCara.frx":10104
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   3000
      Picture         =   "FrmNewCara.frx":10416
      Top             =   1320
      Width           =   240
   End
End
Attribute VB_Name = "FrmNewCara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IndexHead As Integer
Private Sub DirPJ_Click(Index As Integer)
10        If Index = 0 Then
20            Select Case UserRaza
                  Case eRaza.Humano
30                    If UserSexo = eGenero.Hombre Then
40                        If (IndexHead - 1) < LBound(HeadHombre.Humano()) Then Exit _
                              Sub
50                        IndexHead = IndexHead - 1
60                        Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Humano(IndexHead) & ".jpg")
                          
70                    Else
80                        If (IndexHead - 1) < LBound(HeadMujer.Humano()) Then Exit _
                              Sub
90                        IndexHead = IndexHead - 1
100                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Humano(IndexHead) & ".jpg")
110                   End If
                      
120               Case eRaza.Elfo
130                   If UserSexo = eGenero.Hombre Then
140                       If (IndexHead - 1) < LBound(HeadHombre.Elfo()) Then Exit Sub
150                       IndexHead = IndexHead - 1
160                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Elfo(IndexHead) & ".jpg")
                          
170                   Else
180                       If (IndexHead - 1) < LBound(HeadMujer.Elfo()) Then Exit Sub
190                       IndexHead = IndexHead - 1
200                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Elfo(IndexHead) & ".jpg")
210                   End If
220               Case eRaza.ElfoOscuro
230                   If UserSexo = eGenero.Hombre Then
240                       If (IndexHead - 1) < LBound(HeadHombre.ElfoDrow()) Then _
                              Exit Sub
250                       IndexHead = IndexHead - 1
260                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.ElfoDrow(IndexHead) & ".jpg")
                          
270                   Else
280                       If (IndexHead - 1) < LBound(HeadMujer.ElfoDrow()) Then Exit _
                              Sub
290                       IndexHead = IndexHead - 1
300                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.ElfoDrow(IndexHead) & ".jpg")
310                   End If
320               Case eRaza.Gnomo
330                   If UserSexo = eGenero.Hombre Then
340                       If (IndexHead - 1) < LBound(HeadHombre.Gnomo()) Then Exit _
                              Sub
350                       IndexHead = IndexHead - 1
360                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Gnomo(IndexHead) & ".jpg")
                          
370                   Else
380                       If (IndexHead - 1) < LBound(HeadMujer.Gnomo()) Then Exit Sub
390                       IndexHead = IndexHead - 1
400                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Gnomo(IndexHead) & ".jpg")
410                   End If
420               Case eRaza.Enano
430                   If UserSexo = eGenero.Hombre Then
440                       If (IndexHead - 1) < LBound(HeadHombre.Enano()) Then Exit _
                              Sub
450                       IndexHead = IndexHead - 1
460                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Enano(IndexHead) & ".jpg")
                          
470                   Else
480                       If (IndexHead - 1) < LBound(HeadMujer.Enano()) Then Exit Sub
490                       IndexHead = IndexHead - 1
500                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Enano(IndexHead) & ".jpg")
510                   End If
520           End Select
530       Else
540           Select Case UserRaza
                  Case eRaza.Humano
550                   If UserSexo = eGenero.Hombre Then
560                       If (IndexHead + 1) > UBound(HeadHombre.Humano()) Then Exit _
                              Sub
570                       IndexHead = IndexHead + 1
580                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Humano(IndexHead) & ".jpg")
                          
590                   Else
600                       If (IndexHead + 1) > UBound(HeadMujer.Humano()) Then Exit _
                              Sub
610                       IndexHead = IndexHead + 1
620                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Humano(IndexHead) & ".jpg")
630                   End If
                      
640               Case eRaza.Elfo
650                   If UserSexo = eGenero.Hombre Then
660                       If (IndexHead + 1) > UBound(HeadHombre.Elfo()) Then Exit Sub
670                       IndexHead = IndexHead + 1
680                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Elfo(IndexHead) & ".jpg")
                          
690                   Else
700                       If (IndexHead + 1) > UBound(HeadMujer.Elfo()) Then Exit Sub
710                       IndexHead = IndexHead + 1
720                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Elfo(IndexHead) & ".jpg")
730                   End If
740               Case eRaza.ElfoOscuro
750                   If UserSexo = eGenero.Hombre Then
760                       If (IndexHead + 1) > UBound(HeadHombre.ElfoDrow()) Then _
                              Exit Sub
770                       IndexHead = IndexHead + 1
780                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.ElfoDrow(IndexHead) & ".jpg")
                          
790                   Else
800                       If (IndexHead + 1) > UBound(HeadMujer.ElfoDrow()) Then Exit _
                              Sub
810                       IndexHead = IndexHead + 1
820                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.ElfoDrow(IndexHead) & ".jpg")
830                   End If
840               Case eRaza.Gnomo
850                   If UserSexo = eGenero.Hombre Then
860                       If (IndexHead + 1) > UBound(HeadHombre.Gnomo()) Then Exit _
                              Sub
870                       IndexHead = IndexHead + 1
880                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Gnomo(IndexHead) & ".jpg")
                          
890                   Else
900                       If (IndexHead + 1) > UBound(HeadMujer.Gnomo()) Then Exit Sub
910                       IndexHead = IndexHead + 1
920                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Gnomo(IndexHead) & ".jpg")
930                   End If
940               Case eRaza.Enano
950                   If UserSexo = eGenero.Hombre Then
960                       If (IndexHead + 1) > UBound(HeadHombre.Enano()) Then Exit _
                              Sub
970                       IndexHead = IndexHead + 1
980                       Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadHombre.Enano(IndexHead) & ".jpg")
                          
990                   Else
1000                      If (IndexHead + 1) > UBound(HeadMujer.Enano()) Then Exit Sub
1010                      IndexHead = IndexHead + 1
1020                      Me.PicHead.Picture = LoadPicture(DirGraficos & "\HEAD\" & _
                              HeadMujer.Enano(IndexHead) & ".jpg")
1030                  End If
1040          End Select
1050      End If
End Sub

Private Sub Form_Load()
10        IndexHead = 1
End Sub

Private Sub Image1_Click()
10        If UserSexo = eGenero.Hombre Then
20            Select Case UserRaza
                  Case eRaza.Humano
30                    WriteHead HeadHombre.Humano(IndexHead)
40                Case eRaza.Elfo
50                    WriteHead HeadHombre.Elfo(IndexHead)
60                Case eRaza.ElfoOscuro
70                    WriteHead HeadHombre.ElfoDrow(IndexHead)
80                Case eRaza.Gnomo
90                    WriteHead HeadHombre.Gnomo(IndexHead)
100               Case eRaza.Enano
110                   WriteHead HeadHombre.Enano(IndexHead)
120           End Select
130       Else
140           Select Case UserRaza
                  Case eRaza.Humano
150                   WriteHead HeadMujer.Humano(IndexHead)
160               Case eRaza.Elfo
170                   WriteHead HeadMujer.Elfo(IndexHead)
180               Case eRaza.ElfoOscuro
190                   WriteHead HeadMujer.ElfoDrow(IndexHead)
200               Case eRaza.Gnomo
210                   WriteHead HeadMujer.Gnomo(IndexHead)
220               Case eRaza.Enano
230                   WriteHead HeadMujer.Enano(IndexHead)
240           End Select
250       End If
          
260       Unload Me
End Sub

Private Sub imgCerrar_Click()
10        Unload Me
End Sub

