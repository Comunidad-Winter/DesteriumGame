VERSION 5.00
Begin VB.Form FrmPjsMao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "FrmPjsMao.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTittle 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3960
      Width           =   4335
   End
   Begin VB.TextBox txtDsp 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "0"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtGld 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ListBox lstPjs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1320
   End
   Begin VB.ListBox lstCopyPjs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   2640
      TabIndex        =   2
      Top             =   1560
      Width           =   1320
   End
   Begin VB.ListBox lstMercado 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4125
      Left            =   7800
      TabIndex        =   0
      Top             =   1800
      Width           =   3720
   End
   Begin VB.Image imgOffer 
      Height          =   615
      Left            =   7920
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Image ImgInfo 
      Height          =   375
      Left            =   8280
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Image ImgAddMao 
      Height          =   375
      Left            =   4560
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblValor 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   730
      TabIndex        =   1
      Top             =   3420
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   9240
      Top             =   8520
      Width           =   2535
   End
End
Attribute VB_Name = "FrmPjsMao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Image1_Click()
    SelectedListMAO = 0
10    Unload Me
      'FrmMercado.SetFocus
End Sub


Private Sub imgInfo_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        If lstMercado.ListIndex < 0 Then
30            MsgBox "Selecciona una publicación previamente. También podes utilizar Doble Click."
40            Exit Sub
50        End If

          If lstMercado.List(lstMercado.ListIndex) = "(Vacio)" Then
            MsgBox "Selecciona una publicación que no esté vacía"
            Exit Sub
          End If
          
          SelectedListMAO = lstMercado.ListIndex + 1
60        Call WriteRequestInfoMAO
70        Unload Me
End Sub

Private Sub imgComprar_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        If lstPjs.ListIndex = -1 Then
30            MsgBox "Selecciona un personaje"
40            Exit Sub
50        End If
          
60        If lblValor.Caption = "X CAMBIO" Then
70            MsgBox "El personaje con el que intentas cambiar está en MODO CAMBIO"
80            Exit Sub
90        End If

100       If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
110           If MsgBox("¿Seguro que desea comprar el personaje " & _
                  lstPjs.List(lstPjs.ListIndex) & "?", vbYesNo + vbQuestion, _
                  "Mercado Desterium") = vbYes Then
120               Protocol.WriteBuyPj (UCase$(lstPjs.List(lstPjs.ListIndex)))
130           End If
140       End If
          
End Sub

Private Sub ImgOfrecer_Click()
10        Call Audio.PlayWave(SND_CLICK)
          
20        If lstPjs.ListIndex = -1 Then
30            MsgBox "Selecciona un personaje"
40            Exit Sub
50        End If
          
60        If Not lblValor.Caption = "X CAMBIO" Then
70            MsgBox "El personaje con el que intentas cambiar no está en MODO CAMBIO"
80            Exit Sub
90        End If
          
100       If lstPjs.List(lstPjs.ListIndex) <> vbNullString Then
110           If _
                  MsgBox("¿Seguro que deseas ofertar tu personaje a cambio de este? Tu contraseña/pin/email pasarán a ser los datos del personaje recibido.", _
                  vbYesNo + vbQuestion, "Desterium AO") = vbYes Then
                  Dim strPin As String
120               strPin = _
                      InputBox("Escriba el pin de su personaje. Recuerde la distinción entre mayúsculas y minusculas.")
                  
130               If strPin = vbNullString Then Exit Sub
                  
140               Call _
                      Protocol.WriteMercadoInvitation(UCase$(lstPjs.List(lstPjs.ListIndex)), _
                      strPin)
150           Else
160               Exit Sub
170           End If
180       End If

End Sub

Private Sub ImgAddMao_Click()

    On Err GoTo ErrHandler
    
        Dim A As Long
        Dim TempS As String
    
        If lstCopyPjs.ListCount <= 0 Then
            MsgBox "Tienes que seleccionar al menos 1 personaje"
            Exit Sub
        End If
        
        If Val(txtGld.Text) < 0 Then
            MsgBox "El valor de las Monedas de Oro debe ser positivo o cero."
            Exit Sub
        End If
        
        If Val(txtDsp.Text) < 0 Then
            MsgBox "El valor de las monedas Dsp debe ser positivo o cero."
            Exit Sub
        End If
        
        If Len(txtTittle.Text) < 8 Then
            MsgBox "El título debe contener 8 carácteres como mínimo. Ten en cuenta que es una linea por publicación."
            Exit Sub
        End If
        
        For A = 0 To lstCopyPjs.ListCount - 1
            TempS = TempS & lstCopyPjs.List(A) & "-"
        Next A
        
        TempS = mid$(TempS, 1, Len(TempS) - 1)
        
        
        
        WritePublicationMAO Val(txtGld.Text), Val(txtDsp.Text), TempS, txtTittle.Text, 0
        Unload Me
        
ErrHandler:
    'Did detected an invalid message??
70        If Err.number = CustomMessages.InvalidMessageErrCode Then
80            Call MsgBox("El Mensaje " & TempS & " es inválido. Modifiquelo por favor.")
90        End If
    
End Sub

Private Sub imgOffer_Click()
    FrmOfferMao.Show
    Unload Me
End Sub

Private Sub lstCopyPjs_dblClick()
    
    lstCopyPjs.RemoveItem lstCopyPjs.ListIndex
    
End Sub

Private Sub lstPjs_dblClick()
    Dim SelectedPj As String
    Dim A As Long
    
    SelectedPj = lstPjs.List(lstPjs.ListIndex)
    
    For A = 0 To lstCopyPjs.ListCount - 1
        If UCase$(SelectedPj) = lstCopyPjs.List(A) Then
            MsgBox "El personaje ya está agregado"
            Exit Sub
        End If
    Next A
    
    lstCopyPjs.AddItem SelectedPj
    
End Sub

