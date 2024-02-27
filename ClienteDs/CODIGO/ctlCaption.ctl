VERSION 5.00
Begin VB.UserControl CaptionControl 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Sombra 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Sombra 
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "CaptionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get font() As font
10        Set font = Caption.font
End Property

Property Let font(ByVal nFont As font)
10        Caption.font = nFont
20        Sombra(0).font = nFont
30        Sombra(1).font = nFont
40        PropertyChanged "Font"
End Property

Property Get Text() As String
10        Text = Caption.Caption
End Property

Property Let Text(ByVal strCaption As String)
10        Caption.Caption = strCaption
20        Sombra(0).Caption = strCaption
30        Sombra(1).Caption = strCaption
40        PropertyChanged "Text"
End Property

Private Sub UserControl_Resize()
          
10        With Caption
20            .Left = 25
30            .Top = 25
40            .Width = Width
50            .Height = Height
60        End With
          
70        With Sombra(0)
80            .Left = 50
90            .Top = 50
100           .Width = Width
110           .Height = Height
120       End With
          
130       With Sombra(1)
140           .Left = 0
150           .Top = 0
160           .Width = Width
170           .Height = Height
180       End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

10    On Error Resume Next

20        Call PropBag.WriteProperty("Font", Caption.font, Caption.font)
30        Call PropBag.WriteProperty("Text", Caption.Caption, "CaptionControl")
40        Call PropBag.WriteProperty("Text", Sombra(0).Caption, "CaptionControl")
50        Call PropBag.WriteProperty("Text", Sombra(1).Caption, "CaptionControl")

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

10    On Error Resume Next

20        Set font = PropBag.ReadProperty("Font", Caption.font)
30        Caption.Caption = PropBag.ReadProperty("Text", "CaptionControl")
40        Sombra(0).Caption = PropBag.ReadProperty("Text", "CaptionControl")
50        Sombra(1).Caption = PropBag.ReadProperty("Text", "CaptionControl")

End Sub
