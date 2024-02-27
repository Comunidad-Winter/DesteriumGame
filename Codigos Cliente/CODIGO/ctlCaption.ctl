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
    Set font = Caption.font
End Property

Property Let font(ByVal nFont As font)
    Caption.font = nFont
    Sombra(0).font = nFont
    Sombra(1).font = nFont
    PropertyChanged "Font"
End Property

Property Get Text() As String
    Text = Caption.Caption
End Property

Property Let Text(ByVal strCaption As String)
    Caption.Caption = strCaption
    Sombra(0).Caption = strCaption
    Sombra(1).Caption = strCaption
    PropertyChanged "Text"
End Property

Private Sub UserControl_Resize()
    
    With Caption
        .Left = 25
        .Top = 25
        .Width = Width
        .Height = Height
    End With
    
    With Sombra(0)
        .Left = 50
        .Top = 50
        .Width = Width
        .Height = Height
    End With
    
    With Sombra(1)
        .Left = 0
        .Top = 0
        .Width = Width
        .Height = Height
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

On Error Resume Next

    Call PropBag.WriteProperty("Font", Caption.font, Caption.font)
    Call PropBag.WriteProperty("Text", Caption.Caption, "CaptionControl")
    Call PropBag.WriteProperty("Text", Sombra(0).Caption, "CaptionControl")
    Call PropBag.WriteProperty("Text", Sombra(1).Caption, "CaptionControl")

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

    Set font = PropBag.ReadProperty("Font", Caption.font)
    Caption.Caption = PropBag.ReadProperty("Text", "CaptionControl")
    Sombra(0).Caption = PropBag.ReadProperty("Text", "CaptionControl")
    Sombra(1).Caption = PropBag.ReadProperty("Text", "CaptionControl")

End Sub
