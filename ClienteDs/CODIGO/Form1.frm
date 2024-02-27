VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mercado de Personajes (VENTA)"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Comandos"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar Lista"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info pj"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
WriteSendInfo List1.ListIndex + 1
End Sub

Private Sub Command2_Click()
List1.Clear
End Sub

Private Sub Command3_Click()
ShowConsoleMsg "Mercado> Bienvenido al sistema de Mercado AO! En este espacio te mostraremos los comandos útiles de este sistema.", 250, 250, 150, True
ShowConsoleMsg "-/POSTEAR Sirve para poner en venta el personaje que desees, ingresando la cantidad de oro, el nivel mínimo de intercambio y el personaje que va a recibir el dinero en caso de que alguien lo compre.", 250, 250, 150, False
ShowConsoleMsg "-/MERCADO Sirve para ver los personajes en el mercado.", 250, 250, 150, False
ShowConsoleMsg "-/SOLICITARPJ Sirve para enviar una notificación al usuario que tiene en venta su personaje para que acepte o no el intercambio.", 250, 250, 150, False
ShowConsoleMsg "-/CANCELARSOLICITUDPJ Sirve para cancelar el cambio de personaje con el usuario que solicito el cambio con el tuyo.", 250, 250, 150, False
ShowConsoleMsg "-/DENEGARSOLICITUD Sirve para cancelar el cambio de personaje cuando se lo envias al personaje que deseas cambiar.", 250, 250, 150, False
ShowConsoleMsg "-/COMPRAR Sirve para comprar el personaje por oro.", 250, 250, 150, False
ShowConsoleMsg "-/CAMBIAR Sirve para cambiar tu personaje con el que hayas solicitado previamente.", 250, 250, 150, False
ShowConsoleMsg "-/QUITARPJ Sirve para sacar el personaje del Mercado Ao.", 250, 250, 150, False
End Sub
