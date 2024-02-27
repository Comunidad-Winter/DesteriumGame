VERSION 5.00
Begin VB.Form Vercaptions 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desterium AO"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAPTIONS DE: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Vercaptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As _
    Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength Lib "user32" Alias _
    "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal _
    wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public CANTv As Byte

Public Function Listar() As String
      Static alter As String
      Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long
          'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
10        handle = GetWindow(hWnd, GW_HWNDFIRST)

20        CANTv = 0
          'Este bucle va a recorrer todas las ventanas.
          'cuando GetWindow devielva un 0, es por que no hay mas
30        Do While handle <> 0
              'Tenemos que comprobar que la ventana es una de tipo visible
40            If IsWindowVisible(handle) Then
                  'Obtenemos el número de caracteres de la ventana
50                lenT = GetWindowTextLength(handle)
                  'si es el número anterior es mayor a 0
60                If lenT > 0 Then
                      'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
70                    titulo = String$(lenT, 0)
                      'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                      'y también debemos pasarle el Hwnd de dicha ventana
80                    ret = GetWindowText(handle, titulo, lenT + 1)
90                    titulo$ = Left$(titulo, ret)
                      'La agregamos al ListBox
100                   Listar = titulo & "#" & Listar
110                   CANTv = CANTv + 1
120               End If
130           End If
              'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
140           handle = GetWindow(handle, GW_HWNDNEXT)
150          Loop
End Function

Private Sub Command1_Click()
10    Unload Me
End Sub

Private Sub Form_Load()

10    Debug.Print Listar
      Dim raa As String
20    raa = Listar
30    Debug.Print raa

End Sub

