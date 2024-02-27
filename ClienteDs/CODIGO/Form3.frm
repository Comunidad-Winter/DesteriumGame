VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "CAPTIONS DE "
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
    Lib "user32" ( _
        ByVal hwnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
    Lib "user32" _
    Alias "GetWindowTextLengthA" ( _
        ByVal hwnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
    Lib "user32" _
    Alias "GetWindowTextA" ( _
        ByVal hwnd As Long, _
        ByVal lpString As String, _
        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
    Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public CANTv As Byte

Public Function Listar() As String
Static alter As String
Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hwnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                'La agregamos al ListBox
                Listar = titulo & "#" & Listar
                CANTv = CANTv + 1
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
       Loop
End Function

Private Sub Form_Load()

Debug.Print Listar
Dim raa As String
raa = Listar
Debug.Print raa

End Sub
