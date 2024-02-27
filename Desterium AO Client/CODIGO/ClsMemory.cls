VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Private Type s_Data
        Block_Memory()  As Long
End Type
 
Private Const MAXLEN    As Long = 10240
 
Private m_Password      As Long
Private m_Array(1)      As s_Data
 
Public Sub Initialize(ByRef arr_Memory() As Long, ByVal i_Password As Long)
 
'
' @ Inicializa la clase.
 
'Doy dimensiones para los arrays _
de datos encriptados y originales.
 
Dim loopX   As Long
   
    m_Password = i_Password
   
    For loopX = 0 To 1
        ReDim m_Array(loopX).Block_Memory(LBound(arr_Memory()) To UBound(arr_Memory())) As Long
    Next loopX
   
    'Guardo el bloque de memoria para los datos originales.
    For loopX = LBound(m_Array(0).Block_Memory()) To UBound(m_Array(0).Block_Memory())
        m_Array(0).Block_Memory(loopX) = arr_Memory(loopX)
    Next loopX
   
    'Encripto y guardo el bloque de memoria para los datos encriptados.
    For loopX = LBound(m_Array(1).Block_Memory()) To UBound(m_Array(1).Block_Memory())
        m_Array(1).Block_Memory(loopX) = Encrypt_Data(arr_Memory(loopX))
    Next loopX
   
End Sub
 
Public Sub Decrypt_All_Data(ByRef return_Array() As Long)
 
'
' @ Desencripta toda la información que contiene la clase _
y la devuelve en el array "return_Array"
 
With m_Array(1)
 
     Dim loopX  As Long
     
     ReDim return_Array(LBound(.Block_Memory()) To UBound(.Block_Memory())) As Long
     
     'Llena el array
     For loopX = LBound(return_Array()) To UBound(return_Array())
         return_Array(loopX) = Decrypt_Data(.Block_Memory(loopX))
     Next loopX
 
End With
 
End Sub
 
Private Function Encrypt_Data(ByVal i_Data As Long) As Long
 
'
' @ Codifica un dato.
 
Encrypt_Data = (i_Data Xor m_Password)
 
End Function
 
Public Function Decrypt_Data(ByVal i_Data As Long)
 
'
' @ Descodifica un dato.
 
Decrypt_Data = (m_Password Xor i_Data)
 
End Function
 
Public Function Return_Element_Original(ByVal i_Element As Long) As Long
 
'
' @ Devuelve un elemento del array de datos originales.
 
    With m_Array(0)
         
         If (i_Element > UBound(.Block_Memory())) Then Exit Function
         
         Return_Element_Original = .Block_Memory(i_Element)
         
    End With
 
End Function
 
Public Function Return_Element_Crypted(ByVal i_Element As Long) As Long
 
'
' @ Devuelve un elemento del array de datos encriptados _
    SIN DESENCRIPTAR.
 
If i_Element > UBound(m_Array(1).Block_Memory()) Then Exit Function
 
Return_Element_Crypted = m_Array(1).Block_Memory(i_Element)
 
End Function
 
Public Function Return_Element_Decrypted(ByVal i_Element As Long) As Long
 
'
' @ Devuelve un elemento del array de datos encriptados, ya desencriptad.
 
If i_Element > UBound(m_Array(1).Block_Memory()) Then Exit Function
 
Return_Element_Decrypted = Decrypt_Data(m_Array(1).Block_Memory(i_Element))
 
End Function
 
Public Function Set_Element(ByVal i_Value As Long) As Long
 
'
' @ Setea un elemento al array de datos originales _
    y devuelve el índice en el array.
   
    'Busco un índice.
    Set_Element = Find_Free_Index(0)
 
    'ÍNDICE INVÁLIDO?
    If (Set_Element = -1) Then Exit Function
   
    With m_Array(0)
         .Block_Memory(Set_Element) = i_Value
    End With
     
End Function
 
Private Function Find_Free_Index(ByVal i_Mode As Byte) As Long
 
'
' @ Busca un slot libre en el array.
 
Dim i   As Long
 
'Recorro el array buscando un valor nulo (0)
For i = 1 To UBound(m_Array(i_Mode).Block_Memory())
    'Si lo encuentro salgo del bucle.
    If m_Array(i_Mode).Block_Memory(i) = 0 Then Exit For
Next i
 
'Si el valor encontrado es > al valor máximo del array _
entonces redimensiono el array y devuelvo el último slot.
 
If (i > UBound(m_Array(i_Mode).Block_Memory())) Then
   'Compruebo no pasar el máximo tamaño
   If (UBound(m_Array(i_Mode).Block_Memory()) + 1) < MAXLEN Then
      ReDim m_Array(i_Mode).Block_Memory(LBound(m_Array(i_Mode).Block_Memory()) To UBound(m_Array(i_Mode).Block_Memory())) As Long
      Find_Free_Index = UBound(m_Array(i_Mode).Block_Memory())
   Else
      'Muy poco probable, pero si tenemos el array
      'Lleno devolvemos un índice inválido!
      Find_Free_Index = -1
   End If
Else
   Find_Free_Index = i
End If
 
End Function



