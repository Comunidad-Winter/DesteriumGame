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
 
'Diseñado & Implementado por maTih.-
 
'Función del código: El objetivo es tener un array grande
'                    Con 2 elementos, en el elemento 0 [PRIMER]
'                    Vamos a guardar los datos originales
'                    Y en el segundo elemento la data encriptada.
 
Private Type s_Data
        Block_Memory()  As Long
End Type
 
Private Const MAXLEN    As Long = 15530
 
Private m_Password      As Long
Private m_Array(1)      As s_Data
 
Public Sub Initialize(ByRef arr_Memory() As Long, ByVal i_Password As Long)
       
      '
      ' @ Inicializa la clase.
       
      'Doy dimensiones para los arrays de datos encriptados y originales.
       
      Dim loopX   As Long
         
10        m_Password = i_Password
         
20        For loopX = 0 To 1
30            ReDim m_Array(loopX).Block_Memory(LBound(arr_Memory()) To _
                  UBound(arr_Memory())) As Long
40        Next loopX
         
          'Guardo el bloque de memoria para los datos originales.
50        For loopX = LBound(m_Array(0).Block_Memory()) To _
              UBound(m_Array(0).Block_Memory())
60            m_Array(0).Block_Memory(loopX) = arr_Memory(loopX)
70        Next loopX
         
          'Encripto y guardo el bloque de memoria para los datos encriptados.
80        For loopX = LBound(m_Array(1).Block_Memory()) To _
              UBound(m_Array(1).Block_Memory())
90            m_Array(1).Block_Memory(loopX) = Encrypt_Data(arr_Memory(loopX))
100       Next loopX
         
End Sub
 
Public Sub Decrypt_All_Data(ByRef return_Array() As Long)
       
      '
      ' @ Desencripta toda la información que contiene la clase y la devuelve en el array "return_Array"
       
10    With m_Array(1)
       
           Dim loopX  As Long
           
20         ReDim return_Array(LBound(.Block_Memory()) To UBound(.Block_Memory())) As _
               Long
           
           'Llena el array
30         For loopX = LBound(return_Array()) To UBound(return_Array())
40             return_Array(loopX) = Decrypt_Data(.Block_Memory(loopX))
50         Next loopX
       
60    End With
       
End Sub
 
Private Function Encrypt_Data(ByVal i_Data As Long) As Long
       
      '
      ' @ Codifica un dato.
       
10    Encrypt_Data = (i_Data Xor m_Password)
       
End Function
 
Public Function Decrypt_Data(ByVal i_Data As Long)
       
      '
      ' @ Descodifica un dato.
       
10    Decrypt_Data = (m_Password Xor i_Data)
       
End Function
 
Public Function Return_Element_Original(ByVal i_Element As Long) As Long
       
      '
      ' @ Devuelve un elemento del array de datos originales.
       
10        With m_Array(0)
               
20             If (i_Element > UBound(.Block_Memory())) Then Exit Function
               
30             Return_Element_Original = .Block_Memory(i_Element)
               
40        End With
       
End Function
 
Public Function Return_Element_Crypted(ByVal i_Element As Long) As Long
       
      '
      ' @ Devuelve un elemento del array de datos encriptados SIN DESENCRIPTAR.
       
10    If i_Element > UBound(m_Array(1).Block_Memory()) Then Exit Function
       
20    Return_Element_Crypted = m_Array(1).Block_Memory(i_Element)
       
End Function
 
Public Function Return_Element_Decrypted(ByVal i_Element As Long) As Long
       
      '
      ' @ Devuelve un elemento del array de datos encriptados, ya desencriptad.
       
10    If i_Element > UBound(m_Array(1).Block_Memory()) Then Exit Function
       
20    Return_Element_Decrypted = Decrypt_Data(m_Array(1).Block_Memory(i_Element))
       
End Function
 
Public Function Set_Element(ByVal i_Value As Long) As Long
       
      '
      ' @ Setea un elemento al array de datos originales y devuelve el índice en el array.
         
          'Busco un índice.
10        Set_Element = Find_Free_Index(0)
       
          'ÍNDICE INVÁLIDO?
20        If (Set_Element = -1) Then Exit Function
         
30        With m_Array(0)
40             .Block_Memory(Set_Element) = i_Value
50        End With
           
End Function
 
Private Function Find_Free_Index(ByVal i_Mode As Byte) As Long
       
      '
      ' @ Busca un slot libre en el array.
       
      Dim i   As Long
       
      'Recorro el array buscando un valor nulo (0)
10    For i = 1 To UBound(m_Array(i_Mode).Block_Memory())
          'Si lo encuentro salgo del bucle.
20        If m_Array(i_Mode).Block_Memory(i) = 0 Then Exit For
30    Next i
       
      'Si el valor encontrado es > al valor máximo del array entonces redimensiono el array y devuelvo el último slot.
       
40    If (i > UBound(m_Array(i_Mode).Block_Memory())) Then
         'Compruebo no pasar el máximo tamaño
50       If (UBound(m_Array(i_Mode).Block_Memory()) + 1) < MAXLEN Then
60          ReDim m_Array(i_Mode).Block_Memory(LBound(m_Array(i_Mode).Block_Memory()) _
                To UBound(m_Array(i_Mode).Block_Memory())) As Long
70          Find_Free_Index = UBound(m_Array(i_Mode).Block_Memory())
80       Else
            'Muy poco probable, pero si tenemos el array
            'Lleno devolvemos un índice inválido!
90          Find_Free_Index = -1
100      End If
110   Else
120      Find_Free_Index = i
130   End If
       
End Function





