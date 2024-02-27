VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAntiDos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const MaximoConexionesPorIP = 5
Private m_coleccion As New Collection
Public Function MaxConexiones(sIp As String) As Boolean
       
      Dim i As Integer
      Dim oIp As New UserIpAdress
       
10    oIp.ip = sIp
20    oIp.Conexiones = 1
30    m_coleccion.Add oIp
       
40    For i = 1 To m_coleccion.Count
50        If m_coleccion.Item(i).ip = sIp Then
60            m_coleccion.Item(i).Conexiones = m_coleccion.Item(i).Conexiones + 1
70            MaxConexiones = (m_coleccion.Item(i).Conexiones >= MaximoConexionesPorIP)
80            Exit Function
90        End If
100   Next
       
110   MaxConexiones = False
       
End Function
 
Public Function RestarConexion(sIp As String)
10    On Error GoTo Errhandler
      Dim i As Integer
       
20    For i = 1 To m_coleccion.Count
30       If m_coleccion.Item(i).ip = sIp Then
40          m_coleccion.Item(i).Conexiones = m_coleccion.Item(i).Conexiones - 1
50          If m_coleccion.Item(i).Conexiones = 0 Then Call m_coleccion.Remove(i)
60          Exit Function
70       End If
80    Next
       
90    Exit Function
       
Errhandler:
100       Call LogError("Error en RestarConexion " & Err.Description)
         
End Function
