VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserIpAdress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private m_ip As String
Private m_conexiones As Integer
Private m_personajes_creados As Integer
Public Property Get PersonajesCreados() As Integer
10    PersonajesCreados = m_personajes_creados
End Property
Public Property Let PersonajesCreados(cant As Integer)
10    m_personajes_creados = cant
End Property
Public Property Get ip() As String
10    ip = m_ip
End Property
Public Property Let ip(ByVal Address As String)
10    m_ip = Address
End Property
Public Property Let Conexiones(Cantidad As Integer)
10    m_conexiones = Cantidad
End Property
Public Property Get Conexiones() As Integer
10    Conexiones = m_conexiones
End Property
