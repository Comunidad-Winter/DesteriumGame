Attribute VB_Name = "EventosDS"
Option Explicit

Public Const MAX_EVENT_SIMULTANEO As Byte = 5

Public Enum eModalityEvent
    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Aracnus = 4
    HombreLobo = 5
    Minotauro = 6
    Busqueda = 7
    Unstoppable = 8
    Invasion = 9
    Enfrentamientos = 10
End Enum

Public Function strModality(ByVal Modality As eModalityEvent) As String
10        Select Case Modality
              Case eModalityEvent.CastleMode
20                strModality = "CastleMode"
                  
30            Case eModalityEvent.DagaRusa
40                strModality = "DagaRusa"
                  
50            Case eModalityEvent.DeathMatch
60                strModality = "DeathMatch"
                  
70            Case eModalityEvent.Aracnus
80                strModality = "Aracnus"
                  
90            Case eModalityEvent.HombreLobo
100               strModality = "HombreLobo"
                  
110           Case eModalityEvent.Minotauro
120               strModality = "Minotauro"
                  
130           Case eModalityEvent.Busqueda
140               strModality = "Busqueda"
                  
150           Case eModalityEvent.Unstoppable
160               strModality = "Unstoppable"
              
170           Case eModalityEvent.Invasion
180               strModality = "Invasion"
              
190           Case eModalityEvent.Enfrentamientos
200               strModality = "Enfrentamientos"
              
210       End Select
End Function

Public Function ModalityByte(ByVal Modality As String) As String
10        Select Case UCase$(Modality)
              Case "CASTLEMODE"
20                ModalityByte = 1
                  
30            Case "DAGARUSA"
40                ModalityByte = 2
                  
50            Case "DEATHMATCH"
60                ModalityByte = 3
                  
70            Case "ARACNUS"
80                ModalityByte = 4
              
90            Case "HOMBRELOBO"
100               ModalityByte = 5
                  
110           Case "MINOTAURO"
120               ModalityByte = 6
                  
130           Case "BUSQUEDA"
140               ModalityByte = 7
              
150           Case "UNSTOPPABLE"
160               ModalityByte = 8
              
170           Case "INVASION"
180               ModalityByte = 9
              
190           Case "1VS1", "2VS2", "3VS3", "4VS4", "5VS5", "6VS6", "7VS7", "8VS8", _
                  "9VS9", "10VS10", "11VS11", "12VS12", "13VS13", "14VS14", "15VS15", _
                  "20VS20", "25VS25"
200               ModalityByte = 10
210           Case Else
220               ModalityByte = 0
230       End Select
End Function

