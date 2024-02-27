Attribute VB_Name = "modAnticheat"
Option Explicit

Public Enum eTipo
    BotonLanzar = 1
    BotonHechizos = 2
    BotonInventario = 3
    ListaLanzar = 4
    InventarioObj = 5
    InvObjHechizo = 6
    LanzarAutocurar = 7
End Enum


Public Sub DetectoAnticheat(ByVal userIndex As Integer, ByVal tipo As eTipo)
    Dim nString As String
    Select Case tipo
        Case eTipo.BotonLanzar
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Boton lanzar>>"
            
        Case eTipo.BotonHechizos
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Boton hechizos>>"
            
        Case eTipo.BotonInventario
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Boton inventario>>"
            Exit Sub
            
        Case eTipo.InventarioObj
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Intervalo de cambio al inventario>>"
            
        Case eTipo.InvObjHechizo
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Intervalo de cambio al inventario>>"
            
        Case eTipo.ListaLanzar
            nString = "El anticheat ha detectado macro de clicks en la actividad <<Intervalo entre elegir hechizo y boton lanzar>>"
            
    End Select
    
    SendData SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Anticheat> USUARIO: " & UserList(userIndex).Name & " Se ha detectado posible uso de cheats", FontTypeNames.FONTTYPE_BRONCE)
          
    nString = "******************************************" & vbCrLf & _
            nString & vbCrLf & "Usuario: " & UserList(userIndex).Name & vbCrLf & _
            "Fecha y hora: " & Date & " - " & time & vbCrLf & _
            "******************************************"
    
    Call LogAnalisisPatrones(nString)
End Sub


Private Sub LogAnalisisPatrones(ByVal nString As String)
      '***************************************************
      'Author: el_Santo43
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
          

          Dim nfile As Integer
20        nfile = FreeFile ' obtenemos un canal
30        Open App.Path & "\AntiCheatSanto.log" For Append Shared As #nfile
40        Print #nfile, nString
50        Print #nfile, ""
60        Close #nfile
          
70        Exit Sub

Errhandler:
End Sub
