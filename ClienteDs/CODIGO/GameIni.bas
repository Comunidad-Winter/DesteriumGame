Attribute VB_Name = "GameIni"
'Desterium AO 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Desterium AO is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit



Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
    bNoSoundEffects As Boolean
    sGraficos   As String * 13
    bGuildNews  As Boolean ' 11/19/09
    bDie        As Boolean ' 11/23/09 - FragShooter
    bKill       As Boolean ' 11/23/09 - FragShooter
    byMurderedLevel As Byte ' 11/23/09 - FragShooter
    bActive     As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
    'Aca iria lo demas
End Type

Public Type addSetupOption
    bGameCombat     As Boolean
    bFPS            As Boolean
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni
Public TSetup As addSetupOption

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
10        Cabecera.Desc = _
              "Desterium AO by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
20        Cabecera.CRC = Rnd * 100
30        Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
          Dim n As Integer
          Dim GameIni As tGameIni
10        n = FreeFile
20        Open App.path & "\init\Inicio.con" For Binary As #n
30        Get #n, , MiCabecera
          
40        Get #n, , GameIni
          
50        Close #n
60        LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
10    On Local Error Resume Next

      Dim n As Integer
20    n = FreeFile
30    Open App.path & "\init\Inicio.con" For Binary As #n
40    Put #n, , MiCabecera
50    Put #n, , GameIniConfiguration
60    Close #n
End Sub

Public Function ReadOptionIni() As addSetupOption
      Dim sOption As addSetupOption
      Dim n As Integer
10        n = FreeFile
20        Open App.path & "\Init\Config.con" For Binary As #n
30            Get #n, , MiCabecera
40            Get #n, , sOption
50        Close #n
60        ReadOptionIni = sOption
End Function
