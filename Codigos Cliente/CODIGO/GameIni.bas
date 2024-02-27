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

Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
"GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long '//Disco.

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
    Cabecera.Desc = "Desterium AO by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim n As Integer
    Dim GameIni As tGameIni
    n = FreeFile
    Open App.path & "\init\Inicio.con" For Binary As #n
    Get #n, , MiCabecera
    
    Get #n, , GameIni
    
    Close #n
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
On Local Error Resume Next

Dim n As Integer
n = FreeFile
Open App.path & "\init\Inicio.con" For Binary As #n
Put #n, , MiCabecera
Put #n, , GameIniConfiguration
Close #n
End Sub
Function GetSerialNumber(strDrive As String) As Long '//Disco.
Dim SerialNum As Long
Dim res As Long
Dim Temp1 As String
Dim Temp2 As String
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
res = GetVolumeInformation(strDrive, Temp1, _
Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
GetSerialNumber = SerialNum
End Function
Public Function ReadOptionIni() As addSetupOption
Dim sOption As addSetupOption
Dim n As Integer
    n = FreeFile
    Open App.path & "\Init\Config.con" For Binary As #n
        Get #n, , MiCabecera
        Get #n, , sOption
    Close #n
    ReadOptionIni = sOption
End Function
