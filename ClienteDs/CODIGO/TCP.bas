Attribute VB_Name = "Mod_TCP"
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
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
10    PuedoQuitarFoco = True
      'PuedoQuitarFoco = Not frmEstadisticas.Visible And '                 Not frmGuildAdm.Visible And '                 Not frmGuildDetails.Visible And '                 Not frmGuildBrief.Visible And '                 Not frmGuildFoundation.Visible And '                 Not frmGuildLeader.Visible And '                 Not frmCharInfo.Visible And '                 Not frmGuildNews.Visible And '                 Not frmGuildSol.Visible And '                 Not frmCommet.Visible And '                 Not frmPeaceProp.Visible
      '
End Function

Sub Login()
10        If EstadoLogin = E_MODO.Dados Then
20            Call Protocol.WriteThrowDices
110       End If

    If EstadoLogin = e_NewAccount Then
        WriteNewAccount
    ElseIf EstadoLogin = e_ConnectAccount Then
        WriteLoginAccount
    ElseIf EstadoLogin = e_LoginCharAccount Then
        WriteLoginCharAccount
    ElseIf EstadoLogin = e_CreateCharAccount Then
        WriteCreateCharAccount
    ElseIf EstadoLogin = e_RemoveCharAccount Then
        WriteRemoveCharAccount
    ElseIf EstadoLogin = e_RecoverAccount Then
        WriteRecoverAccount
    ElseIf EstadoLogin = e_ChangePasswdAccount Then
        WriteChangePasswdAccount
    ElseIf EstadoLogin = e_Temporal Then
        WriteAddTemporal
    End If
    
120       DoEvents
         
130       Call FlushBuffer
End Sub
