Attribute VB_Name = "modPrivateMessages"
' Reparado por Lorwik

Option Explicit

Public Sub AgregarMensaje(ByVal UserIndex As Integer, _
                          ByRef Autor As String, _
                          ByRef Mensaje As String)

    '<EhHeader>
    On Error GoTo AgregarMensaje_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Agrega un nuevo mensaje privado a un usuario online.
    '***************************************************
    Dim LoopC As Long

    With UserList(UserIndex)

        If .UltimoMensaje < MAX_PRIVATE_MESSAGES Then
            .UltimoMensaje = .UltimoMensaje + 1
        Else

            For LoopC = 1 To MAX_PRIVATE_MESSAGES - 1
                .Mensajes(LoopC) = .Mensajes(LoopC + 1)
            Next

        End If
        
        With .Mensajes(.UltimoMensaje)
            .Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"
            .Nuevo = True

        End With
        
        Call WriteConsoleMsg(UserIndex, "¡Has recibido un mensaje privado de un Game Master!", FontTypeNames.FONTTYPE_GM)

    End With

    '<EhFooter>
    Exit Sub

AgregarMensaje_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.AgregarMensaje " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub AgregarMensajeOFF(ByRef Destinatario As String, _
                             ByRef Autor As String, _
                             ByRef Mensaje As String)

    '<EhHeader>
    On Error GoTo AgregarMensajeOFF_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Agrega un nuevo mensaje privado a un usuario offline.
    '***************************************************
    Dim UltimoMensaje As Byte

    Dim Charfile      As String

    Dim Contenido     As String

    Dim LoopC         As Long

    Charfile = CharPath & Destinatario & ".chr"
    UltimoMensaje = CByte(GetVar(Charfile, "MENSAJES", "UltimoMensaje"))
    Contenido = UCase$(Autor) & ": " & Mensaje & " (" & Now & ")"

    If UltimoMensaje < MAX_PRIVATE_MESSAGES Then
        UltimoMensaje = UltimoMensaje + 1
    Else

        For LoopC = 1 To MAX_PRIVATE_MESSAGES - 1
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1))
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
        Next LoopC

    End If
        
    Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje, Contenido)
    Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje & "_NUEVO", 1)
    
    Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", UltimoMensaje)
    '<EhFooter>
    Exit Sub

AgregarMensajeOFF_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.AgregarMensajeOFF " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function TieneMensajesNuevos(ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo TieneMensajesNuevos_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Determina si el usuario tiene mensajes nuevos.
    '***************************************************
    Dim LoopC As Long

    For LoopC = 1 To MAX_PRIVATE_MESSAGES

        If UserList(UserIndex).Mensajes(LoopC).Nuevo Then
            TieneMensajesNuevos = True

            Exit Function

        End If

    Next LoopC
    
    TieneMensajesNuevos = False
    '<EhFooter>
    Exit Function

TieneMensajesNuevos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.TieneMensajesNuevos " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub GuardarMensajes(ByRef IUser As User, ByRef Manager As clsIniManager)

    '<EhHeader>
    On Error GoTo GuardarMensajes_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Guarda los mensajes del usuario.
    '***************************************************
    Dim LoopC As Long
    
    With IUser
        Call Manager.ChangeValue("MENSAJES", "UltimoMensaje", CStr(.UltimoMensaje))
        
        For LoopC = 1 To MAX_PRIVATE_MESSAGES
            Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC, .Mensajes(LoopC).Contenido)

            If .Mensajes(LoopC).Nuevo Then
                Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 1)
            Else
                Call Manager.ChangeValue("MENSAJES", "MSJ" & LoopC & "_NUEVO", 0)

            End If

        Next LoopC

    End With

    '<EhFooter>
    Exit Sub

GuardarMensajes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.GuardarMensajes " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub CargarMensajes(ByVal UserIndex As Integer, ByRef Manager As clsIniManager)

    '<EhHeader>
    On Error GoTo CargarMensajes_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Carga los mensajes del usuario.
    '***************************************************
    Dim LoopC As Long

    With UserList(UserIndex)
        .UltimoMensaje = val(Manager.GetValue("MENSAJES", "UltimoMensaje"))
        
        For LoopC = 1 To MAX_PRIVATE_MESSAGES

            With .Mensajes(LoopC)
                .Nuevo = val(Manager.GetValue("MENSAJES", "MSJ" & LoopC & "_NUEVO"))
                .Contenido = CStr(Manager.GetValue("MENSAJES", "MSJ" & LoopC))

            End With

        Next LoopC

    End With

    '<EhFooter>
    Exit Sub

CargarMensajes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.CargarMensajes " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub LimpiarMensajeSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo LimpiarMensajeSlot_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Limpia el un mensaje de un usuario online.
    '***************************************************
    With UserList(UserIndex).Mensajes(Slot)
        .Contenido = vbNullString
        .Nuevo = False

    End With

    '<EhFooter>
    Exit Sub

LimpiarMensajeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.LimpiarMensajeSlot " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub LimpiarMensajes(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo LimpiarMensajes_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Limpia los mensajes del slot.
    '***************************************************
    Dim LoopC As Long

    With UserList(UserIndex)
        .UltimoMensaje = 0
        
        For LoopC = 1 To MAX_PRIVATE_MESSAGES
            Call LimpiarMensajeSlot(UserIndex, LoopC)
        Next LoopC

    End With

    '<EhFooter>
    Exit Sub

LimpiarMensajes_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.LimpiarMensajes " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub BorrarMensaje(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo BorrarMensaje_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Borra un mensaje de un usuario.
    '***************************************************
    Dim LoopC As Long

    With UserList(UserIndex)

        If Slot > .UltimoMensaje Or Slot < 1 Then Exit Sub

        If Slot = .UltimoMensaje Then
            Call LimpiarMensajeSlot(UserIndex, Slot)
        Else

            For LoopC = Slot To MAX_PRIVATE_MESSAGES - 1
                .Mensajes(LoopC) = .Mensajes(LoopC + 1)
            Next LoopC

            Call LimpiarMensajeSlot(UserIndex, .UltimoMensaje)

        End If
        
        .UltimoMensaje = .UltimoMensaje - 1

    End With

    '<EhFooter>
    Exit Sub

BorrarMensaje_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.BorrarMensaje " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub BorrarMensajeOFF(ByVal UserName As String, ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo BorrarMensajeOFF_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 04/08/2011
    'Borra un mensaje de un usuario.
    '***************************************************
    Dim Charfile      As String

    Dim UltimoMensaje As Byte

    Dim LoopC         As Long

    Charfile = CharPath & UserName & ".chr"
    
    UltimoMensaje = GetVar(Charfile, "MENSAJES", "UltimoMensaje")
    
    If Slot > UltimoMensaje Or Slot < 1 Then Exit Sub
    
    If Slot = UltimoMensaje Then
        Call WriteVar(Charfile, "MENSAJES", "MSJ" & Slot, vbNullString)
        Call WriteVar(Charfile, "MENSAJES", "MSJ" & Slot & "_Nuevo", vbNullString)
    Else

        For LoopC = Slot To UltimoMensaje - 1
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1))
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", GetVar(Charfile, "MENSAJES", "MSJ" & LoopC + 1 & "_NUEVO"))
        Next LoopC

        Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje, vbNullString)
        Call WriteVar(Charfile, "MENSAJES", "MSJ" & UltimoMensaje & "_Nuevo", vbNullString)

    End If
    
    Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", UltimoMensaje - 1)
    '<EhFooter>
    Exit Sub

BorrarMensajeOFF_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.BorrarMensajeOFF " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub LimpiarMensajesOFF(ByVal UserName As String)

    '<EhHeader>
    On Error GoTo LimpiarMensajesOFF_Err

    '</EhHeader>

    '***************************************************
    'Author: Amraphen
    'Last Modification: 18/08/2011
    'Borra los mensajes de un usuario offline.
    '***************************************************
    Dim Charfile      As String

    Dim UltimoMensaje As Byte

    Dim LoopC         As Long

    Charfile = CharPath & UserName & ".chr"
    
    UltimoMensaje = GetVar(Charfile, "MENSAJES", "UltimoMensaje")
    
    If UltimoMensaje > 0 Then

        For LoopC = 1 To UltimoMensaje
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC, vbNullString)
            Call WriteVar(Charfile, "MENSAJES", "MSJ" & LoopC & "_NUEVO", vbNullString)
        Next LoopC
        
        Call WriteVar(Charfile, "MENSAJES", "UltimoMensaje", 0)

    End If

    '<EhFooter>
    Exit Sub

LimpiarMensajesOFF_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modPrivateMessages.LimpiarMensajesOFF " & "at line " & Erl
        
    '</EhFooter>
End Sub
