Attribute VB_Name = "mMascotas"
Option Explicit

Private Const ARCHIVE As String = "MASCOTAS.DAT"

Private Type tMascota

    Name As String
    MinHit As Integer
    MaxHit As Integer
    MinHitMag As Integer
    MaxHitMag As Integer
    Spells(1 To 35) As Integer
    
    SoloMagia As Boolean
    SoloGolpe As Boolean

End Type

Public Mascotas() As tMascota

Public Function Mascota_Index(ByVal UserIndex As Integer) As Integer

    '<EhHeader>
    On Error GoTo Mascota_Index_Err

    '</EhHeader>

    With UserList(UserIndex)

        'Druidas
        If .Clase = eClass.Druid Then

            Select Case .Raza

                Case eRaza.Humano, eRaza.Gnomo, eRaza.Enano
                    Mascota_Index = 78

                    Exit Function

                Case eRaza.Elfo, eRaza.Drow
                    Mascota_Index = 96

                    Exit Function

            End Select

        End If
        
        ' Clerigos
        If .Clase = eClass.Cleric Then
            Mascota_Index = 92

            Exit Function

        End If
        
        ' Bardos
        If .Clase = eClass.Bard Then
            Mascota_Index = 94

            Exit Function

        End If
        
        ' Magos
        If .Clase = eClass.Mage Then
            Mascota_Index = 93

            Exit Function

        End If
        
        ' Paladines-Asesinos
        If .Clase = eClass.Assasin Or .Clase = eClass.Paladin Then
            Mascota_Index = 115

        End If
    
    End With
    
    '<EhFooter>
    Exit Function

Mascota_Index_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMascotas.Mascota_Index " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub Mascotas_AddNew(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '<EhHeader>
    On Error GoTo Mascotas_AddNew_Err

    '</EhHeader>
    
    Dim Slot As Byte

    Dim Obj  As Obj

    With UserList(UserIndex)
        
        'Slot = Mascota_FreeSlot(UserIndex)
        
        WriteConsoleMsg UserIndex, "Sabemos lo importante que es este sistema para vos. Te prometemos un nuevo sistema de domar mascotas, donde podrás entrenarlas. Estamos trabajando en ello. Mientras tanto tendrás los hechizos para invocar mascotas momentaneas.", FontTypeNames.FONTTYPE_INFO

        Exit Sub
        
        If Slot = 0 Then

            'WriteConsoleMsg UserIndex, "No tienes lugar para mas mascotas.", FontTypeNames.FONTTYPE_INFO
            'Exit Sub
        End If
        
        If RandomNumber(1, 100) <= 77 Then
            WriteConsoleMsg UserIndex, "No has logrado domar a la criatura.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
        Obj.ObjIndex = Npclist(NpcIndex).MonturaIndex
        Obj.Amount = 1
        
        If Not MeterItemEnInventario(UserIndex, Obj) Then
            WriteConsoleMsg UserIndex, "No tienes lugar en tu inventario. SI o SI debes tenerla en él.", FontTypeNames.FONTTYPE_INFO

            Exit Sub

        End If
        
    End With

    '<EhFooter>
    Exit Sub

Mascotas_AddNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMascotas.Mascotas_AddNew " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DoEquita(ByVal UserIndex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)

    '<EhHeader>
    On Error GoTo DoEquita_Err

    '</EhHeader>

    With UserList(UserIndex)
        
        If .flags.Montando = 0 Then
            .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.MonturaSlot = Slot
            .Char.Head = 0

            If .flags.Muerto = 0 Then
                .Char.Body = Montura.Ropaje
            Else
                .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
                .Char.Head = iCabezaMuerto(Escriminal(UserIndex))

            End If

            .Char.Head = UserList(UserIndex).OrigChar.Head
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = .Char.CascoAnim
            .flags.Montando = 1
            .Invent.Object(Slot).Equipped = 1
        Else

            If .Invent.MonturaObjIndex <> .Invent.Object(Slot).ObjIndex Then
                Call WriteConsoleMsg(UserIndex, "Esta no es la montura a la que estabas subido.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If
            
            .Invent.Object(Slot).Equipped = 0
            .flags.Montando = 0
            
            If .flags.Muerto = 0 Then
                .Char.Head = UserList(UserIndex).OrigChar.Head

                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
                Else
                    Call DarCuerpoDesnudo(UserIndex)

                End If

                If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
                If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
            Else
                .Char.Body = iCuerpoMuerto(Escriminal(UserIndex))
                .Char.Head = iCabezaMuerto(Escriminal(UserIndex))
                    
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                      
                Dim A As Long
                      
                For A = 1 To MAX_AURAS
                    .Char.AuraIndex(A) = NingunAura
                Next A

            End If

        End If
      
        Call WriteChangeInventorySlot(UserIndex, Slot)
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        Call WriteMontateToggle(UserIndex)

    End With

    '<EhFooter>
    Exit Sub

DoEquita_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mMascotas.DoEquita " & "at line " & Erl

    '</EhFooter>
End Sub
