Attribute VB_Name = "Mod_Balance"
Option Explicit

Private Const MAX_NIVEL_VIDA_PROMEDIO = 17

Public Type tRango

    minimo As Integer
    maximo As Integer

End Type

Private Type tBalance

    Hp As Integer
    Man As Integer

End Type

Public BalanceStats(1 To NUMCLASES, 1 To NUMRAZAS) As tBalance

' Adicionales de Vida
Public Const AdicionalHPGuerrero = 2 'HP adicionales cuando sube de nivel

Public Const AdicionalHPCazador = 1

Public Const AdicionalSTLadron = 3

Public Const AdicionalSTLeñador = 23

Public Const AdicionalSTPescador = 20

Public Const AdicionalSTMinero = 25

Public Const AumentoSTDef        As Byte = 15

Public Const AumentoStBandido    As Byte = AumentoSTDef + 23

Public Const AumentoSTLadron     As Byte = AumentoSTDef + 3

Public Const AumentoSTMago       As Byte = AumentoSTDef - 1

Public Const AumentoSTTrabajador As Byte = AumentoSTDef + 25

' El balance LVL TEMP sería el nivel "1" del personaje. En base al balance default del juego (TDS)
' Si se suma 32 + los 15 niveles que el juego tiene da 47. No es un número al azar
Public Const BALANCE_LVL_TEMP    As Byte = 0

Public Const POCION_ROJA_NEWBIE = 205

Public Const POCION_AZUL_NEWBIE = 206

Public Const POCION_AMARILLA_NEWBIE = 207

Public Const POCION_VERDE_NEWBIE = 208

Public Const VESTIMENTA_WAR_NEWBIE = 203

Public Const VESTIMENTA_MAG_NEWBIE = 204

Public Const DAGA_NEWBIE = 202

Public Sub LoadSetInitial_Class(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo LoadSetInitial_Class_Err

    '</EhHeader>

    Dim UserClase As eClass

    Dim Slot      As Long

    With UserList(UserIndex)

        Call LimpiarInventario(UserIndex)

        'Pociones Rojas (Newbie)
        Slot = 1
        .Invent.Object(Slot).ObjIndex = POCION_ROJA_NEWBIE
        .Invent.Object(Slot).Amount = 150
        
        'Pociones azules (Newbie)
        If .Stats.MaxMan > 0 Then
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = POCION_AZUL_NEWBIE
            .Invent.Object(Slot).Amount = 100
          
        End If
            
        .Invent.Object(Slot).Amount = 10 'Pociones Amarillas
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = POCION_AMARILLA_NEWBIE
        
        'Pociones Amarillas
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = POCION_VERDE_NEWBIE
        .Invent.Object(Slot).Amount = 10
    
        Dim Escudo     As Obj

        Dim Casco      As Obj

        Dim Armadura   As Obj

        Dim Arma       As Obj

        Dim Anillo     As Obj

        Dim Municiones As Obj
    
        Escudo.Amount = 1
        Casco.Amount = 1
        Armadura.Amount = 1
        Arma.Amount = 1
        Anillo.Amount = 1
        Municiones.Amount = 1
             
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
        .Char.Body = 0
        .Char.ShieldAnim = NingunEscudo
         
        Select Case .Clase
    
            Case eClass.Mage
                Escudo.ObjIndex = 0
                Casco.ObjIndex = 173 ' Sombrero de Mago
                Armadura.ObjIndex = 91 'Tunica legendaria
                Arma.ObjIndex = 171 ' Báculo Engarzado
            
            Case eClass.Cleric
                Escudo.ObjIndex = 117 ' Escudo Imperial
                Casco.ObjIndex = 119 ' Casco de Hierro Completo
                Armadura.ObjIndex = 104 ' Placas de Acero
                Arma.ObjIndex = 141 ' Hacha Dos Filos
        
            Case eClass.Paladin
                Escudo.ObjIndex = 117 ' Escudo Imperial
                Casco.ObjIndex = 119 ' Casco de Hierro Completo
                Armadura.ObjIndex = 104 ' Placas de Acero
                Arma.ObjIndex = 142 ' Espada de Plata
        
            Case eClass.Warrior
                Escudo.ObjIndex = 117 ' Escudo Imperial
                Casco.ObjIndex = 119 ' Casco de Hierro Completo
                Armadura.ObjIndex = 104 ' Placas de Acero
                Arma.ObjIndex = 142 ' Espada de Plata
        
            Case eClass.Assasin
                Escudo.ObjIndex = 333 ' Escudo de Tortuga
                Casco.ObjIndex = 118 ' Casco de Hierro
                Armadura.ObjIndex = 105 ' Armadura Klox
                Arma.ObjIndex = 139 ' Puñal Infernal
            
            Case eClass.Bard
                Escudo.ObjIndex = 333 ' Escudo de Tortuga
                Casco.ObjIndex = 132 ' Casco de Hierro
                Armadura.ObjIndex = 91 'Tunica legendaria
                Arma.ObjIndex = 0
                Anillo.ObjIndex = 167 ' Laud Mágico
            
            Case eClass.Druid
                Escudo.ObjIndex = 0
                Casco.ObjIndex = 0
                Armadura.ObjIndex = 91 'Tunica legendaria
                Arma.ObjIndex = 0
                Anillo.ObjIndex = 168 ' Anillo Mágico
            
            Case eClass.Hunter
                Escudo.ObjIndex = 333 ' Escudo de Tortuga
                Casco.ObjIndex = 131 ' Capucha de Cazador
                Armadura.ObjIndex = 96 ' Armadura de Cazador
                Arma.ObjIndex = 212 ' Arco de Cazador
                Municiones.ObjIndex = 154 ' Flechas
        
            Case Else
            
        End Select

        If UserList(UserIndex).ServerSelected = 3 Then
            Armadura.ObjIndex = 216 ' Vestimentas Polar

        End If
            
        If Armadura.ObjIndex > 0 Then
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = Armadura.ObjIndex
    
            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
          
            .Invent.ArmourEqpSlot = Slot
            .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
        
        End If
    
        If Arma.ObjIndex > 0 Then
            Slot = Slot + 1

            .Invent.Object(Slot).ObjIndex = Arma.ObjIndex

            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
          
            .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.WeaponEqpSlot = Slot
          
            .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)
    
        End If
    
        If Municiones.ObjIndex > 0 Then
            Slot = Slot + 1

            .Invent.Object(Slot).ObjIndex = Municiones.ObjIndex

            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
          
            .Invent.MunicionEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.MunicionEqpSlot = Slot
    
        End If
            
        If Escudo.ObjIndex > 0 Then
            Slot = Slot + 1

            .Invent.Object(Slot).ObjIndex = Escudo.ObjIndex

            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
          
            .Invent.EscudoEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.EscudoEqpSlot = Slot
          
            .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

        End If
    
        If Casco.ObjIndex > 0 Then
            Slot = Slot + 1

            .Invent.Object(Slot).ObjIndex = Casco.ObjIndex

            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
          
            .Invent.CascoEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.CascoEqpSlot = Slot
          
            .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    
        End If
        
        ' Total Items
        .Invent.NroItems = Slot
        
        Call User_GenerateNewHead(UserIndex, 1)
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
        
        Call UpdateUserInv(True, UserIndex, 0)

        ' Spells
        Dim A As Long

        For A = 1 To MAXUSERHECHIZOS
            .Stats.UserHechizos(A) = 0
        Next A
            
        If .Stats.MaxMan > 0 Then
            .Stats.UserHechizos(35) = 10 'RemoverParalisis
            .Stats.UserHechizos(34) = 24 ' 'Inmovilizar
            .Stats.UserHechizos(33) = 9 ' 'Paralizar
            .Stats.UserHechizos(32) = 15 ' 'Tormenta de Fuego
            .Stats.UserHechizos(31) = 23 ' 'Descarga Eléctrica
            .Stats.UserHechizos(30) = 14 ' 'Invisibilidad
                
            If .Stats.MaxMan > 1000 Then
                .Stats.UserHechizos(29) = 25 ' 'Apocalipsis

            End If

        End If

        Call UpdateUserHechizos(True, UserIndex, 0)
        
    End With

    '<EhFooter>
    Exit Sub

LoadSetInitial_Class_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.LoadSetInitial_Class " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ApplySetInitial_Newbie(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ApplySetInitial_Newbie_Err

    '</EhHeader>

    With UserList(UserIndex)

        '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
        Dim Slot      As Byte

        Dim IsPaladin As Boolean
        
        IsPaladin = .Clase = eClass.Paladin
        
        'Pociones Rojas (Newbie)
        Slot = 1
        .Invent.Object(Slot).ObjIndex = POCION_ROJA_NEWBIE
        .Invent.Object(Slot).Amount = 150
        
        If POCION_AZUL_NEWBIE > 0 Then

            'Pociones azules (Newbie)
            If .Stats.MaxMan > 0 Or IsPaladin Then
                Slot = Slot + 1
                .Invent.Object(Slot).ObjIndex = POCION_AZUL_NEWBIE
                .Invent.Object(Slot).Amount = 100
              
            End If

        End If
        
        If POCION_AMARILLA_NEWBIE > 0 Then
            'Pociones Amarillas
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = POCION_AMARILLA_NEWBIE
            .Invent.Object(Slot).Amount = 10

        End If
        
        If POCION_VERDE_NEWBIE > 0 Then
            'Pociones Amarillas
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = POCION_VERDE_NEWBIE
            .Invent.Object(Slot).Amount = 10

        End If
        
        ' Ropa (Newbie)
        Slot = Slot + 1

        Select Case .Clase

            Case eClass.Assasin, eClass.Paladin, eClass.Hunter, eClass.Warrior, eClass.Thief
                .Invent.Object(Slot).ObjIndex = VESTIMENTA_WAR_NEWBIE

            Case eClass.Mage, eClass.Druid, eClass.Cleric, eClass.Bard
                .Invent.Object(Slot).ObjIndex = VESTIMENTA_MAG_NEWBIE

        End Select
        
        ' Equipo ropa
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1
        .flags.Desnudo = 0
        .Invent.ArmourEqpSlot = Slot
        .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
        'Call DarCuerpoDesnudo(UserIndex)
        .Char.Body = GetArmourAnim(UserIndex, .Invent.ArmourEqpObjIndex)
        
        If DAGA_NEWBIE > 0 Then
            'Arma (Newbie)
            Slot = Slot + 1
    
            .Invent.Object(Slot).ObjIndex = DAGA_NEWBIE
            
            ' Equipo arma
            .Invent.Object(Slot).Amount = 1
            .Invent.Object(Slot).Equipped = 1
        
            .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.WeaponEqpSlot = Slot
            
            .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Raza, .Invent.WeaponEqpObjIndex)

        End If

        ' Sin casco y escudo
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
          
        ' Total Items
        .Invent.NroItems = Slot

    End With

    '<EhFooter>
    Exit Sub

ApplySetInitial_Newbie_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.ApplySetInitial_Newbie " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ApplySpellsStats(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo ApplySpellsStats_Err

    '</EhHeader>

    ' Aplicamos hechizos
    With UserList(UserIndex)

        If .Clase = eClass.Mage Or .Clase = eClass.Cleric Or .Clase = eClass.Druid Or .Clase = eClass.Bard Or .Clase = eClass.Assasin Or .Clase = eClass.Paladin Then
            
            .Stats.UserHechizos(1) = 2      ' Dardo Mágico
            .Stats.UserHechizos(2) = 1      ' Curar Veneno
            .Stats.UserHechizos(3) = 3      ' Curar heridas Leves

        End If

    End With

    '<EhFooter>
    Exit Sub

ApplySpellsStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.ApplySpellsStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub InitialUserStats(ByRef IUser As User)

    '<EhHeader>
    On Error GoTo InitialUserStats_Err

    '</EhHeader>
    
    ' Reset automático para incrementar de forma automática
    ' Adaptado a 0.11.5
    Dim LoopC As Integer

    Dim MiInt As Long

    Dim ln    As String
            
    With IUser
    
        .Stats.UserAtributos(eAtributos.Fuerza) = 18 + Balance.ModRaza(.Raza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = 18 + Balance.ModRaza(.Raza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = 18 + Balance.ModRaza(.Raza).Inteligencia
        .Stats.UserAtributos(eAtributos.Carisma) = 18 + Balance.ModRaza(.Raza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = 18 + Balance.ModRaza(.Raza).Constitucion
        
        ' Skills en 0
        For LoopC = 1 To NUMSKILLS
            .Stats.UserSkills(LoopC) = 0
            'Call CheckEluSkill(UserIndex, LoopC, True)
        Next LoopC

        ' Vida
        'MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
        .Stats.MaxHp = 20 ' + MiInt
        .Stats.MinHp = 20 ' + MiInt

        ' Energia
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)

        If MiInt = 1 Then MiInt = 2
        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt
          
        ' Agua y comida
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinHam = 100

        '<-----------------MANA----------------------->
        If .Clase = eClass.Mage Then  'Cambio en mana inicial (ToxicWaste)
            MiInt = 100
            .Stats.MaxMan = MiInt
            .Stats.MinMan = MiInt
        ElseIf .Clase = eClass.Cleric Or .Clase = eClass.Druid Or .Clase = eClass.Bard Or .Clase = eClass.Assasin Then
            .Stats.MaxMan = 50
            .Stats.MinMan = 50
        Else
            .Stats.MaxMan = 0
            .Stats.MinMan = 0

        End If

        .Stats.MinMan = .Stats.MaxMan
        
        .Stats.MaxHit = 2
        .Stats.MinHit = 1

        .Stats.Exp = 0
        .Stats.Elu = 300
        .Stats.Elv = 1
        .Stats.SkillPts = 10

    End With

    '<EhFooter>
    Exit Sub

InitialUserStats_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.InitialUserStats " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub UserLevelEditation(ByRef IUser As User, _
                              ByVal Elv As Byte, _
                              ByVal UserUps As Byte)
    ' Procedimiento creado para entrenamiento de personajes de nivel 1 a 15. (Editar con f1)
    
    Dim LoopC  As Integer

    Dim NewHp  As Integer

    Dim NewMan As Integer

    Dim NewSta As Integer

    Dim NewHit As Integer
    
    On Error GoTo UserLevelEditation_Error

    With IUser
        
        ' Quitamos los Itmes Newbies
        'Call QuitarNewbieObj(UserIndex)
        
        NewMan = .Stats.MaxMan
        NewHp = 0
        NewHit = 0
        NewSta = 0
        
        'Nivel 2 a 15
        For LoopC = 2 To Elv
            NewMan = NewMan + Balance_AumentoMANA(.Clase, .Raza, NewMan)
            NewSta = NewSta + Balance_AumentoSTA(.Clase)
            NewHit = NewHit + Balance_AumentoHIT(.Clase, LoopC)
        Next LoopC
        
        ' Nueva vida
        .Stats.MaxHp = getVidaIdeal(Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion)) + UserUps
        
        If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
        
        ' Nueva energía
        .Stats.MaxSta = .Stats.MaxSta + NewSta

        If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

        ' Nueva maná
        .Stats.MaxMan = NewMan

        If .Stats.MaxMan > STAT_MAXMAN Then .Stats.MaxMan = STAT_MAXMAN
        
        ' Nuevo golpe máximo y mínimo
        .Stats.MaxHit = .Stats.MaxHit + NewHit
        .Stats.MinHit = .Stats.MinHit + NewHit

        If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then .Stats.MaxHit = STAT_MAXHIT_UNDER36
        If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then .Stats.MinHit = STAT_MAXHIT_UNDER36
        
        .Stats.MinMan = .Stats.MaxMan
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinSta = .Stats.MaxSta
        .Stats.Elv = Elv
        .Stats.Elu = EluUser(.Stats.Elv)
        .Stats.Gld = 0
        .Stats.MaxAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinAGU = .Stats.MaxAGU
        .Stats.MinHam = .Stats.MinHam

    End With

    On Error GoTo 0

    Exit Sub

UserLevelEditation_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure UserLevelEditation of Módulo mBalance in line " & Erl

End Sub

Public Function Balance_AumentoSTA(ByVal UserClase As Byte) As Integer
    ' Aumento de energía
    
    On Error GoTo Balance_AumentoSTA_Error
            
    Select Case UserClase

        Case eClass.Thief
            Balance_AumentoSTA = AumentoSTLadron

        Case eClass.Mage
            Balance_AumentoSTA = AumentoSTMago

        Case Else
            Balance_AumentoSTA = AumentoSTDef

    End Select
        
    On Error GoTo 0

    Exit Function

Balance_AumentoSTA_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoSTA of Módulo mBalance in line " & Erl

End Function

Public Function Balance_AumentoHIT(ByVal UserClase As Byte, ByVal Elv As Byte) As Integer

    ' Aumento de HIT por nivel
    On Error GoTo Balance_AumentoHIT_Error
            
    Select Case UserClase

        Case eClass.Warrior, eClass.Hunter
            Balance_AumentoHIT = IIf(Elv > 35, 2, 3)
                    
        Case eClass.Paladin
            Balance_AumentoHIT = IIf(Elv > 35, 1, 3)
                    
        Case eClass.Thief
            Balance_AumentoHIT = 2
                         
        Case eClass.Mage
            Balance_AumentoHIT = 1
                   
        Case eClass.Cleric
            Balance_AumentoHIT = 2
                    
        Case eClass.Druid
            Balance_AumentoHIT = 2
                     
        Case eClass.Assasin
            Balance_AumentoHIT = IIf(Elv > 35, 1, 3)

        Case eClass.Bard
            Balance_AumentoHIT = 2
                    
        Case Else
            Balance_AumentoHIT = 2

    End Select

    On Error GoTo 0

    Exit Function

Balance_AumentoHIT_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoHIT of Módulo mBalance in line " & Erl

End Function

Public Function Balance_AumentoMANA(ByVal Class As Byte, ByVal Raze As Byte, ByRef TempMan As Integer) As Integer

    ' Aumento de maná según clase
    '<EhHeader>
    On Error GoTo Balance_AumentoMANA_Err

    '</EhHeader>
    
    Dim UserInteligencia As Byte

    UserInteligencia = 18 + Balance.ModRaza(Raze).Inteligencia
    
    On Error GoTo Balance_AumentoMANA_Error

    Select Case Class
                    
        Case eClass.Paladin
            Balance_AumentoMANA = UserInteligencia
                         
        Case eClass.Mage

            If Raze = Enano Then
                Balance_AumentoMANA = 2 * UserInteligencia
            ElseIf (TempMan >= 2000) Then
                Balance_AumentoMANA = (3 * UserInteligencia) / 2
            Else
                Balance_AumentoMANA = 3 * UserInteligencia

            End If
                   
        Case eClass.Druid, eClass.Bard, eClass.Cleric
            Balance_AumentoMANA = 2 * UserInteligencia

        Case eClass.Assasin
            Balance_AumentoMANA = UserInteligencia
                    
        Case Else
            Balance_AumentoMANA = 0
            
    End Select
        
    On Error GoTo Balance_AumentoMANA_Err

    Exit Function

Balance_AumentoMANA_Error:

    LogError "Error " & Err.number & " (" & Err.description & ") in procedure Balance_AumentoMANA of Módulo mBalance in line " & Erl

    '<EhFooter>
    Exit Function

Balance_AumentoMANA_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.Balance_AumentoMANA " & "at line " & Erl
        
    '</EhFooter>
End Function

' Retorna la vida ideal que deberia tener el personaje para su nivel
Public Function getVidaIdeal(ByVal Elv As Byte, ByVal Class As Byte, ByVal Constitucion As Byte) As Single

    '<EhHeader>
    On Error GoTo getVidaIdeal_Err

    '</EhHeader>

    Dim promedio     As Single

    Dim vidaBase     As Integer

    Dim rangoAumento As tRango
    
    vidaBase = 20
    
    rangoAumento = getRangoAumentoVida(Class, Constitucion)
    promedio = ((rangoAumento.minimo + rangoAumento.maximo) / 2)
    
    getVidaIdeal = ((vidaBase + (Elv - 1) * promedio))

    '<EhFooter>
    Exit Function

getVidaIdeal_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.getVidaIdeal " & "at line " & Erl
        
    '</EhFooter>
End Function

' El personaje aumenta su vida
Public Function obtenerAumentoHp(ByVal UserIndex As Integer) As Byte

    '<EhHeader>
    On Error GoTo obtenerAumentoHp_Err

    '</EhHeader>

    ' Calculo de vida
    Dim vidaPromedio  As Integer

    Dim promedio      As Single

    Dim vidaIdeal     As Single
        
    Dim minimoAumento As Integer

    Dim maximoAumento As Integer

    Dim aumentoHp     As Byte
    
    Dim rangoAumento  As tRango

    Dim vidaBase      As Integer
    
    Dim Random        As Integer
    
    vidaBase = 20

    rangoAumento = getRangoAumentoVida(UserList(UserIndex).Clase, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    promedio = (rangoAumento.minimo + rangoAumento.maximo) / 2
    vidaIdeal = vidaBase + (UserList(UserIndex).Stats.Elv - 2) * promedio
        
    Dim puntosAumento As Integer
        
    puntosAumento = Int(RandomNumber(rangoAumento.minimo, rangoAumento.maximo))
                
    If UserList(UserIndex).Stats.MaxHp < vidaIdeal + 1.5 Then
        aumentoHp = maxi(puntosAumento, Int(0.5 + promedio))
    Else
        aumentoHp = puntosAumento
            
        If rangoAumento.minimo = puntosAumento Then
            If RandomNumber(1, 100) <= 50 Then
                aumentoHp = puntosAumento + 1

            End If

        End If

    End If
    
    obtenerAumentoHp = aumentoHp

    '<EhFooter>
    Exit Function

obtenerAumentoHp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.obtenerAumentoHp " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub RecompensaPorNivel(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo RecompensaPorNivel_Err

    '</EhHeader>

    Dim Obj        As Obj

    Dim aumentoHp  As Single

    Dim AumentoMan As Integer

    Dim AumentoHit As Integer

    Dim AumentoSta As Integer
    
    Dim Texto      As String
    
    Dim Ups        As Single
    
    With UserList(UserIndex)

        If .Stats.Elv = STAT_MAXELV Then
            .Stats.Exp = 0
            .Stats.Elu = 0

        End If
        
        Texto = "Nivel '" & .Stats.Elv & "' "

        AumentoSta = Balance_AumentoSTA(.Clase)
        aumentoHp = obtenerAumentoHp(UserIndex)
        AumentoMan = Balance_AumentoMANA(.Clase, .Raza, .Stats.MaxMan)
        AumentoHit = Balance_AumentoHIT(.Clase, .Stats.Elv)
        
        .Stats.MaxHp = .Stats.MaxHp + aumentoHp
        .Stats.MaxMan = .Stats.MaxMan + AumentoMan
        .Stats.MinHit = .Stats.MinHit + AumentoHit
        .Stats.MaxHit = .Stats.MaxHit + AumentoHit
        .Stats.MaxSta = .Stats.MaxSta + AumentoSta
        
        If .Stats.Elv < 36 Then
            If .Stats.MinHit > STAT_MAXHIT_UNDER36 Then .Stats.MinHit = STAT_MAXHIT_UNDER36
        Else

            If .Stats.MinHit > STAT_MAXHIT_OVER36 Then .Stats.MinHit = STAT_MAXHIT_OVER36

        End If
        
        If .Stats.Elv < 36 Then
            If .Stats.MaxHit > STAT_MAXHIT_UNDER36 Then .Stats.MaxHit = STAT_MAXHIT_UNDER36
        Else

            If .Stats.MaxHit > STAT_MAXHIT_OVER36 Then .Stats.MaxHit = STAT_MAXHIT_OVER36

        End If

        Ups = .Stats.MaxHp - Mod_Balance.getVidaIdeal(.Stats.Elv, .Clase, .Stats.UserAtributos(eAtributos.Constitucion))
            
        Texto = Texto & "Vida +'" & aumentoHp & "' puntos de vida. Ups: " & IIf((Ups = 0), "Ninguno", Ups) & "."

        If AumentoMan Then Texto = Texto & " Maná +" & AumentoMan
        
        If AumentoHit Then
            Texto = Texto & " Golpe +" & AumentoHit & ". "

        End If
        
        Call WriteConsoleMsg(UserIndex, Texto, FontTypeNames.FONTTYPE_INFO)
        
        Call Logs_User(.Name, eLog.eUser, eLvl, "paso a nivel " & .Stats.Elv & " gano HP: " & aumentoHp)
          
        Call WriteUpdateUserStats(UserIndex)
    
    End With

    '<EhFooter>
    Exit Sub

RecompensaPorNivel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.RecompensaPorNivel " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Reset_DesquiparAll(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Reset_DesquiparAll_Err

    '</EhHeader>

    With UserList(UserIndex)

        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

        End If
        
        ' Desequipamos la montura
        If .Invent.MonturaObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MonturaSlot)

        End If
        
        ' Desequipamos el pendiente
        If .Invent.PendientePartyObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.PendientePartySlot)

        End If
            
        ' Desequipamos la reliquia
        If .Invent.ReliquiaObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.ReliquiaSlot)

        End If
        
        ' Desequipamos el Objeto mágico (Laudes y Anillos mágicos)
        If .Invent.MagicObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)

        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

        End If
        
        'desequipar aura
        If .Invent.AuraEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.AuraEqpSlot)

        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

        End If
        
        'desequipar anillo magico/laud
        If .Invent.MagicSlot > 0 Then
            Call Desequipar(UserIndex, .Invent.MagicSlot)

        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

        End If
    
    End With

    '<EhFooter>
    Exit Sub

Reset_DesquiparAll_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.Reset_DesquiparAll " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Obtiene el Promedio de Aumento de vida del Personaje
Public Function getPromedioAumentoVida(ByVal Class As Byte, ByVal Constitucion As Byte) As Single

    '<EhHeader>
    On Error GoTo getPromedioAumentoVida_Err

    '</EhHeader>

    Dim Rango As tRango
    
    Rango = getRangoAumentoVida(Class, Constitucion)
    
    getPromedioAumentoVida = (Rango.maximo + Rango.minimo) / 2

    '<EhFooter>
    Exit Function

getPromedioAumentoVida_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.getPromedioAumentoVida " & "at line " & Erl
        
    '</EhFooter>
End Function

' Retrona el minimo/maximo de puntos de vida que pude subir este usuario por nivel.
Public Function getRangoAumentoVida(ByVal Class As Byte, ByVal Constitucion As Byte) As tRango

    '<EhHeader>
    On Error GoTo getRangoAumentoVida_Err

    '</EhHeader>

    getRangoAumentoVida.maximo = 0
    getRangoAumentoVida.minimo = 0

    Select Case Class

        Case eClass.Warrior

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 12

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 12

                Case 19
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 6 + AdicionalHPCazador
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Hunter
    
            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 11

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 19
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Paladin

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 9
                    getRangoAumentoVida.maximo = 11

                Case 20
                    getRangoAumentoVida.minimo = 8
                    getRangoAumentoVida.maximo = 11

                Case 19
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 11

                Case 18
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 11

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 + AdicionalHPCazador

            End Select

        Case eClass.Thief

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case 16, 17
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 7

                Case 16
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 6

                Case 14
                    getRangoAumentoVida.minimo = 2
                    getRangoAumentoVida.maximo = 6

                Case 13
                    getRangoAumentoVida.minimo = 2
                    getRangoAumentoVida.maximo = 5

                Case 12
                    getRangoAumentoVida.minimo = 1
                    getRangoAumentoVida.maximo = 5

                Case 11
                    getRangoAumentoVida.minimo = 1
                    getRangoAumentoVida.maximo = 4

                Case 10
                    getRangoAumentoVida.minimo = 0
                    getRangoAumentoVida.maximo = 4

                Case Else
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

            End Select
    
        Case eClass.Mage

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 8

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 8

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case 18
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = 8

                Case Else
                    getRangoAumentoVida.minimo = 3
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPGuerrero

            End Select
                
        Case eClass.Cleric

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case eClass.Druid

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select
        
        Case eClass.Assasin

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case eClass.Bard

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 7
                    getRangoAumentoVida.maximo = 10

                Case 20
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 10

                Case 19
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 18
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case Else
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

        Case Else

            Select Case Constitucion

                Case 21
                    getRangoAumentoVida.minimo = 6
                    getRangoAumentoVida.maximo = 9

                Case 20
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = 9

                Case 19
                    getRangoAumentoVida.minimo = 4
                    getRangoAumentoVida.maximo = 8

                Case Else
                    getRangoAumentoVida.minimo = 5
                    getRangoAumentoVida.maximo = Constitucion \ 2 - AdicionalHPCazador

            End Select

    End Select

    '<EhFooter>
    Exit Function

getRangoAumentoVida_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.Mod_Balance.getRangoAumentoVida " & "at line " & Erl
        
    '</EhFooter>
End Function

