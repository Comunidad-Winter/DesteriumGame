Attribute VB_Name = "mNewChars"
Option Explicit

Public Sub IUser_Editation_Skills(ByRef Temp As User)

    '<EhHeader>
    On Error GoTo IUser_Editation_Skills_Err

    '</EhHeader>
    
    ' Todas
    Temp.Stats.UserSkills(eSkill.Resistencia) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Tacticas) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Armas) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Comerciar) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Navegacion) = 35
    Temp.Stats.UserSkills(eSkill.Ocultarse) = RandomNumber(1, 37)
           
    ' Clases con maná
    If Temp.Stats.MaxMan > 0 Then
        Temp.Stats.UserSkills(eSkill.Magia) = LevelSkill(Temp.Stats.Elv).LevelValue ' + RandomNumber(1, 7)

    End If
    
    ' Todas menos mago APUÑALAN
    If Temp.Clase <> eClass.Mage Then
        Temp.Stats.UserSkills(eSkill.Apuñalar) = LevelSkill(Temp.Stats.Elv).LevelValue

    End If
    
    ' Escudos
    If Temp.Clase <> eClass.Mage And Temp.Clase <> eClass.Druid Then
        Temp.Stats.UserSkills(eSkill.Defensa) = LevelSkill(Temp.Stats.Elv).LevelValue

    End If
    
    If Temp.Clase = eClass.Hunter Or Temp.Clase = eClass.Warrior Then
        Temp.Stats.UserSkills(eSkill.Proyectiles) = LevelSkill(Temp.Stats.Elv).LevelValue

    End If
        
    ' Armas de proyectiles
    If Temp.Clase = eClass.Hunter Then
        Temp.Stats.UserSkills(eSkill.Ocultarse) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(1, 14)

    End If
    
    ' Ocultarse & Robar
    If Temp.Clase = eClass.Thief Then
        Temp.Stats.UserSkills(eSkill.Ocultarse) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(20, 40)
        Temp.Stats.UserSkills(eSkill.Robar) = LevelSkill(Temp.Stats.Elv).LevelValue + RandomNumber(20, 40)

    End If
    
    ' Mineria, Tala, Pesca
    Temp.Stats.UserSkills(eSkill.Pesca) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Mineria) = LevelSkill(Temp.Stats.Elv).LevelValue
    Temp.Stats.UserSkills(eSkill.Talar) = LevelSkill(Temp.Stats.Elv).LevelValue
        
    ' Checking Final
    Dim A As Long
        
    For A = 1 To NUMSKILLS
        
        If Temp.Stats.UserSkills(A) > 100 Then
            Temp.Stats.UserSkills(A) = 100

        End If
            
    Next A
    
    '<EhFooter>
    Exit Sub

IUser_Editation_Skills_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.FrmPanelCreator.IUser_Editation_Skills " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub IUser_Editation_Reputacion_Frags(ByRef Temp As User, ByVal Frags As Integer)

    '<EhHeader>
    On Error GoTo IUser_Editation_Reputacion_Frags_Err

    '</EhHeader>
    Dim L As Long

    Frags = RandomNumber(val(FrmPanelCreator.txtFrags(0).Text), val(FrmPanelCreator.txtFrags(1).Text))
    
    With Temp.Reputacion

        If RandomNumber(1, 100) >= 50 Then
            .AsesinoRep = (vlASESINO * 2)

            If .AsesinoRep > MAXREP Then .AsesinoRep = MAXREP
             
            .BandidoRep = Frags + RandomNumber(1, 12) * 1000
             
            .BurguesRep = 0
            .NobleRep = 0
            .PlebeRep = 0
            Temp.Faction.FragsCiu = Frags
            Temp.Faction.FragsCri = RandomNumber(1, Frags)
            Temp.Faction.FragsOther = Temp.Faction.FragsCri + Temp.Faction.FragsCiu
            
        Else
            .NobleRep = (vlNoble * Frags)

            If .NobleRep > MAXREP Then .NobleRep = MAXREP
            Temp.Faction.FragsCri = Frags
            Temp.Faction.FragsOther = Temp.Faction.FragsCri

        End If
        
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = L / 6
        .promedio = L

    End With

    '<EhFooter>
    Exit Sub

IUser_Editation_Reputacion_Frags_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.FrmPanelCreator.IUser_Editation_Reputacion_Frags " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub IUser_Editation_Spells(ByRef Temp As User)

    Dim Slot As Byte
    
    With Temp
    
        Slot = 35
        .Stats.UserHechizos(Slot) = 10  ' Remover Parálisis
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 24  ' Inmovilizar
    
        If .Stats.MaxMan >= 1000 Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 25  ' Apocalipsis

        End If
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 23  ' Descarga eléctrica
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 15  ' Tormenta de Fuego
    
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 18  ' Celeridad
    
        If .Clase <> eClass.Mage Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 20  ' Fuerza

        End If
        
        Slot = Slot - 1
        .Stats.UserHechizos(Slot) = 14  ' Invisibilidad
    
        If .Stats.MaxMan >= 1000 Then
            Slot = Slot - 1
            .Stats.UserHechizos(Slot) = 11  ' Resucitar

        End If
    
        Slot = Slot - 1
     
        If .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
            .Stats.UserHechizos(Slot) = 53  ' Invocar Mascotas
        Else
            .Stats.UserHechizos(Slot) = 52  ' Invocar Mascotas

        End If
    
        ' Hechizos básicos fuera de onda
        Slot = 1
        .Stats.UserHechizos(Slot) = 2  ' Curar Veneno
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 8  ' Misil Mágico
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 17  ' Invocar Zombies
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 3  ' Curar heridas leves
        Slot = Slot + 1
        .Stats.UserHechizos(Slot) = 5  ' Curar heridas Graves
        
        If .Clase = eClass.Mage Then
            Slot = Slot + 1
            .Stats.UserHechizos(Slot) = 20  ' Fuerza

        End If
    
    End With

End Sub

