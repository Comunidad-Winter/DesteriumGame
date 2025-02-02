VERSION 5.00
Begin VB.Form FrmObject_Info 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Informacion del Objeto"
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmObject_Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   194
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tUpdate 
      Interval        =   40
      Left            =   240
      Top             =   0
   End
   Begin VB.Image imgUnload 
      Height          =   435
      Left            =   2400
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "FrmObject_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PIXEL_HEIGHT  As Integer = 240 '220

Private Const TWIP_TO_PIXEL As Integer = 1

Public pixelHeight          As Long

Private Height_Original     As Integer

Private Width_Original      As Integer

Public FormMovement         As clsFormMovementManager

Private Type tLine

    Text As String
    Color As Long
    Font As Integer
    Size As Integer

End Type

Public LastLine  As Integer

Private lines()  As tLine

Private Loading  As Boolean

Private Armadura As Integer, Arma As Integer, Casco As Integer, Escudo As Integer

Private Heading  As E_Heading

Private Sub Add_Line(ByVal Text As String, _
                     ByVal Color As Long, _
                     Font As Byte, _
                     Size As Byte)

    '<EhHeader>
    On Error GoTo Add_Line_Err

    '</EhHeader>
    
    LastLine = LastLine + 1
    ReDim Preserve lines(0 To LastLine) As tLine
    
    With lines(LastLine)
        .Text = Text
        .Color = Color
        .Font = Font
        .Size = Size

    End With
    
    pixelHeight = pixelHeight + (PIXEL_HEIGHT * TWIP_TO_PIXEL)
    '<EhFooter>
    Exit Sub

Add_Line_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Add_Line " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' # Dibuja como se ve el objeto equipado
Private Sub Object_RenderEquiped(ByVal ObjIndex As Integer)

    Dim X As Integer

    Dim Y As Integer
    
    X = 130
    Y = 80

    Dim Depth As Single

    Dim Mult  As Byte

    Dim Div   As Byte
    
    Depth = 2
    Heading = CharList(UserCharIndex).Heading

    With CharList(UserCharIndex)

        If .iBody > 0 And Not .Invisible Then
            
            If .iHead > 0 And ObjData(ObjIndex).ObjType <> otTransformVIP Then
                Call Draw_Grh_Menu(HeadData(.iHead).Head(Heading), X + (BodyData(.iBody).HeadOffset.X), Y + (BodyData(.iBody).HeadOffset.Y), To_Depth(Depth + 3), 1, 0, , , , eTechnique.t_Alpha, , , , , True)

            End If
            
            If Armadura > 0 Then
                Call Draw_Grh_Menu(BodyData(Armadura).Walk(Heading), X, Y, To_Depth(Depth + 2), 1, 0, 0, , , eTechnique.t_Alpha, , , , , True)
            Else
                Call Draw_Grh_Menu(BodyData(.iBody).Walk(Heading), X, Y, To_Depth(Depth + 2), 1, 0, 0, , , eTechnique.t_Alpha, , , , , True)

            End If
            
            If Arma > 0 Then
                Call Draw_Grh_Menu(WeaponAnimData(Arma).WeaponWalk(Heading), X, Y, To_Depth(Depth + 3), 1, 1, 0, , , eTechnique.t_Alpha, , , , , True)
            Else

                If .Arma.WeaponWalk(Heading).GrhIndex > 0 Then Call Draw_Grh_Menu(.Arma.WeaponWalk(Heading), X, Y, To_Depth(Depth + 3), 1, 1, 0, , , eTechnique.t_Alpha, , , , , True)

            End If
            
            If Casco > 0 Then
                Call Draw_Grh_Menu(CascoAnimData(Casco).Head(Heading), X + BodyData(.iBody).HeadOffset.X, Y + (BodyData(.iBody).HeadOffset.Y), To_Depth(Depth + 3, , , 3), 1, 0, , , , eTechnique.t_Alpha, , , , True)
            Else

                If .Casco.Head(Heading).GrhIndex > 0 Then
                    Call Draw_Grh_Menu(.Casco.Head(Heading), X + BodyData(.iBody).HeadOffset.X, Y + ((BodyData(.iBody).HeadOffset.Y)), To_Depth(Depth + 3, , , 3), 1, 0, , , , eTechnique.t_Alpha, , , , True)

                End If

            End If
            
            If Escudo > 0 Then
                Call Draw_Grh_Menu(ShieldAnimData(Escudo).ShieldWalk(Heading), X, Y, To_Depth(Depth + 3, , , 4), 1, 1, 0, , , eTechnique.t_Alpha, , , , , True)
            Else

                If .Escudo.ShieldWalk(Heading).GrhIndex > 0 Then Call Draw_Grh_Menu(.Escudo.ShieldWalk(Heading), X, Y, To_Depth(Depth + 3, , , 4), 1, 1, 0, , , eTechnique.t_Alpha, , , , , True)

            End If
                
        End If
    
    End With

End Sub

Private Sub Prepare_Npcs()

    '<EhHeader>
    On Error GoTo Prepare_Npcs_Err

    '</EhHeader>
    
    LastLine = 0
    pixelHeight = 0
    ReDim Preserve lines(LastLine) As tLine
     
    Dim Npc As tNpcs
    
    Npc = NpcList(SelectedNpcIndex)

    ' Nombre de la criatura
    Call Add_Line(Npc.Name, ARGB(255, 255, 255, 255), eFonts.f_Medieval, 15)

    ' Comerciantes & Npcs interactivos
    If 0 = 0 Then
        
    Else
        ' Criaturas que atacan

        If Npc.MinHit > 0 Then
            Call Add_Line("Hit: " & Npc.MinHit & "/" & Npc.MaxHit, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

        End If
        
        If Npc.Def > 0 Then
            Call Add_Line("Def: " & Npc.Def, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

        End If
        
        If Npc.DefM > 0 Then
            Call Add_Line("RM: " & Npc.DefM, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

        End If
        
    End If

    Me.Height = Me.Height + pixelHeight
    MirandoObjetos = True

    '<EhFooter>
    Exit Sub

Prepare_Npcs_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Prepare_Npcs " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Prepare_Object()

    '<EhHeader>
    On Error GoTo Prepare_Object_Err

    '</EhHeader>
        
    '  Me.visible = False
    LastLine = 0
    pixelHeight = 0
    ReDim Preserve lines(LastLine) As tLine
     
    Dim Obj As tObjData

    Dim A   As Long
        
    If SelectedObjIndex = 0 Then Exit Sub
    Obj = ObjData(SelectedObjIndex)
    
    ' Nombre del Objeto
    Call Add_Line(Obj.Name, ARGB(247, 232, 80, 255), eFonts.f_Tahoma, 16)
        
    ' Hit del Objeto
    If Obj.MinHit > 0 Then
        Call Add_Line("Hit: " & Obj.MinHit & "/" & Obj.MaxHit, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If
        
    ' Def del Objeto
    If Obj.MinDef > 0 Then
        Call Add_Line("Def: " & Obj.MinDef & "/" & Obj.MaxDef, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If
    
    ' Def Mag del Objeto
    If Obj.MinDefRM > 0 Then
        Call Add_Line("RM: " & Obj.MinDefRM & "/" & Obj.MaxDefRM, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If
        
    ' Proyectiles
    If Obj.Proyectil > 0 Then
        Call Add_Line("Necesita municiones", ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If
    
    If Obj.LvlMin > 0 Then
        Call Add_Line("Nivel M�nimo: " & Obj.LvlMin, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If
    
    If Obj.LvlMax > 0 Then
        Call Add_Line("Nivel M�ximo: " & Obj.LvlMax, ARGB(255, 255, 255, 255), eFonts.f_Verdana, 14)

    End If

    If Obj.Skin > 0 Then
        If Obj.ValueGLD > 0 Then
            Call Add_Line("Oro: " & PonerPuntos(Obj.ValueGLD), ARGB(255, 228, 0, 255), eFonts.f_Verdana, 14)

        End If
        
        If Obj.ValueDSP > 0 Then
            Call Add_Line("Dsp: " & PonerPuntos(Obj.ValueDSP), ARGB(255, 175, 0, 255), eFonts.f_Verdana, 14)

        End If

    End If
        
    If Obj.Points > 0 Then
        Call Add_Line("Puntos DS (x1): " & PonerPuntos(CalculateSellPrice(CSng(Obj.Points), 1)), ARGB(255, 228, 0, 255), eFonts.f_Verdana, 14)
        
    End If

    If Obj.ObjType = otTeleportInvoker Then
        If Obj.TimeWarp > 0 Then
            Call Add_Line("Aparece en " & Obj.TimeWarp & " segundos", ARGB(180, 244, 190, 255), eFonts.f_Verdana, 14)

        End If
        
        If Obj.TimeDuration > 0 Then
            Call Add_Line("Visible durante " & Obj.TimeDuration & " segundos", ARGB(248, 190, 155, 255), eFonts.f_Verdana, 14)

        End If
        
        If Obj.PuedeInsegura = 0 Then
            Call Add_Line("Invocable Zona Segura", ARGB(180, 244, 82, 255), eFonts.f_Verdana, 14)
        Else
            Call Add_Line("Invocable Zona Insegura", ARGB(190, 70, 50, 255), eFonts.f_Verdana, 14)

        End If
    
    End If
    
    If Obj.RemoveObj > 0 Then
        Call Add_Line("OBJETO USABLE", ARGB(214, 10, 10, 255), eFonts.f_Verdana, 14)

    End If
        
    ' Atributos del Objeto
    If Obj.SkillNum > 0 Then

        For A = 1 To Obj.SkillNum
            Call Add_Line("+" & Obj.Skill(A).Amount & " " & InfoSkill(Obj.Skill(A).Selected).Name, InfoSkill(Obj.Skill(A).Selected).Color, eFonts.f_Verdana, 14)
        Next A

    End If

    ' Atributos Extremos del Objeto
    If Obj.SkillsEspecialNum > 0 Then

        For A = 1 To Obj.SkillsEspecialNum
            Call Add_Line("+" & Obj.SkillsEspecial(A).Amount & " " & InfoSkillEspecial(Obj.SkillsEspecial(A).Selected).Name, InfoSkillEspecial(Obj.SkillsEspecial(A).Selected).Color, eFonts.f_Verdana, 14)
        Next A

    End If
        
    ' Objetos que requiere para comprarlo. Asi el npc puede fabricarlo.
    If Obj.Upgrade.RequiredCant > 0 Then
        Call Add_Line("Requiere:", ARGB(228, 20, 10, 255), eFonts.f_Tahoma, 16)
            
        For A = 1 To Obj.Upgrade.RequiredCant
            Call Add_Line(ObjData(Obj.Upgrade.Required(A).ObjIndex).Name & "(" & Obj.Upgrade.Required(A).Amount & ")", ARGB(240, 160, 160, 255), eFonts.f_Verdana, 14)
        Next A

    End If
        
    If Obj.GuildLvl > 0 Then
        Call Add_Line("Requiere Clan Lvl " & Obj.GuildLvl, ARGB(214, 10, 10, 255), eFonts.f_Verdana, 14)

    End If
        
    If Obj.Skin = 0 Then
        If Obj.NoSeCae > 0 Or Obj.ObjType = otBarcos Then
            Call Add_Line("NO SE CAE AL MORIR", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Else
            Call Add_Line("SE CAE AL MORIR", ARGB(214, 10, 10, 255), eFonts.f_Verdana, 14)

        End If

    End If
        
    ' # Render object equiped
    If Obj.ObjType = otTransformVIP Then
        Armadura = Obj.Anim
    Else
        Armadura = IIf(Obj.ObjType = otarmadura, Obj.Anim, 0)
        Arma = IIf(Obj.ObjType = otWeapon, Obj.Anim, 0)
        Casco = IIf(Obj.ObjType = otcasco, Obj.Anim, 0)
        Escudo = IIf(Obj.ObjType = otescudo, Obj.Anim, 0)

    End If
        
    If Arma > 0 Or Armadura > 0 Or Casco > 0 Or Escudo > 0 Then
        ' Linea en blanco
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Add_Line(" ", ARGB(50, 180, 8, 255), eFonts.f_Verdana, 14)
        Call Object_RenderEquiped(SelectedObjIndex)

    End If
        
    Me.Height = Me.Height + pixelHeight
    MirandoObjetos = True

    '<EhFooter>
    Exit Sub

Prepare_Object_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Prepare_Object " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Initial_Form()

    '<EhHeader>
    On Error GoTo Initial_Form_Err

    '</EhHeader>
        
    MirandoObjetos = True
    Me.visible = False
    Width_Original = 2900
    Height_Original = 490
    
    Me.Width = Width_Original
    Me.Height = Height_Original
    
    Call Prepare_Object
        
    If g_Captions(eCaption.cObjectInfo) > 0 Then
        Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.cObjectInfo))

    End If
        
    g_Captions(eCaption.cObjectInfo) = wGL_Graphic.Create_Device_From_Display(FrmObject_Info.hWnd, FrmObject_Info.ScaleWidth, FrmObject_Info.ScaleHeight)
    Render_Obj

    Me.visible = True
    Exit Sub

Initial_Form_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Initial_Form " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Close_Form()

    '<EhHeader>
    On Error GoTo Close_Form_Err

    '</EhHeader>

    Unload Me
    '<EhFooter>
    Exit Sub

Close_Form_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Close_Form " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Load()

    Set FormMovement = New clsFormMovementManager
    
    Call FormMovement.Initialize(Me, 32)
    
    Initial_Form
    
End Sub

Public Sub Render_Obj()

    '<EhHeader>
    On Error GoTo Render_Obj_Err

    '</EhHeader>
    
    Dim A        As Long

    Dim Y_Avance As Long
    
    Dim Color    As Long

    Dim Tier     As Byte
    
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.cObjectInfo))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, FrmObject_Info.ScaleWidth, FrmObject_Info.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    ' Color del fondo
    
    Tier = 0
    
    Select Case Tier

        Case 2
            Color = ARGB(50, 255, 0, 255)
    
        Case 3
            Color = ARGB(0, 240, 255, 255)
            
        Case 4
            Color = ARGB(255, 0, 240, 255)
            
        Case 5
            Color = ARGB(255, 255, 0, 255)
        
        Case Else
            Color = ARGB(255, 255, 255, 255)

    End Select
    
    ' Borde Superior
    Call Draw_Texture_Graphic_Gui(129, 0, 0, To_Depth(1), 193, 16, 0, 0, 193, 16, Color, 0, eTechnique.t_Default)
    Y_Avance = 16
      
    For A = 1 To UBound(lines)

        With lines(A)
            Call Draw_Text(.Font, .Size, 15, Y_Avance + 1, To_Depth(3), 0, .Color, FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .Text, False, True)
            Call Draw_Texture_Graphic_Gui(130, 0, Y_Avance, To_Depth(1), 193, 16, 0, 0, 193, 16, Color, 0, eTechnique.t_Default)
            Y_Avance = Y_Avance + 16
            
            ' Call Draw_Texture_Graphic_Gui(130, 0, Y_Avance, To_Depth(1), 193, 16, 0, 0, 193, 16, ARGB(255, 255, 255, 255), 0, eTechnique.t_Default)
            
        End With

    Next A

    If Arma > 0 Or Armadura > 0 Or Casco > 0 Or Escudo > 0 Then
        Call Object_RenderEquiped(SelectedObjIndex)

    End If

    Call Draw_Texture_Graphic_Gui(131, 0, Y_Avance, To_Depth(2), 193, 16, 0, 0, 193, 16, Color, 0, eTechnique.t_Default)
    
    Call wGL_Graphic_Renderer.Flush

    '<EhFooter>
    Exit Sub

Render_Obj_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObject_Info.Render_Obj " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoObjetos = False
    
    ' If FrmMain.visible Then
    'FrmMain.SetFocus
    ' End If
    
End Sub

Private Sub imgUnload_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    Unload Me

End Sub

