VERSION 5.00
Begin VB.Form FrmObjetive 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Objetivos Argentum"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmObjetive.frx":0000
   LinkTopic       =   "Objetivos Argentum"
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picQuest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1710
      Left            =   195
      MousePointer    =   99  'Custom
      ScaleHeight     =   114
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   825
      Width           =   4815
   End
   Begin VB.Timer tUpdate 
      Interval        =   40
      Left            =   4560
      Top             =   360
   End
   Begin VB.Image imgWeb 
      Height          =   495
      Left            =   1440
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image ImgNext 
      Height          =   375
      Left            =   4560
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgLast 
      Height          =   375
      Left            =   240
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgMision 
      Height          =   210
      Index           =   2
      Left            =   720
      Top             =   2220
      Width           =   2775
   End
   Begin VB.Image imgMision 
      Height          =   210
      Index           =   1
      Left            =   720
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Image imgMision 
      Height          =   255
      Index           =   0
      Left            =   720
      Top             =   1740
      Width           =   2295
   End
   Begin VB.Image imgReward 
      Height          =   495
      Left            =   1560
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Image imgItem 
      Height          =   615
      Index           =   3
      Left            =   2880
      Top             =   5880
      Width           =   600
   End
   Begin VB.Image imgItem 
      Height          =   615
      Index           =   2
      Left            =   2130
      Top             =   5880
      Width           =   600
   End
   Begin VB.Image imgItem 
      Height          =   615
      Index           =   1
      Left            =   1380
      Top             =   5880
      Width           =   600
   End
   Begin VB.Image imgItem 
      Height          =   615
      Index           =   0
      Left            =   600
      Top             =   5880
      Width           =   600
   End
   Begin VB.Image imgSelected 
      Height          =   375
      Index           =   1
      Left            =   2760
      Top             =   960
      Width           =   1815
   End
   Begin VB.Image imgSelected 
      Height          =   375
      Index           =   0
      Left            =   720
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image imgUnload 
      Height          =   315
      Left            =   4800
      Picture         =   "FrmObjetive.frx":000C
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "FrmObjetive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ListQuests As clsGraphicalList

Public ViewDesc   As Boolean

Public Enum eBotonSelected

    eNormal = 0
    eAltoRiesgo = 1

End Enum

Public BotonSelected  As eBotonSelected

Private clsFormulario As clsFormMovementManager

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me

    End If

End Sub

Private Sub Form_Load()

    #If ModoBig = 0 Then
        ' Handles Form movement (drag and dr|op).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me
    #End If
    
    g_Captions(eCaption.e_Objetivos) = wGL_Graphic.Create_Device_From_Display(Me.hWnd, Me.ScaleWidth, Me.ScaleHeight)
    
    Dim A As Long
    
    If NpcsUser_QuestIndex > 0 Then
        If QuestList(NpcsUser_QuestIndex).RewardObj > 0 Then

            For A = imgItem.LBound To imgItem.UBound

                If (A + 1) < QuestList(NpcsUser_QuestIndex).RewardObj Then
                    imgItem(A).Enabled = True
                    imgItem(A).ToolTipText = ObjData(QuestList(NpcsUser_QuestIndex).RewardObjs(A + 1).ObjIndex).Name & " (x" & QuestList(NpcsUser_QuestIndex).RewardObjs(A + 1).Amount & ")"
                Else
                    imgItem(A).Enabled = False

                End If

            Next A
        
        Else
            imgItem(A).Enabled = False

        End If
        
    Else
        BotonSelected = eAltoRiesgo

    End If
    
    Render_Objetive

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ViewDesc = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call wGL_Graphic.Destroy_Device(g_Captions(eCaption.e_Objetivos))
  
    Set ListQuests = Nothing
  
End Sub

Private Sub imgItem_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)

End Sub

Private Sub imgLast_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If NpcsUser_QuestIndex = 0 Then Exit Sub
    
    If QuestList(NpcsUser_QuestIndex).LastQuest = 0 Then
        Call ShowConsoleMsg("¡No hay nada para atrás muchacho!")
        Exit Sub

    End If
    
    NpcsUser_QuestIndex = QuestList(NpcsUser_QuestIndex).LastQuest
    Call Quests_CheckViewObjs(NpcsUser_QuestIndex)

End Sub

Private Sub imgMision_Click(Index As Integer)
    Call Audio.PlayInterface(SND_CLICK)

End Sub

Private Sub imgMision_MouseMove(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    ' ViewDesc = True
End Sub

Private Sub ImgNext_Click()
    Call Audio.PlayInterface(SND_CLICK)
    
    If NpcsUser_QuestIndex = 0 Then Exit Sub
    
    If QuestList(NpcsUser_QuestIndex).NextQuest = 0 Then
        Call ShowConsoleMsg("¡Ya has terminado todas los objetivos principales del juego! ¡Eres un Record!")
        Exit Sub

    End If
    
    NpcsUser_QuestIndex = QuestList(NpcsUser_QuestIndex).NextQuest
    Call Quests_CheckViewObjs(NpcsUser_QuestIndex)

End Sub

Private Sub imgReward_Click()

    If Not PuedeReclamar Then Exit Sub
    Call Audio.PlayInterface(SND_CLICK)
    
    'Call WriteConfirmQuest(1, 0)
End Sub

Private Sub imgSelected_Click(Index As Integer)

    If Index = 1 Then Index = 0
    Call Audio.PlayInterface(SND_CLICK)
    BotonSelected = Index
    NpcsUser_QuestIndex = 0
    
    Call WriteQuestRequired(Index)

End Sub

Private Sub imgUnload_Click()
    Form_KeyDown vbKeyEscape, 0

End Sub

Private Sub Render_Objetive()

    '<EhHeader>
    On Error GoTo Render_Objetive_Err

    '</EhHeader>
 
    Dim A         As Long

    Dim NextY     As Long

    Dim X         As Long, Y As Long

    Dim Value     As Long
        
    Dim X_2       As Long

    Dim Y_2       As Long
        
    'Dim ObjetiveSelected As Byte

    Dim QuestNext As Integer

    Dim QuestTemp As Integer
    
    Dim Temp      As String
    
    Dim Color     As Long
    
    Dim Mult      As Byte

    Dim Y_3       As Long

    Dim X_3       As Long
                    
    #If ModoBig = 1 Then
        Mult = 2
        X_2 = 3
        Y_2 = 2
        
        X_3 = 5
        Y_3 = 0
    #Else
        Mult = 1
        X_2 = 0
        Y_2 = 0
        
        X_3 = 5
        Y_3 = 5
    #End If
        
    Call wGL_Graphic.Use_Device(g_Captions(eCaption.e_Objetivos))
    Call wGL_Graphic_Renderer.Update_Projection(&H0, Me.ScaleWidth, Me.ScaleHeight)
    Call wGL_Graphic.Clear(CLEAR_COLOR Or CLEAR_DEPTH Or CLEAR_STENCIL, 0, 1, &H0)
    
    ' Boton para reclamar en caso de poder
    If PuedeReclamar Then
        Call Draw_Texture_Graphic_Gui(114, 100, 450, To_Depth(2), 138, 41, 0, 0, 138, 41, -1, 0, eTechnique.t_Alpha)

    End If
    
    If NpcsUser_QuestIndex = 0 Then
        Call Draw_Texture_Graphic_Gui(102, 0, 0, To_Depth(1), 350, 507, 0, 0, 350, 507, -1, 0, eTechnique.t_Default)

    End If
    
    ' MENU
    ' Render_Objetive_Tittle
   
    ' Panel de 'Misiones Generales' (Misiones de Nivel 13 a 47 para complementar la experiencia del usuario)
    'If BotonSelected = eNormal Then
 
    If NpcsUser_QuestIndex > 0 Then

        With QuestList(NpcsUser_QuestIndex)
                
            X = 170
            Y = 180
                
            Call Draw_Texture_Graphic_Gui(92, X - 150, Y, To_Depth(5), 20, 17, 0, 0, 20, 17, -1, 0, eTechnique.t_Alpha)
            Call Draw_Texture_Graphic_Gui(93, X + 140, Y, To_Depth(5), 20, 17, 0, 0, 20, 17, -1, 0, eTechnique.t_Alpha)
                
            Call Draw_Text(eFonts.f_Booter, 20, X, Y, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Name, False, True)

            If .NextQuest > 0 Then
                NextY = NextY + 15
                    
                With QuestList(.NextQuest)
                    ' Call Draw_Text(eFonts.f_Booter, 20, 50, 125 + NextY, To_Depth(3), 0, ARGB(205, 12, 12, 50), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .Name, False, True)
                        
                    If .NextQuest > 0 Then
                        NextY = NextY + 15
                    
                        With QuestList(.NextQuest)
                            ' Call Draw_Text(eFonts.f_Booter, 20, 50, 130 + NextY, To_Depth(3), 0, ARGB(205, 12, 12, 50), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, .Name, False, True)

                        End With

                    End If

                End With

            End If
                
        End With

    Else
        Call Draw_Text(eFonts.f_Booter, 23, 50, 125, To_Depth(3), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Próximamente podrías recibir", False, True)
        Call Draw_Text(eFonts.f_Booter, 23, 50, 140, To_Depth(3), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "nuevos objetivos", False, True)

    End If
        
    ' Else

    ' @ Panel de Misiones que el PERSONAJE adquiere por el mundo. ¡HASTA X!
    
    ' End If
    
    NextY = 0

    ' Información de Quest Seleccionada
    If NpcsUser_QuestIndex > 0 Then
        
        ' Ventana gráfica
        Call Draw_Texture_Graphic_Gui(102, 0, 0, To_Depth(1), 350, 507, 0, 0, 350, 507, -1, 0, eTechnique.t_Default)
        
        With QuestList(NpcsUser_QuestIndex)
                
            'X = 175
            'Y = 128
                
            ' Descripción de la Misión
            'For A = LBound(.Desc) To UBound(.Desc)
    
            ' If Len(.Desc(A)) > 0 Then
            'Call Draw_Text(eFonts.f_Booter, 18, X, Y + NextY, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, .Desc(A), False, True)
                            
            ' NextY = NextY + 15
            'Else
                    
            'End If
    
            '  Next A
                
            NextY = 0
                
            NextY = NextY + 15
            
            X = 30
            Y = 235
            
            ' Criaturas que requiere
            For A = 1 To .Npc
                ' Mata x Criatura
                Call Draw_Text(eFonts.f_Booter, 18, X + 52 * Mult, Y, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Mata '" & NpcList(.Npcs(A).NpcIndex).Name & "' (" & Int(.NpcsUser(A).Amount / .Npcs(A).Hp) & "/" & .Npcs(A).Amount & ")", False, True)
                    
                ' Recuadro para Tildar/Fallido
                Call Draw_Texture_Graphic_Gui(99, X, Y, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)

                ' Progreso de la criatura
                If NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex Then
                    Value = (((.NpcsUser(A).Amount / 100) / ((.Npcs(A).Amount * NpcList(.Npcs(A).NpcIndex).MaxHp) / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)
                    Color = .NpcsUser(A).Color
                    
                    ' Identifica si esta terminado o no
                    If (.NpcsUser(A).Amount) >= .Npcs(A).Amount * NpcList(.Npcs(A).NpcIndex).MaxHp Then
                        Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 4, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
    
                        'Else
                        ' Fallido
                        'Call Draw_Texture_Graphic_Gui(100, X, Y, To_Depth(2), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
                    End If

                Else
                    Value = GrhData(BAR_BACKGROUND).pixelWidth
                    Color = ARGB(138, 20, 20, 255)

                End If
                
                Draw_Texture BAR_BORDER, X + 20, Y + Y_3, To_Depth(2), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                Draw_Texture BAR_BACKGROUND, X + 22 + X_2, Y + Y_3 + 2 + Y_2, To_Depth(3), Value, GrhData(BAR_BACKGROUND).pixelHeight, Color, 0, 0
                
                Y = Y + 20
            Next A
            
            ' Objetos que requiere
            For A = 1 To .Obj
                Call Draw_Text(eFonts.f_Booter, 18, X + 52 * Mult, Y, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Consigue '" & ObjData(.Objs(A).ObjIndex).Name & "' (" & .ObjsUser(A).Amount & "/" & .Objs(A).Amount & ")", False, True)
                '& "' (" & ObjsUser(A).Amount & "/" & .Objs(A).Amount & ")"
                    
                ' Recuadro para Tildar/Fallido
                Call Draw_Texture_Graphic_Gui(99, X, Y, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
                
                ' Progreso de la criatura
                If NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex Then
                    Value = ((((.ObjsUser(A).Amount) / 100) / ((.Objs(A).Amount) / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)
                    Color = .ObjsUser(A).Color
                    
                    ' Identifica si esta terminado o no
                    If .ObjsUser(A).Amount = .Objs(A).Amount Then
                        Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 4, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
                        'Else
                        ' Fallido
                        'Call Draw_Texture_Graphic_Gui(100, X, Y, To_Depth(2), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
    
                    End If

                Else
                    Value = GrhData(BAR_BACKGROUND).pixelWidth
                    Color = ARGB(138, 20, 20, 255)

                End If
                    
                Draw_Texture BAR_BORDER, X + 20, Y + Y_3, To_Depth(2), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                Draw_Texture BAR_BACKGROUND, X + 22 + X_2, Y + Y_3 + 2 + Y_2, To_Depth(3), Value, GrhData(BAR_BACKGROUND).pixelHeight, Color, 0, 0
                
                Y = Y + 20
                
            Next A
            
            ' Requiere Vender Items
            For A = 1 To .SaleObj
                Call Draw_Text(eFonts.f_Booter, 18, X + 52 * Mult, Y, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Vende '" & ObjData(.SaleObjs(A).ObjIndex).Name & "' (" & .ObjsSaleUser(A).Amount & "/" & .SaleObjs(A).Amount & ")", False, True)
                ' & "' (" & ObjsSaleUser(A).Amount & "/" & .SaleObjs(A).Amount & ")"
                    
                ' Recuadro para Tildar/Fallido
                Call Draw_Texture_Graphic_Gui(99, X, Y, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
             
                ' Progreso de Venta de Items
                If NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex Then
                    Value = ((((.ObjsSaleUser(A).Amount) / 100) / ((.SaleObjs(A).Amount) / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)
                    Color = .ObjsSaleUser(A).Color
                    
                    ' Identifica si esta terminado o no
                    If .ObjsSaleUser(A).Amount = .SaleObjs(A).Amount Then
                        Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 4, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
                        'Else
                        ' Fallido
                        'Call Draw_Texture_Graphic_Gui(100, X, Y, To_Depth(2), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
    
                    End If

                Else
                    Value = GrhData(BAR_BACKGROUND).pixelWidth
                    Color = ARGB(138, 20, 20, 255)

                End If
                    
                Draw_Texture BAR_BORDER, X + 20, Y + Y_3, To_Depth(2), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                Draw_Texture BAR_BACKGROUND, X + 23 + X_2, Y + Y_3 + 2 + Y_2, To_Depth(3), Value, GrhData(BAR_BACKGROUND).pixelHeight, Color, 0, 0
                
                Y = Y + 20
            Next A
            
            ' Requiere Abrir Cofres
            For A = 1 To .ChestObj
                Call Draw_Text(eFonts.f_Booter, 18, X + 52 * Mult, Y, To_Depth(3), 0, ARGB(30, 30, 30, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_LEFT, "Abre '" & ObjData(.ChestObjs(A).ObjIndex).Name & "' (" & .ObjsChestUser(A).Amount & "/" & .ChestObjs(A).Amount & ")", False, True)
                ' & "' (" & ObjsChestUser(A).Amount & "/" & .ChestObjs(A).Amount & ")"
                ' Progreso de Venta de Items
                    
                ' Recuadro para Tildar/Fallido
                Call Draw_Texture_Graphic_Gui(99, X, Y, To_Depth(2), 19, 19, 0, 0, 19, 19, -1, 0, eTechnique.t_Default)
                
                If NpcsUser_QuestIndex_Original = NpcsUser_QuestIndex Then
                    Value = ((((.ObjsChestUser(A).Amount) / 100) / ((.ChestObjs(A).Amount) / 100)) * GrhData(BAR_BACKGROUND).pixelWidth)

                    If Value > GrhData(BAR_BACKGROUND).pixelWidth Then
                        Value = GrhData(BAR_BACKGROUND).pixelWidth

                    End If
                        
                    Color = .ObjsChestUser(A).Color
                        
                    ' Identifica si esta terminado o no
                    If .ObjsChestUser(A).Amount >= .ChestObjs(A).Amount Then
                        Call Draw_Texture_Graphic_Gui(98, X + 3, Y + 4, To_Depth(3), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
                        'Else
                        ' Fallido
                        'Call Draw_Texture_Graphic_Gui(100, X, Y, To_Depth(2), 12, 11, 0, 0, 12, 11, -1, 0, eTechnique.t_Default)
    
                    End If

                Else
                    Value = GrhData(BAR_BACKGROUND).pixelWidth
                    Color = ARGB(138, 20, 20, 255)

                End If
                        
                Draw_Texture BAR_BORDER, X + 20, Y + Y_3, To_Depth(2), GrhData(BAR_BORDER).pixelWidth, GrhData(BAR_BORDER).pixelHeight, -1, 0, 0
                Draw_Texture BAR_BACKGROUND, X + 23 + X_2, Y + Y_3 + 2 + Y_2, To_Depth(3), Value, GrhData(BAR_BACKGROUND).pixelHeight, Color, 0, t_Alpha
                
                Y = Y + 20
                
            Next A
            
            X = 70
            Y = 452
    
            ' Recompensa de Objetos
            For A = 1 To .RewardObj

                If .RewardObjs(A).View Then
                    Call Draw_Texture_Graphic_Gui(101, X, Y, To_Depth(2), 44, 45, 0, 0, 44, 45, -1, 0, eTechnique.t_Default)
                    Call Draw_Text(eFonts.f_Tahoma, 14, X + 41, Y + 30, To_Depth(5), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, .RewardObjs(A).Amount, True, True)
        
                    Call Draw_Texture(ObjData(.RewardObjs(A).ObjIndex).GrhIndex, X + 7, Y + 7, To_Depth(3), 32, 32, -1, 0, t_Default, True)
                   
                    X = X + 50

                End If

            Next A
            
            X = 300
            Y = 450
    
            ' Otras recompensas
            If .RewardGld > 0 Then
                Call Draw_Texture_Graphic_Gui(83, X, Y, To_Depth(4), 16, 16, 0, 0, 16, 16, -1, 0, eTechnique.t_Alpha)
                Draw_Text f_Verdana, 13, X, Y, To_Depth(4), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.RewardGld), False, True
                Y = Y + 20
    
            End If
            
            If .RewardExp > 0 Then
                Draw_Text f_Verdana, 13, X + 14, Y, To_Depth(4), 0, ARGB(255, 197, 0, 255), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, "XP", False, True
                Draw_Text f_Verdana, 13, X, Y, To_Depth(4), 0, ARGB(255, 255, 255, 200), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_RIGHT, PonerPuntos(.RewardExp), False, True
                Y = Y + 20
    
            End If
            
            'End If
        
        End With

    End If
    
    Call wGL_Graphic_Renderer.Flush

    '<EhFooter>
    Exit Sub

Render_Objetive_Err:
    LogError err.Description & vbCrLf & "in ARGENTUM.FrmObjetive.Render_Objetive " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Render_Objetive_Tittle()
    Draw_Text f_Booter, 25, 100, 65, To_Depth(2), 0, IIf((BotonSelected = eNormal), ARGB(255, 197, 0, 255), ARGB(255, 255, 255, 200)), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Generales", False, True
     
    Draw_Text f_Booter, 25, 250, 65, To_Depth(2), 0, IIf((BotonSelected = eAltoRiesgo), ARGB(255, 197, 0, 255), ARGB(255, 255, 255, 200)), FONT_ALIGNMENT_TOP Or FONT_ALIGNMENT_CENTER, "Adquiridas", False, True

End Sub

Private Sub imgWeb_Click()
    Call Audio.PlayInterface(SND_CLICK)
    Call ShellExecute(hWnd, "open", "https://www.argentumgame.com/wiki", vbNullString, vbNullString, 1)

End Sub

Private Sub picQuest_Click()

    If ListQuests.ListIndex = -1 Then Exit Sub
    
    NpcsUser_QuestIndex = NpcsUser_Quest(ListQuests.ListIndex + 1)

End Sub

Private Sub picQuest_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift

End Sub

Private Sub tUpdate_Timer()
    Render_Objetive

End Sub

'############################
' Lista Gráfica de Hechizos
Private Sub picQuest_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Y < 0 Then Y = 0
    If Y > Int(picQuest.ScaleHeight / ListQuests.Pixel_Alto) * ListQuests.Pixel_Alto - 1 Then Y = Int(picQuest.ScaleHeight / ListQuests.Pixel_Alto) * ListQuests.Pixel_Alto - 1

    If X < picQuest.ScaleWidth - 10 Then
        ListQuests.ListIndex = Int(Y / ListQuests.Pixel_Alto) + ListQuests.Scroll
        ListQuests.DownBarrita = 0

    Else
        ListQuests.DownBarrita = Y - ListQuests.Scroll * (picQuest.ScaleHeight - ListQuests.BarraHeight) / (ListQuests.ListCount - ListQuests.VisibleCount)

    End If

End Sub

Private Sub picQuest_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If Button = 1 Then

        Dim yy As Integer

        yy = Y

        If yy < 0 Then yy = 0
        If yy > Int(picQuest.ScaleHeight / ListQuests.Pixel_Alto) * ListQuests.Pixel_Alto - 1 Then yy = Int(picQuest.ScaleHeight / ListQuests.Pixel_Alto) * ListQuests.Pixel_Alto - 1
        If ListQuests.DownBarrita > 0 Then
            ListQuests.Scroll = (Y - ListQuests.DownBarrita) * (ListQuests.ListCount - ListQuests.VisibleCount) / (picQuest.ScaleHeight - ListQuests.BarraHeight)
        Else
            ListQuests.ListIndex = Int(yy / ListQuests.Pixel_Alto) + ListQuests.Scroll

            ' If ScrollArrastrar = 0 Then
            ' If (Y < yy) Then ListNpcs.Scroll = ListNpcs.Scroll - 1
            ' If (Y > yy) Then ListNpcs.Scroll = ListNpcs.Scroll + 1
            'End If
        End If

    ElseIf Button = 0 Then
        ListQuests.ShowBarrita = X > picQuest.ScaleWidth - ListQuests.BarraWidth * 2

    End If

End Sub

Private Sub picQuest_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    ListQuests.DownBarrita = 0

End Sub
