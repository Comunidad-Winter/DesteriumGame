Attribute VB_Name = "m_Damages"
Option Explicit
 
Const DAMAGE_TIME   As Integer = 57
Const DAMAGE_FONT_S As Byte = 17
 
Enum EDType
     edPuñal = 1                'Apuñalo.
     edNormal = 2               'Hechizo o golpe común.
End Enum

Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Long  'Cantidad de daño.
     ColorRGB       As Long     'Color.
     DamageType     As EDType   'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     TimeRendered   As Integer  'Tiempo transcurrido.
     Downloading    As Byte     'Contador para la posicion Y.
     Activated      As Boolean  'Si está activado..
End Type
 
Sub Initialize()
       
      'INICIA EL FONTTYPE
       
10    With DNormalFont
           
20         .Size = 8
30         .italic = False
40         .bold = True
50         .Name = "Tahoma"
           
60    End With
       
       
End Sub
 
Sub Create(ByVal X As Byte, ByVal Y As Byte, ByVal ColorRGB As Long, ByVal _
    DamageValue As Long, ByVal edMode As Byte)

       
      'INICIA EL FONTTYPE APU
       
10    With MapData(X, Y).Damage
           
20         .Activated = True
30         .ColorRGB = ColorRGB
40         .DamageType = edMode
50         .DamageVal = DamageValue
60         .TimeRendered = 0
70         .Downloading = 0
           

80            With .DamageFont
90                 .Size = 8
100                .Name = "Tahoma"
110                .bold = True
120                Exit Sub
130           End With


140        .DamageFont = DNormalFont
150        .DamageFont.Size = 8
           
160   End With
       
End Sub

Sub Draw(ByVal X As Byte, ByVal Y As Byte, ByVal PixelX As Integer, ByVal _
    PixelY As Integer)
       
      ' @ Dibuja un daño
       
10    With MapData(X, Y).Damage
           
20         If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
30            If .TimeRendered < DAMAGE_TIME Then
                 
                 'Sumo el contador del tiempo.
40               .TimeRendered = .TimeRendered + 1
                 
50               If (.TimeRendered / 2) > 0 Then
60                   .Downloading = (.TimeRendered / 2)
70               End If
                 
80               .ColorRGB = ModifyColour(.TimeRendered, .DamageType)
                     
                     
                 #If Wgl = 1 Then
                    EngineWgl.Draw_Text f_Tahoma, 16, PixelX, PixelY - .Downloading, 0#, 0#, .ColorRGB, FONT_ALIGNMENT_BASELINE, "" & .DamageVal, True
                 #Else
90                  RenderTextCentered PixelX, PixelY - .Downloading, "" & .DamageVal, .ColorRGB, .DamageFont, False
                 #End If
                 
                 'Si llego al tiempo lo limpio
100              If .TimeRendered >= DAMAGE_TIME Then
110                 Clear X, Y
120              End If
                 
130        End If
             
140   End With
       
End Sub
 
Sub Clear(ByVal X As Byte, ByVal Y As Byte)
       
      ' @ Limpia todo.
       
10    With MapData(X, Y).Damage
20         .Activated = False
30         .ColorRGB = 0
40         .DamageVal = 0
50         .TimeRendered = 0
60    End With
       
End Sub
 
Function ModifyColour(ByVal TimeNowRendered As Byte, ByVal DamageType As Byte) _
    As Long
      ' @ Se usa para el "efecto" de desvanecimiento.
       
10    Select Case DamageType
                         
             Case EDType.edPuñal
20                ModifyColour = RGB(255, 255, 184)
                  'ModifyColour = GetPuñalNewColour()
                         
30           Case EDType.edNormal
40                ModifyColour = RGB(255 - (TimeNowRendered * 3), 0, 0)
50    End Select
       
End Function
