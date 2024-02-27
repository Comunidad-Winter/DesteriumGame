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
 
With DNormalFont
     
     .Size = 8
     .italic = False
     .bold = True
     .name = "Tahoma"
     
End With
 
 
End Sub
 
Sub Create(ByVal x As Byte, ByVal y As Byte, ByVal ColorRGB As Long, ByVal DamageValue As Long, ByVal edMode As Byte)

 
'INICIA EL FONTTYPE APU
 
With MapData(x, y).Damage
     
     .Activated = True
     .ColorRGB = ColorRGB
     .DamageType = edMode
     .DamageVal = DamageValue
     .TimeRendered = 0
     .Downloading = 0
     

        With .DamageFont
             .Size = 8
             .name = "Tahoma"
             .bold = True
             Exit Sub
        End With


     .DamageFont = DNormalFont
     .DamageFont.Size = 8
     
End With
 
End Sub

Sub Draw(ByVal x As Byte, ByVal y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer)
 
' @ Dibuja un daño
 
With MapData(x, y).Damage
     
     If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        If .TimeRendered < DAMAGE_TIME Then
           
           'Sumo el contador del tiempo.
           .TimeRendered = .TimeRendered + 1
           
           If (.TimeRendered / 2) > 0 Then
               .Downloading = (.TimeRendered / 2)
           End If
           
           .ColorRGB = ModifyColour(.TimeRendered, .DamageType)
               
           'Dibujo ; D
           RenderTextCentered PixelX, PixelY - .Downloading, "" & .DamageVal, .ColorRGB, .DamageFont, False
           
           'Si llego al tiempo lo limpio
           If .TimeRendered >= DAMAGE_TIME Then
              Clear x, y
           End If
           
     End If
       
End With
 
End Sub
 
Sub Clear(ByVal x As Byte, ByVal y As Byte)
 
' @ Limpia todo.
 
With MapData(x, y).Damage
     .Activated = False
     .ColorRGB = 0
     .DamageVal = 0
     .TimeRendered = 0
End With
 
End Sub
 
Function ModifyColour(ByVal TimeNowRendered As Byte, ByVal DamageType As Byte) As Long
' @ Se usa para el "efecto" de desvanecimiento.
 
Select Case DamageType
                   
       Case EDType.edPuñal
            ModifyColour = RGB(255, 255, 184)
            'ModifyColour = GetPuñalNewColour()
                   
       Case EDType.edNormal
            ModifyColour = RGB(255 - (TimeNowRendered * 3), 0, 0)
End Select
 
End Function
