Attribute VB_Name = "Carteles"
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

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Integer)
10    If Not Cartel Then
20        Leyenda = Ley
30        textura = Grh
40        Cartel = True
50        ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                      
          Dim i As Integer, k As Integer, anti As Integer
60        anti = 1
70        k = 0
80        i = 0
90        Call DarFormato(Leyenda, i, k, anti)
100       i = 0
110       Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
              
120          i = i + 1
130       Loop
140       ReDim Preserve LeyendaFormateada(0 To i)
150   Else
160       Exit Sub
170   End If
End Sub


Private Function DarFormato(s As String, i As Integer, k As Integer, anti As _
    Integer)
10    If anti + i <= Len(s) + 1 Then
20        If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) _
              Then
30            LeyendaFormateada(k) = mid(s, anti, i + 1)
40            k = k + 1
50            anti = anti + i + 1
60            i = 0
70        Else
80            i = i + 1
90        End If
100       Call DarFormato(s, i, k, anti)
110   End If
End Function


Sub DibujarCartel()
10    If Not Cartel Then Exit Sub
      Dim X As Integer, Y As Integer
20    X = XPosCartel + 20
30    Y = YPosCartel + 60
40    Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0)
      Dim j As Integer, desp As Integer

50    For j = 0 To UBound(LeyendaFormateada)
60        RenderText X, Y + desp, LeyendaFormateada(j), vbWhite, frmMain.font
70        desp = desp + (frmMain.font.Size) + 5
80    Next
End Sub


