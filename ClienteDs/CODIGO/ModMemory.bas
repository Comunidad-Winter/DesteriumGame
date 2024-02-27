Attribute VB_Name = "ModResolution"
Option Explicit
 
 
Public Enum eIntervalos
       USE_ITEM_U = 1       'Usar items con la U.
       CAST_ATTACK          'Hechizo - ataque.
       CAST_SPELL           'Hechizo - hechizo.
       Attack               'Golpe - golpe.
End Enum
 
Public INT_ATTACK           As Long
Public INT_CAST_ATTACK      As Long
Public INT_CAST_SPELL       As Long
Public INT_USEITEMU         As Long
 
Public Sub Generate_Array()
10    INT_ATTACK = 1200
20    INT_CAST_SPELL = 1100
30    INT_CAST_ATTACK = 1000
40    INT_USEITEMU = 435
       
End Sub


