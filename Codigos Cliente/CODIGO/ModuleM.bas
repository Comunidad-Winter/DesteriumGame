Attribute VB_Name = "ModuleM"
Option Explicit
 
'programado por maTih.-
 
Public Enum eIntervalos
       USE_ITEM_U = 1       'Usar items con la U.
       CAST_ATTACK          'Hechizo - ataque.
       CAST_SPELL           'Hechizo - hechizo.
       Attack               'Golpe - golpe.
End Enum
 
Public Const INT_PASSWORD   As Long = 555
 
Public INT_ATTACK           As Long
Public INT_CAST_ATTACK      As Long
Public INT_CAST_SPELL       As Long
Public INT_USEITEMU         As Long
 
Public e_Pointers(1 To 4)   As Long
Public e_Intervals(1 To 4)  As Long
Public int_Memory           As New clsMemory
 
Public Sub Generate_Array()
 
'
' @ Acá van los intervalos !
 
'Dim null_Array(1 To 1) As Long 'usamos esto para inicializar la clase : p
Dim loopC As Long
 
INT_ATTACK = 1200
INT_CAST_SPELL = 1100
INT_CAST_ATTACK = 900
INT_USEITEMU = 450
 
'setea los intervalos del array original.
e_Intervals(eIntervalos.Attack) = INT_ATTACK
e_Intervals(eIntervalos.CAST_ATTACK) = INT_CAST_ATTACK
e_Intervals(eIntervalos.CAST_SPELL) = INT_CAST_SPELL
e_Intervals(eIntervalos.USE_ITEM_U) = INT_USEITEMU
 
Call int_Memory.Initialize(e_Intervals(), INT_PASSWORD)
 
End Sub
 
Public Sub Check_All()
 
'
' @ Busca edición en los intervalos.
 
Dim loopC   As Long
 
For loopC = eIntervalos.USE_ITEM_U To eIntervalos.Attack
    'Si el dato original es distinto al encriptado entonces está editado.
    If int_Memory.Return_Element_Original(loopC) <> int_Memory.Return_Element_Decrypted(loopC) Then
       MsgBox "Se detectó una edición en la memoria del juego, el mismo se cerrará.", vbCritical
       Call Mod_General.CloseClient
       Exit For
    End If
Next loopC
 
End Sub
