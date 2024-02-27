Attribute VB_Name = "Mod_Compresion"
Option Explicit
 
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal F As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal F As String, ByVal t As Long, ByVal r As String)
 
Public Function MD5String(ByVal p As String) As String
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function
 
Public Function MD5File(ByVal F As String) As String
    Dim r As String * 32
    r = Space(32)
    MDFile F, r
    MD5File = r
End Function
