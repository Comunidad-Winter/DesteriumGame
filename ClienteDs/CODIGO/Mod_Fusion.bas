Attribute VB_Name = "Mod_Compresion"
Option Explicit
 
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal F As String, ByVal r As _
    String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal F As String, ByVal t _
    As Long, ByVal r As String)
 
Public Function MD5String(ByVal p As String) As String
          Dim r As String * 32, t As Long
10        r = Space(32)
20        t = Len(p)
30        MDStringFix p, t, r
40        MD5String = r
End Function
 
Public Function MD5File(ByVal F As String) As String
          Dim r As String * 32
10        r = Space(32)
20        MDFile F, r
30        MD5File = r
End Function

