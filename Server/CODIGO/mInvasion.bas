Attribute VB_Name = "mInvasion"
Option Explicit

Public Type t_Rectangle

    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer

End Type

Type t_SpawnBox

    TopLeft As WorldPos
    BottomRight As WorldPos
    Heading As eHeading
    CoordMuralla As Integer
    LegalBox As t_Rectangle

End Type

' WyroX: Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(r As t_Rectangle, _
                                ByVal X As Integer, _
                                ByVal Y As Integer) As Boolean

    If X < r.X1 Then Exit Function
    If X > r.X2 Then Exit Function
    If Y < r.Y1 Then Exit Function
    If Y > r.Y2 Then Exit Function
    InsideRectangle = True

End Function
