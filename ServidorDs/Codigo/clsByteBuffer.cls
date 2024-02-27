VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                    ByRef destination As Any, _
                                    ByRef source As Any, _
                                    ByVal length As Long)

Private data() As Byte

Private CurrentPos As Long
Private lastPos As Long
Private hFile As Long

Private Const INVALID_FILE_HANDLER As Integer = -1
Private Const DEFAULT_MAX_SIZE_FILE As Long = 65535
Private Const BYTE_SIZE As Byte = 1
Private Const BOOL_SIZE As Byte = 2
Private Const INTEGER_SIZE As Byte = 2
Private Const LONG_SIZE As Byte = 4
Private Const SINGLE_SIZE As Byte = 4
Private Const DOUBLE_SIZE As Byte = 8
Private Const STRING_LENGTH_SIZE As Byte = 2

Public Sub initializeReader(ByRef arrayByte() As Byte)
10    lastPos = UBound(arrayByte)
20    ReDim data(lastPos) As Byte

30    Call CopyMemory(data(0), arrayByte(0), lastPos + 1)
40    CurrentPos = 0
End Sub

Public Sub initializeWriter(ByVal fileHandler As Integer)
10    ReDim data(DEFAULT_MAX_SIZE_FILE * 20) As Byte

20    hFile = fileHandler
30    CurrentPos = 0
40    lastPos = -1
End Sub

Public Sub getBytes(ByRef destination() As Byte, Optional ByVal length As Long = -1)
10    If length >= 0 Then
20        Call CopyMemory(destination(0), data(CurrentPos), length)
30    Else
40        Call CopyMemory(destination(0), data(0), lastPos + 1)
50    End If
End Sub

Public Function getByte() As Byte
10    getByte = data(CurrentPos)
20    CurrentPos = CurrentPos + BYTE_SIZE
End Function

Public Function getBoolean() As Boolean
10    Call CopyMemory(getBoolean, data(CurrentPos), BOOL_SIZE)
20    CurrentPos = CurrentPos + BOOL_SIZE
End Function

Public Function getInteger() As Integer
10    Call CopyMemory(getInteger, data(CurrentPos), INTEGER_SIZE)
20    CurrentPos = CurrentPos + INTEGER_SIZE
End Function

Public Function getLong() As Long
10    Call CopyMemory(getLong, data(CurrentPos), LONG_SIZE)
20    CurrentPos = CurrentPos + LONG_SIZE
End Function

Public Function getSingle() As Single
10    Call CopyMemory(getSingle, data(CurrentPos), SINGLE_SIZE)
20    CurrentPos = CurrentPos + SINGLE_SIZE
End Function

Public Function getDouble() As Double
10    Call CopyMemory(getDouble, data(CurrentPos), DOUBLE_SIZE)
20    CurrentPos = CurrentPos + DOUBLE_SIZE
End Function

Public Function getString(Optional ByVal length As Integer = -1) As String
      Dim ret() As Byte

10    If length < 0 Then
20        length = getInteger()
30        getString = getString(length)
40    Else
50        If length > 0 Then
60            ReDim ret(length - 1) As Byte
              
70            Call CopyMemory(ret(0), data(CurrentPos), length)
              
80            getString = StrConv(ret, vbUnicode)
90            CurrentPos = CurrentPos + length
100       End If
110   End If

End Function

Public Sub putByte(ByVal value As Byte)
10    data(lastPos + 1) = value
20    lastPos = lastPos + BYTE_SIZE
End Sub

Public Sub putBoolean(ByVal value As Boolean)
10    Call CopyMemory(data(lastPos + 1), value, BOOL_SIZE)
20    lastPos = lastPos + BOOL_SIZE
End Sub

Public Sub putInteger(ByVal value As Integer)
10    Call CopyMemory(data(lastPos + 1), value, INTEGER_SIZE)
20    lastPos = lastPos + INTEGER_SIZE
End Sub

Public Sub putLong(ByVal value As Long)
10    Call CopyMemory(data(lastPos + 1), value, LONG_SIZE)
20    lastPos = lastPos + LONG_SIZE
End Sub

Public Sub putSingle(ByVal value As Single)
10    Call CopyMemory(data(lastPos + 1), value, SINGLE_SIZE)
20    lastPos = lastPos + SINGLE_SIZE
End Sub

Public Sub putDouble(ByVal value As Double)
10    Call CopyMemory(data(lastPos + 1), value, DOUBLE_SIZE)
20    lastPos = lastPos + DOUBLE_SIZE
End Sub

Public Sub putString(ByRef Str As String, Optional ByVal withLength As Boolean = True)
      Dim length As Long

10    length = Len(Str)

20    If withLength Then
30        Call putInteger(length)
40        Call putString(Str, False)
50    Else
60        If length > 0 Then
70            Call CopyMemory(data(lastPos + 1), ByVal StrPtr(StrConv(Str, vbFromUnicode)), length)
          
80            lastPos = lastPos + length
90        End If
100   End If
End Sub

Public Sub getVoid(ByVal length As Integer)
10    CurrentPos = CurrentPos + length
End Sub

Public Sub putVoid(ByVal length As Integer)
10    lastPos = lastPos + length
End Sub

Public Sub clearData()
10    ReDim data(DEFAULT_MAX_SIZE_FILE) As Byte

20    CurrentPos = 0
30    lastPos = -1
40    hFile = -1
End Sub

Public Function getLastPos() As Long
10    getLastPos = lastPos
End Function

Public Function getCurrentPos() As Long
10    getCurrentPos = CurrentPos
End Function

Public Function EOF() As Boolean
10    EOF = (CurrentPos > UBound(data))
End Function

Public Sub saveBuffer()
      Dim buf() As Byte

10    If hFile > 0 Then
20        ReDim buf(lastPos) As Byte
          
30        Call CopyMemory(buf(0), data(0), lastPos + 1)
40        Put hFile, , buf
50    End If
End Sub

Private Sub Class_Initialize()
10    hFile = INVALID_FILE_HANDLER
End Sub

Private Sub Class_Terminate()
10    Erase data()
End Sub


