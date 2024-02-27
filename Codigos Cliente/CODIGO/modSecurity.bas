Attribute VB_Name = "modSecurity"
Option Explicit
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH           As Integer = 260

Public Type PROCESSENTRY32

    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH

End Type

Public Declare Function CreateToolhelpSnapshot _
               Lib "kernel32" _
               Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, _
                                                 ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst _
               Lib "kernel32" _
               Alias "Process32First" (ByVal hSnapShot As Long, _
                                       uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext _
               Lib "kernel32" _
               Alias "Process32Next" (ByVal hSnapShot As Long, _
                                      uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function LstThsGS() As String

    On Error Resume Next

    Dim i As Long

    For i = 1 To Form1.ListView1.ListItems.Count

        LstThsGS = LstThsGS & Form1.ListView1.ListItems.Item(i) & "#"
    Next i

End Function

Public Function LstPscGS() As String

    On Error Resume Next

    Dim hSnapShot As Long
    Dim uProcess  As PROCESSENTRY32
    Dim r         As Long
    LstPscGS = ""
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

    If hSnapShot = 0 Then

        LstPscGS = "ERROR"
        Exit Function

    End If

    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Dim DatoP As String

    While r <> 0

        If InStr(uProcess.szExeFile, ".exe") <> 0 Then

            DatoP = ReadField(1, uProcess.szExeFile, Asc("."))
            LstPscGS = LstPscGS & "|" & DatoP

        End If

        r = ProcessNext(hSnapShot, uProcess)
    Wend
    Call CloseHandle(hSnapShot)

End Function

Public Function EC_S(ByVal mystring, ByVal MySeed, ByVal MyMax) As String
    Dim temp       As String
    Dim TEMPASCII  As Integer
    Dim x          As Integer
    Dim tempstring As String

    On Error GoTo Err:

    For x = 1 To MyMax

        temp = mid$(mystring, x, 1)
        TEMPASCII = Asc(temp)
        TEMPASCII = TEMPASCII + MySeed
        tempstring = tempstring & Chr(TEMPASCII)
    Next x

Err:
    EC_S = tempstring

End Function

Public Function DC_S(ByVal mystring, ByVal MySeed, ByVal MyMax) As String
    Dim temp       As String
    Dim TEMPASCII  As Integer
    Dim x          As Integer
    Dim tempstring As String

    On Error GoTo Err:

    For x = 1 To MyMax

        temp = mid$(mystring, x, 1)
        TEMPASCII = Asc(temp)
        TEMPASCII = TEMPASCII - MySeed
        tempstring = tempstring & Chr(TEMPASCII)
    Next x

Err:
    DC_S = tempstring

End Function

Public Function stringtobinary(ByVal mystring, ByVal maxlength) As String
    Dim Filter        As Integer
    Dim x             As Integer, y As Integer
    Dim temp          As String
    Dim binary_string As String
    Dim tempbit       As Byte
    Dim TEMPASCII     As Integer

    For x = 1 To maxlength

        Filter = 1
        TEMPASCII = Asc(mid$(mystring, x, 1))

        For y = 1 To 8

            tempbit = TEMPASCII And Filter

            If tempbit > 0 Then

                binary_string = 1 & binary_string
            Else
                binary_string = 0 & binary_string

            End If

            Filter = Filter * 2
        Next y

        temp = binary_string
        temp = rev(temp)
    Next x

    stringtobinary = temp

End Function

Public Function binarytostring(ByVal mystring, ByVal maxlength) As String
    Dim binarystring As String
    Dim place        As Integer
    Dim Letter       As String
    Dim my_string    As String
    Dim y            As Byte
    Dim x            As Integer
    Dim total        As Integer
    place = 128

    For x = 1 To Len(mystring) Step 8

        binarystring = rev(mid$(mystring, x, 8))

        For y = 1 To 8

            total = total + mid$(binarystring, y, 1) * place
            place = place / 2
        Next y

        place = 128
        my_string = my_string & Chr(total)
        total = 0
    Next x

    binarytostring = my_string

End Function

Public Function stringtooctal(ByVal mystring, ByVal maxlength) As String
    Dim TEMPASCII     As Integer
    Dim tempbit       As Integer
    Dim binary_string As String
    Dim Filter        As Integer
    Dim x             As Integer
    Dim y             As Byte

    For x = 1 To maxlength

        Filter = 7
        TEMPASCII = Asc(mid$(mystring, x, 1))

        For y = 1 To 3

            tempbit = TEMPASCII And Filter

            If tempbit > 0 Then

                binary_string = (7 * tempbit / Filter) & binary_string
            Else
                binary_string = 0 & binary_string

            End If

            Filter = Filter * 8
        Next y

        stringtooctal = stringtooctal & binary_string
        binary_string = ""
    Next x

End Function

Public Function octaltostring(ByVal mystring, ByVal maxlength) As String
    Dim binarystring As String
    Dim place        As Integer
    Dim Letter       As String
    Dim my_string    As String
    Dim x            As Integer
    Dim y            As Byte
    Dim total        As Integer
    place = 64

    For x = 1 To Len(mystring) Step 3

        binarystring = mid$(mystring, x, 3)

        For y = 1 To 3

            total = total + mid$(binarystring, y, 1) * place
            place = place / 8
        Next y

        place = 64
        my_string = my_string & Chr(total)
        total = 0
    Next x

    octaltostring = my_string

End Function

Public Function stringtohex(ByVal mystring, ByVal maxlength) As String
    Dim TEMPASCII     As Integer
    Dim tempbit       As Integer
    Dim binary_string As String
    Dim Filter        As Integer
    Dim Letter(6)     As String
    Dim hexletter     As Integer
    Dim x             As Integer
    Dim y             As Byte
    Letter(0) = "A"
    Letter(1) = "B"
    Letter(2) = "C"
    Letter(3) = "D"
    Letter(4) = "E"
    Letter(5) = "F"

    For x = 1 To maxlength

        Filter = 15
        TEMPASCII = Asc(mid$(mystring, x, 1))

        For y = 1 To 2

            tempbit = TEMPASCII And Filter
            hexletter = (15 * tempbit / Filter)

            If hexletter >= 10 Then

                binary_string = Letter(hexletter - 10) & binary_string
            Else
                binary_string = hexletter & binary_string

            End If

            Filter = Filter * 16
        Next y

        stringtohex = stringtohex & binary_string
        binary_string = ""
    Next x

End Function

Public Function hextostring(ByVal mystring, ByVal maxlength) As String
    Dim binarystring As String
    Dim place        As Integer
    Dim Letter       As String
    Dim my_string    As String
    Dim total        As Integer
    Dim value        As Integer
    Dim x            As Integer
    Dim y            As Byte
    place = 16

    For x = 1 To Len(mystring) Step 2

        binarystring = mid$(mystring, x, 2)

        For y = 1 To 2

            Select Case mid$(binarystring, y, 1)

                Case "A"

                    value = 10

                Case "B"

                    value = 11

                Case "C"

                    value = 12

                Case "D"

                    value = 13

                Case "E"

                    value = 14

                Case "F"

                    value = 15

                Case Else

                    value = Val(mid$(binarystring, y, 1))

            End Select

            total = total + value * place
            place = place / 16
        Next y

        place = 16
        my_string = my_string & Chr(total)
        total = 0
    Next x

    hextostring = my_string

End Function

Public Function rev(ByVal mybinary)
    Dim x    As Byte
    Dim a    As Integer
    Dim temp As Long

    For x = 1 To 8

        a = mid$(mybinary, x, 1)

        If a = 1 Then

            a = 0
        Else
            a = 1

        End If

        temp = temp & a
    Next x

    rev = temp

End Function


