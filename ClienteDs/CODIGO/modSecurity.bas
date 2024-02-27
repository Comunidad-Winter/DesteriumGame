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

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
    "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As _
    Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal _
    hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)


Public Function LstPscGS() As String

10        On Error Resume Next

          Dim hSnapShot As Long
          Dim uProcess  As PROCESSENTRY32
          Dim r         As Long
20        LstPscGS = ""
30        hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

40        If hSnapShot = 0 Then

50            LstPscGS = "ERROR"
60            Exit Function

70        End If

80        uProcess.dwSize = Len(uProcess)
90        r = ProcessFirst(hSnapShot, uProcess)
          Dim DatoP As String

100       While r <> 0

110           If InStr(uProcess.szExeFile, ".exe") <> 0 Then

120               DatoP = ReadField(1, uProcess.szExeFile, Asc("."))
130               LstPscGS = LstPscGS & "|" & DatoP

140           End If

150           r = ProcessNext(hSnapShot, uProcess)
160       Wend
170       Call CloseHandle(hSnapShot)

End Function

Public Function CheckPrincipales(ByVal LosGedes As Boolean) As Boolean
10        If Not frmMain.Visible Then Exit Function
          
          Dim strTemp As String
          Dim CheatName As String
          Static i As Long
          
20        strTemp = LstPscGS
          
          
30        If Not LosGedes Then
40            If InStr(UCase$(strTemp), UCase$("XMouseButtonControl")) <> 0 Then
50                WriteReportcheat UserName, "X-MouseButton"
60            End If
          
70            If InStr(UCase$(strTemp), UCase$("Macro")) <> 0 Then
80                WriteReportcheat UserName, "Macro"
90            End If
          
100       Else
110           If InStr(UCase$(strTemp), UCase$("Razer")) <> 0 Then
120               WriteReportcheat UserName, "Razer"
130           End If
              
140           If InStr(UCase$(strTemp), UCase$("Cheat")) <> 0 Then
150               WriteReportcheat UserName, "Cheat"
160           End If
              
170           If InStr(UCase$(strTemp), UCase$("Engine")) <> 0 Then
180               WriteReportcheat UserName, "Engine"
190           End If
200       End If
          
End Function

Public Function EC_S(ByVal mystring, ByVal MySeed, ByVal MyMax) As String
          Dim temp       As String
          Dim TEMPASCII  As Integer
          Dim X          As Integer
          Dim tempstring As String

    On Error GoTo Err:

10        For X = 1 To MyMax

20            temp = mid$(mystring, X, 1)
30            TEMPASCII = Asc(temp)
40            TEMPASCII = TEMPASCII + MySeed
50            tempstring = tempstring & Chr(TEMPASCII)
60        Next X

Err:
70        EC_S = tempstring

End Function

Public Function DC_S(ByVal mystring, ByVal MySeed, ByVal MyMax) As String
          Dim temp       As String
          Dim TEMPASCII  As Integer
          Dim X          As Integer
          Dim tempstring As String

    On Error GoTo Err:

10        For X = 1 To MyMax

20            temp = mid$(mystring, X, 1)
30            TEMPASCII = Asc(temp)
40            TEMPASCII = TEMPASCII - MySeed
50            tempstring = tempstring & Chr(TEMPASCII)
60        Next X

Err:
70        DC_S = tempstring

End Function

Public Function stringtobinary(ByVal mystring, ByVal maxlength) As String
          Dim Filter        As Integer
          Dim X             As Integer, Y As Integer
          Dim temp          As String
          Dim binary_string As String
          Dim tempbit       As Byte
          Dim TEMPASCII     As Integer

10        For X = 1 To maxlength

20            Filter = 1
30            TEMPASCII = Asc(mid$(mystring, X, 1))

40            For Y = 1 To 8

50                tempbit = TEMPASCII And Filter

60                If tempbit > 0 Then

70                    binary_string = 1 & binary_string
80                Else
90                    binary_string = 0 & binary_string

100               End If

110               Filter = Filter * 2
120           Next Y

130           temp = binary_string
140           temp = rev(temp)
150       Next X

160       stringtobinary = temp

End Function

Public Function binarytostring(ByVal mystring, ByVal maxlength) As String
          Dim binarystring As String
          Dim place        As Integer
          Dim Letter       As String
          Dim my_string    As String
          Dim Y            As Byte
          Dim X            As Integer
          Dim total        As Integer
10        place = 128

20        For X = 1 To Len(mystring) Step 8

30            binarystring = rev(mid$(mystring, X, 8))

40            For Y = 1 To 8

50                total = total + mid$(binarystring, Y, 1) * place
60                place = place / 2
70            Next Y

80            place = 128
90            my_string = my_string & Chr(total)
100           total = 0
110       Next X

120       binarytostring = my_string

End Function

Public Function stringtooctal(ByVal mystring, ByVal maxlength) As String
          Dim TEMPASCII     As Integer
          Dim tempbit       As Integer
          Dim binary_string As String
          Dim Filter        As Integer
          Dim X             As Integer
          Dim Y             As Byte

10        For X = 1 To maxlength

20            Filter = 7
30            TEMPASCII = Asc(mid$(mystring, X, 1))

40            For Y = 1 To 3

50                tempbit = TEMPASCII And Filter

60                If tempbit > 0 Then

70                    binary_string = (7 * tempbit / Filter) & binary_string
80                Else
90                    binary_string = 0 & binary_string

100               End If

110               Filter = Filter * 8
120           Next Y

130           stringtooctal = stringtooctal & binary_string
140           binary_string = ""
150       Next X

End Function

Public Function octaltostring(ByVal mystring, ByVal maxlength) As String
          Dim binarystring As String
          Dim place        As Integer
          Dim Letter       As String
          Dim my_string    As String
          Dim X            As Integer
          Dim Y            As Byte
          Dim total        As Integer
10        place = 64

20        For X = 1 To Len(mystring) Step 3

30            binarystring = mid$(mystring, X, 3)

40            For Y = 1 To 3

50                total = total + mid$(binarystring, Y, 1) * place
60                place = place / 8
70            Next Y

80            place = 64
90            my_string = my_string & Chr(total)
100           total = 0
110       Next X

120       octaltostring = my_string

End Function

Public Function stringtohex(ByVal mystring, ByVal maxlength) As String
          Dim TEMPASCII     As Integer
          Dim tempbit       As Integer
          Dim binary_string As String
          Dim Filter        As Integer
          Dim Letter(6)     As String
          Dim hexletter     As Integer
          Dim X             As Integer
          Dim Y             As Byte
10        Letter(0) = "A"
20        Letter(1) = "B"
30        Letter(2) = "C"
40        Letter(3) = "D"
50        Letter(4) = "E"
60        Letter(5) = "F"

70        For X = 1 To maxlength

80            Filter = 15
90            TEMPASCII = Asc(mid$(mystring, X, 1))

100           For Y = 1 To 2

110               tempbit = TEMPASCII And Filter
120               hexletter = (15 * tempbit / Filter)

130               If hexletter >= 10 Then

140                   binary_string = Letter(hexletter - 10) & binary_string
150               Else
160                   binary_string = hexletter & binary_string

170               End If

180               Filter = Filter * 16
190           Next Y

200           stringtohex = stringtohex & binary_string
210           binary_string = ""
220       Next X

End Function

Public Function hextostring(ByVal mystring, ByVal maxlength) As String
          Dim binarystring As String
          Dim place        As Integer
          Dim Letter       As String
          Dim my_string    As String
          Dim total        As Integer
          Dim value        As Integer
          Dim X            As Integer
          Dim Y            As Byte
10        place = 16

20        For X = 1 To Len(mystring) Step 2

30            binarystring = mid$(mystring, X, 2)

40            For Y = 1 To 2

50                Select Case mid$(binarystring, Y, 1)

                      Case "A"

60                        value = 10

70                    Case "B"

80                        value = 11

90                    Case "C"

100                       value = 12

110                   Case "D"

120                       value = 13

130                   Case "E"

140                       value = 14

150                   Case "F"

160                       value = 15

170                   Case Else

180                       value = Val(mid$(binarystring, Y, 1))

190               End Select

200               total = total + value * place
210               place = place / 16
220           Next Y

230           place = 16
240           my_string = my_string & Chr(total)
250           total = 0
260       Next X

270       hextostring = my_string

End Function

Public Function rev(ByVal mybinary)
          Dim X    As Byte
          Dim a    As Integer
          Dim temp As Long

10        For X = 1 To 8

20            a = mid$(mybinary, X, 1)

30            If a = 1 Then

40                a = 0
50            Else
60                a = 1

70            End If

80            temp = temp & a
90        Next X

100       rev = temp

End Function


