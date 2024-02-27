VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsIniManager.cls - Loads INI files into memory and applies Binary Search to get values at high speed.
' Use it instead of GetVar when reading several values form the same file at once, otherwise it's not usefull.
' Based in the idea of AlejoLP and his clsLeerInis.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Loads a complete INI file into memory and sorts it's data and keys for faster searches.
'It is MUCH faster than GetPrivateProfileStringA if you search for several values within a file, otherwise stick to the API.
'It's particularly usefull when you can keep a whole file in memory such as NPCs.dat'
' Based in the idea of AlejoLP and his clsLeerInis.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.1.0
' @date 20060501

'01/05/2006 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'01/04/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: KeyExists method allows to check for valid section keys.
'
'10/10/2010 - Torres Patricio(Pato)
'   - Add:  AddNode method to add a main node.
'           AddKey method to add key into the refered main node.
'           DumpFile method to dump the data into a file.


Option Explicit
Option Base 0


''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'

Private Type ChildNode
    key As String
    value As String
End Type

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.

Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
Private MainNodes As Long

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      '
      '**************************************************************
          Dim i As Long
          
          'Clean up
10        If MainNodes Then
20            For i = 1 To MainNodes - 1
30                Erase fileData(i).values
40            Next i
              
50            Erase fileData
60        End If
          
70        MainNodes = 0
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue

Public Sub Initialize(ByVal file As String)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 27/07/2006
      'Opens the requested file and loads it's data into memory
      '**************************************************************
          Dim handle As Integer
          Dim Text As String
          Dim Pos As Long
          
          'Prevent memory losses if we are attempting to reload a file....
10        Call Class_Terminate
          
          'Get a free handle and start reading line by line until the end
20        handle = FreeFile
          
30        Open file For Input As handle
          
40        Do Until EOF(handle)
50            Line Input #handle, Text
              
              'Is it null??
60            If Len(Text) Then
                  'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
70                If Left$(Text, 1) = "[" Then
                      'If it has an ending ']' it's a main node, otherwise it's nothing
80                    Pos = InStr(2, Text, "]")
90                    If Pos Then
                          'Add a main node
100                       ReDim Preserve fileData(MainNodes) As MainNode
                          
110                       fileData(MainNodes).Name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                          
120                       MainNodes = MainNodes + 1
130                   End If
140               Else
                      'So it's a value. Check if it has a '=', otherwise it's nothing
150                   Pos = InStr(2, Text, "=")
160                   If Pos Then
                          'Is it under any main node??
170                       If MainNodes Then
180                           With fileData(MainNodes - 1)
                                  'Add it to the main node's value
190                               ReDim Preserve .values(.numValues) As ChildNode
                                  
200                               .values(.numValues).value = Right$(Text, Len(Text) - Pos)
210                               .values(.numValues).key = UCase$(Left$(Text, Pos - 1))
                                  
220                               .numValues = .numValues + 1
230                           End With
240                       End If
250                   End If
260               End If
270           End If
280       Loop
          
290       Close handle
          
          Dim i As Long
          
300       If MainNodes Then
              'Sort main nodes to allow binary search
310           Call SortMainNodes(0, MainNodes - 1)
              
              'Sort values of each node to allow binary search
320           For i = 0 To MainNodes - 1
330               If fileData(i).numValues Then _
                      Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
340           Next i
350       End If
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      'Sorts the list of values in a given MainNode using quicksort,
      'this allows the use of Binary Search for faster searches
      '**************************************************************
          Dim min As Integer      'First item in the list
          Dim max As Integer      'Last item in the list
          Dim comp As String      'Item used to compare
          Dim temp As ChildNode
          
10        min = First
20        max = Last
          
30        With Node
40            comp = .values((min + max) \ 2).key
          
50            Do While min <= max
60                Do While .values(min).key < comp And min < Last
70                    min = min + 1
80                Loop
90                Do While .values(max).key > comp And max > First
100                   max = max - 1
110               Loop
120               If min <= max Then
130                   temp = .values(min)
140                   .values(min) = .values(max)
150                   .values(max) = temp
160                   min = min + 1
170                   max = max - 1
180               End If
190           Loop
200       End With
          
210       If First < max Then SortChildNodes Node, First, max
220       If min < Last Then SortChildNodes Node, min, Last
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      'Sorts the MainNodes list using quicksort,
      'this allows the use of Binary Search for faster searches
      '**************************************************************
          Dim min As Integer      'First item in the list
          Dim max As Integer      'Last item in the list
          Dim comp As String      'Item used to compare
          Dim temp As MainNode
          
10        min = First
20        max = Last
          
30        comp = fileData((min + max) \ 2).Name
          
40        Do While min <= max
50            Do While fileData(min).Name < comp And min < Last
60                min = min + 1
70            Loop
80            Do While fileData(max).Name > comp And max > First
90                max = max - 1
100           Loop
110           If min <= max Then
120               temp = fileData(min)
130               fileData(min) = fileData(max)
140               fileData(max) = temp
150               min = min + 1
160               max = max - 1
170           End If
180       Loop
          
190       If First < max Then SortMainNodes First, max
200       If min < Last Then SortMainNodes min, Last
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValue(ByVal Main As String, ByVal key As String) As String
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      'Returns a value if the key and main node exist, or a nullstring otherwise
      '**************************************************************
          Dim i As Long
          Dim j As Long
          
          'Search for the main node
10        i = FindMain(UCase$(Main))
          
20        If i >= 0 Then
              'If valid, binary search among keys
30            j = FindKey(fileData(i), UCase$(key))
              
              'If we found it we return it
40            If j >= 0 Then GetValue = fileData(i).values(j).value
50        End If
End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Sub ChangeValue(ByVal Main As String, ByVal key As String, ByVal value As String)
      '**************************************************************
      'Author: ZaMa
      'Last Modify Date: 27/05/2009
      'If the key and main node exist, changes the value
      '**************************************************************
          Dim i As Long
          Dim j As Long
          
          'Search for the main node
10        i = FindMain(UCase$(Main))
          
20        If i < 0 Then i = AddNode(Main)
          
          'If valid, binary search among keys
30        j = FindKey(fileData(i), UCase$(key))
          
40        If j < 0 Then
50            Call AddKey(Main, key, value)
60        Else
              'If we found it we change it
70            fileData(i).values(j).value = value
80        End If
End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByRef Node As MainNode, ByVal key As String) As Long
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      'Returns the index of the value which key matches the requested one,
      'or the negation of the position were it should be if not found
      '**************************************************************
          Dim min As Long
          Dim max As Long
          Dim mid As Long
          
10        min = 0
20        max = Node.numValues - 1
          
30        Do While min <= max
40            mid = (min + max) \ 2
              
50            If Node.values(mid).key < key Then
60                min = mid + 1
70            ElseIf Node.values(mid).key > key Then
80                max = mid - 1
90            Else
                  'We found it
100               FindKey = mid
110               Exit Function
120           End If
130       Loop
          
          'Not found, return the negation of the position where it should be
          '(all higher values are to the right of the list and lower values are to the left)
140       FindKey = Not min
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal Name As String) As Long
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 5/01/2006
      'Returns the index of the MainNode which name matches the requested one,
      'or the negation of the position were it should be if not found
      '**************************************************************
          Dim min As Long
          Dim max As Long
          Dim mid As Long
          
10        min = 0
20        max = MainNodes - 1
          
30        Do While min <= max
40            mid = (min + max) \ 2
              
50            If fileData(mid).Name < Name Then
60                min = mid + 1
70            ElseIf fileData(mid).Name > Name Then
80                max = mid - 1
90            Else
                  'We found it
100               FindMain = mid
110               Exit Function
120           End If
130       Loop
          
          'Not found, return the negation of the position where it should be
          '(all higher values are to the right of the list and lower values are to the left)
140       FindMain = Not min
End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal Name As String) As Boolean
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 04/01/2008
      'Returns true of the key exists, false otherwise.
      '**************************************************************
10        KeyExists = FindMain(UCase$(Name)) >= 0
End Function

Private Function AddNode(ByRef Name As String) As Long
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 10/10/2010
      '
      '**************************************************************
      Dim i As Long
      Dim MainPos As Long

10    Name = UCase$(Name)

      'Add a main node
20    ReDim Preserve fileData(MainNodes) As MainNode

30    If MainNodes Then
40        i = MainNodes - 1
50        MainPos = Not FindMain(Name)
          
60        Do While i >= MainPos
70            fileData(i + 1) = fileData(i)
80            i = i - 1
90        Loop
          
100       fileData(MainPos).numValues = 0
110       Erase fileData(MainPos).values
120   End If

130   fileData(MainPos).Name = Name

140   MainNodes = MainNodes + 1

150   AddNode = MainPos
End Function

Private Function AddKey(ByRef Main As String, ByRef key As String, Optional ByRef value As String = vbNullString) As Long
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 10/10/2010
      '
      '**************************************************************
      Dim MainPos As Long
      Dim KeyPos As Long
      Dim i As Long

10    Main = UCase$(Main)
20    key = UCase$(key)

30    MainPos = FindMain(Main)

40    If MainPos < 0 Then MainPos = AddNode(Main)

50    With fileData(MainPos)
          'Add it to the main node's value
60        ReDim Preserve .values(.numValues) As ChildNode
          
70        If .numValues > 0 Then
80            i = .numValues - 1
90            KeyPos = Not FindKey(fileData(MainPos), key)
              
100           Do While i >= KeyPos
110               .values(i + 1) = .values(i)
120               i = i - 1
130           Loop
140       End If
          
150       .values(KeyPos).key = key
160       .values(KeyPos).value = value
          
170       .numValues = .numValues + 1
180   End With

190   AddKey = KeyPos
End Function

Public Sub DumpFile(ByRef file As String)
      '**************************************************************
      'Author: Torres Patricio (Pato)
      'Last Modify Date: 10/10/2010
      '
      '**************************************************************
      Dim hFile As Integer
      Dim i As Long
      Dim j As Long

10    hFile = FreeFile()

20    Open file For Output As hFile

30        For i = 0 To MainNodes - 1
40            With fileData(i)
50                Print #hFile, "[" & .Name & "]"
                  
60                For j = 0 To .numValues - 1
70                    Print #hFile, .values(j).key & "=" & .values(j).value
80                Next j
                  
90                Print #hFile, ""
100           End With
110       Next i
          
120   Close hFile
End Sub








