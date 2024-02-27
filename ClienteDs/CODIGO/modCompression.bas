Attribute VB_Name = "modCompression"
Option Explicit

Public Const GRH_SOURCE_FILE_EXT As String = ".bmp"
Public Const GRH_RESOURCE_FILE As String = "Graficos.AO"
Public Const GRH_PATCH_FILE As String = "Graficos.PATCH"
Public Const MAPS_SOURCE_FILE_EXT As String = ".map"
Public Const MAPS_RESOURCE_FILE As String = "Maps.AO"
Public Const MAPS_PATCH_FILE As String = "Mapas.PATCH"

Public GrhDatContra() As Byte ' Contraseña
Public GrhUsaContra As Boolean ' Usa Contraseña?

Public MapsDatContra() As Byte ' Contraseña
Public MapsUsaContra As Boolean  ' Usa Contraseña?

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER
    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, _
    Src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, _
    Src As Any, ByVal srclen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest _
    As Any, ByRef source As Any, ByVal byteCount As Long)

'BitMaps Strucures
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0
Private Const BI_RLE8 As Long = 1
Private Const BI_RLE4 As Long = 2
Private Const BI_BITFIELDS As Long = 3
Private Const BI_JPG As Long = 4
Private Const BI_PNG As Long = 5

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As _
    Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, _
    ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As _
    Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal _
    lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
    "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As _
    Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

Public Sub GenerateContra(ByVal Contra As String, Optional Modo As Byte = 0)
      '***************************************************
      'Author: ^[GS]^
      'Last Modification: 17/06/2012 - ^[GS]^
      '
      '***************************************************

10    On Error Resume Next

          Dim LoopC As Byte
20        If Modo = 0 Then
30            Erase GrhDatContra
40        ElseIf Modo = 1 Then
50            Erase MapsDatContra
60        End If
          
70        If LenB(Contra) <> 0 Then
80            If Modo = 0 Then
90                ReDim GrhDatContra(Len(Contra) - 1)
100               For LoopC = 0 To UBound(GrhDatContra)
110                   GrhDatContra(LoopC) = Asc(mid(Contra, LoopC + 1, 1))
120               Next LoopC
130               GrhUsaContra = True
140           ElseIf Modo = 1 Then
150               ReDim MapsDatContra(Len(Contra) - 1)
160               For LoopC = 0 To UBound(MapsDatContra)
170                   MapsDatContra(LoopC) = Asc(mid(Contra, LoopC + 1, 1))
180               Next LoopC
190               MapsUsaContra = True
200           End If
210       Else
220           If Modo = 0 Then
230               GrhUsaContra = False
240           ElseIf Modo = 1 Then
250               MapsUsaContra = False
260           End If
270       End If
          
End Sub

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As _
    Currency
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 6/07/2004
      '
      '**************************************************************
          Dim retval As Long
          Dim FB As Currency
          Dim BT As Currency
          Dim FBT As Currency
          
10        retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
          
20        General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal first As _
    Long, ByVal last As Long)
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 08/20/2007
      'Sorts the info headers by their file name using QuickSort.
      '*****************************************************************
          Dim aux As INFOHEADER
          Dim min As Long
          Dim max As Long
          Dim comp As String
          
10        min = first
20        max = last
          
30        comp = InfoHead((min + max) \ 2).strFileName
          
40        Do While min <= max
50            Do While InfoHead(min).strFileName < comp And min < last
60                min = min + 1
70            Loop
80            Do While InfoHead(max).strFileName > comp And max > first
90                max = max - 1
100           Loop
110           If min <= max Then
120               aux = InfoHead(min)
130               InfoHead(min) = InfoHead(max)
140               InfoHead(max) = aux
150               min = min + 1
160               max = max - 1
170           End If
180       Loop
          
190       If first < max Then Call Sort_Info_Headers(InfoHead, first, max)
200       If min < last Then Call Sort_Info_Headers(InfoHead, min, last)
End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As _
    INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal _
    FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 08/21/2007
      'Searches for the specified InfoHeader
      '*****************************************************************
          Dim ReadingHead As Long
          Dim ReadInfoHead As INFOHEADER
          
10        Do Until FirstHead > LastHead
20            ReadingHead = (FirstHead + LastHead) \ 2

30            Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + _
                  1, ReadInfoHead

40            If InfoHead.strFileName = ReadInfoHead.strFileName Then
50                InfoHead = ReadInfoHead
60                BinarySearch = True
70                Exit Function
80            Else
90                If InfoHead.strFileName < ReadInfoHead.strFileName Then
100                   LastHead = ReadingHead - 1
110               Else
120                   FirstHead = ReadingHead + 1
130               End If
140           End If
150       Loop
End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, ByRef FileName As _
    String, ByRef InfoHead As INFOHEADER, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 16/07/2012 - ^[GS]^
      'Retrieves the InfoHead of the specified graphic file
      '*****************************************************************
          Dim ResourceFile As Integer
          Dim ResourceFilePath As String
          Dim FileHead As FILEHEADER
          
10    On Local Error GoTo ErrHandler

20        If Modo = 0 Then
30            ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
40        ElseIf Modo = 1 Then
50            ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
60        End If
          
          'Set InfoHeader we are looking for
70        InfoHead.strFileName = UCase$(FileName)
         
          'Open the binary file
80        ResourceFile = FreeFile()
90        Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
              'Extract the FILEHEADER
100           Get ResourceFile, 1, FileHead
              
              'Check the file for validity
110           If LOF(ResourceFile) <> FileHead.lngFileSize Then
120               MsgBox "Archivo de recursos dañado. " & ResourceFilePath, , "Error"
130               Close ResourceFile
140               Exit Function
150           End If
              
              'Search for it!
160           If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, _
                  Len(FileHead), Len(InfoHead)) Then
170               Get_InfoHeader = True
180           End If
              
190       Close ResourceFile
200   Exit Function

ErrHandler:
210       Close ResourceFile
          
220       Call MsgBox("Error al intentar leer el archivo " & ResourceFilePath & _
              ". Razón: " & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef data() As Byte, Optional Modo As Byte = 0)
      '*****************************************************************
      'Author: Juan Martín Dotuyo Dodero
      'Last Modify Date: 17/07/2012 - ^[GS]^
      'Compresses binary data avoiding data loses
      '*****************************************************************
          Dim Dimensions As Long
          Dim DimBuffer As Long
          Dim BufTemp() As Byte
          Dim LoopC As Long
          
10        Dimensions = UBound(data) + 1
          
          ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
20        DimBuffer = Dimensions * 1.06
          
30        ReDim BufTemp(DimBuffer)
          
40        Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
          
50        Erase data
          
60        ReDim data(DimBuffer - 1)
70        ReDim Preserve BufTemp(DimBuffer - 1)
          
80        data = BufTemp
          
90        Erase BufTemp
          
          ' GSZAO - Seguridad
100       If Modo = 0 And GrhUsaContra = True Then
110           If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 _
                  Then
120               For LoopC = 0 To UBound(GrhDatContra)
130                   data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
140               Next LoopC
150           End If
160       ElseIf Modo = 1 And MapsUsaContra = True Then
170           If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 _
                  Then
180               For LoopC = 0 To UBound(MapsDatContra)
190                   data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
200               Next LoopC
210           End If
220       End If
          ' GSZAO - Seguridad
          
End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long, _
    Optional Modo As Byte = 0)
      '*****************************************************************
      'Author: Juan Martín Dotuyo Dodero
      'Last Modify Date: 16/07/2012 - ^[GS]^
      'Decompresses binary data
      '*****************************************************************
          Dim BufTemp() As Byte
          Dim LoopC As Integer
          
10        ReDim BufTemp(OrigSize - 1)
          
          ' GSZAO - Seguridad
20        If Modo = 0 And GrhUsaContra = True Then
30            If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 _
                  Then
40                For LoopC = 0 To UBound(GrhDatContra)
50                    data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
60                Next LoopC
70            End If
80        ElseIf Modo = 1 And MapsUsaContra = True Then
90            If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 _
                  Then
100               For LoopC = 0 To UBound(MapsDatContra)
110                   data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
120               Next LoopC
130           End If
140       End If
          ' GSZAO - Seguridad
          
150       Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
          
160       ReDim data(OrigSize - 1)
          
170       data = BufTemp
          
180       Erase BufTemp
End Sub

''
' Compresses all graphic files to a resource file.
'
' @param    SourcePath The graphic files folder.
' @param    OutputPath The resource file folder.
' @param    version The resource file version.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Compress_Files(ByRef SourcePath As String, ByRef OutputPath As _
    String, ByVal version As Long, ByRef prgBar As ProgressBar, Optional Modo As _
    Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 17/07/2012 - ^[GS]^
      'Compresses all graphic files to a resource file
      '*****************************************************************
          Dim SourceFileName As String
          Dim OutputFilePath As String
          Dim SourceFile As Long
          Dim OutputFile As Long
          Dim SourceData() As Byte
          Dim FileHead As FILEHEADER
          Dim InfoHead() As INFOHEADER
          Dim LoopC As Long

10    On Local Error GoTo ErrHandler
20        If Modo = 0 Then
30            OutputFilePath = OutputPath & GRH_RESOURCE_FILE
40            SourceFileName = Dir(SourcePath & "*" & GRH_SOURCE_FILE_EXT, vbNormal)
50        ElseIf Modo = 1 Then
60            OutputFilePath = OutputPath & MAPS_RESOURCE_FILE
70            SourceFileName = Dir(SourcePath & "*" & MAPS_SOURCE_FILE_EXT, vbNormal)
80        End If
          
          ' Create list of all files to be compressed
90        While LenB(SourceFileName) <> 0
100           FileHead.lngNumFiles = FileHead.lngNumFiles + 1
              
110           ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
120           InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
              
              'Search new file
130           SourceFileName = Dir()
140       Wend
          
150       If FileHead.lngNumFiles = 0 Then
160           MsgBox "No se encontraron archivos de extensión " & GRH_SOURCE_FILE_EXT _
                  & " en " & SourcePath & ".", , "Error"
170           Exit Function
180       End If
          
190       If Not prgBar Is Nothing Then
200           prgBar.value = 0
210           prgBar.max = FileHead.lngNumFiles + 1
220       End If
          
          'Destroy file if it previuosly existed
230       If LenB(Dir(OutputFilePath, vbNormal)) <> 0 Then
240           Kill OutputFilePath
250       End If
          
          'Finish setting the FileHeader data
260       FileHead.lngFileVersion = version
270       FileHead.lngFileSize = Len(FileHead) + FileHead.lngNumFiles * _
              Len(InfoHead(0))
          
          'Order the InfoHeads
280       Call Sort_Info_Headers(InfoHead(), 0, FileHead.lngNumFiles - 1)
          
          'Open a new file
290       OutputFile = FreeFile()
300       Open OutputFilePath For Binary Access Read Write As OutputFile
              ' Move to the end of the headers, where the file data will actually start
310           Seek OutputFile, FileHead.lngFileSize + 1
              
              ' Process every file!
320           For LoopC = 0 To FileHead.lngNumFiles - 1
                    
330               SourceFile = FreeFile()
340               Open SourcePath & InfoHead(LoopC).strFileName For Binary Access _
                      Read Lock Write As SourceFile
                      
                      'Find out how large the file is and resize the data array appropriately
350                   InfoHead(LoopC).lngFileSizeUncompressed = LOF(SourceFile)
360                   ReDim SourceData(LOF(SourceFile) - 1)
                      
                      'Get the data from the file
370                   Get SourceFile, , SourceData
                      
                      'Compress it
380                   Call Compress_Data(SourceData, Modo)
                      
                      'Store it in the resource file
390                   Put OutputFile, , SourceData
                      
400                   With InfoHead(LoopC)
                          'Set up the info headers
410                       .lngFileSize = UBound(SourceData) + 1
420                       .lngFileStart = FileHead.lngFileSize + 1
                          
                          'Update the file header
430                       FileHead.lngFileSize = FileHead.lngFileSize + .lngFileSize
440                   End With
                      
450                   Erase SourceData
                  
460               Close SourceFile
              
                  'Update progress bar
470               If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
480               DoEvents
490           Next LoopC
              
              'Store the headers in the file
500           Seek OutputFile, 1
510           Put OutputFile, , FileHead
520           Put OutputFile, , InfoHead
              
          'Close the file
530       Close OutputFile
          
540       Erase InfoHead
550       Erase SourceData
          
560       Compress_Files = True
570   Exit Function

ErrHandler:
580       Erase SourceData
590       Erase InfoHead
600       Close OutputFile
          
610       Call MsgBox("No se pudo crear el archivo binario. Razón: " & Err.number & _
              " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, ByRef InfoHead _
    As INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 16/07/2012 - ^[GS]^
      'Retrieves a byte array with the compressed data from the specified file
      '*****************************************************************
          Dim ResourceFilePath As String
          Dim ResourceFile As Integer
          
10    On Local Error GoTo ErrHandler
20        If Modo = 0 Then
30            ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
40        ElseIf Modo = 1 Then
50            ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
60        End If
          
          'Size the Data array
70        ReDim data(InfoHead.lngFileSize - 1)
          
          'Open the binary file
80        ResourceFile = FreeFile
90        Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
              'Get the data
100           Get ResourceFile, InfoHead.lngFileStart, data
          'Close the binary file
110       Close ResourceFile
          
120       Get_File_RawData = True
130   Exit Function

ErrHandler:
140       Close ResourceFile
End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As _
    INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 16/07/2012 - ^[GS]^
      'Extract the specific file from a resource file
      '*****************************************************************
10    On Local Error GoTo ErrHandler
          
20        If Get_File_RawData(ResourcePath, InfoHead, data, Modo) Then
              'Decompress all data
30            If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then
40                Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed, Modo)
50            End If
              
60            Extract_File = True
70        End If
80    Exit Function

ErrHandler:
90        Call MsgBox("Error al intentar decodificar recursos. Razón: " & Err.number _
              & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Extracts all files from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    OutputPath The folder where graphic files will be extracted.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Extract_Files(ByRef ResourcePath As String, ByRef OutputPath As _
    String, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 17/07/2012 - ^[GS]^
      'Extracts all files from a resource file
      '*****************************************************************
          Dim LoopC As Long
          Dim ResourceFile As Integer
          Dim ResourceFilePath As String
          Dim OutputFile As Integer
          Dim SourceData() As Byte
          Dim FileHead As FILEHEADER
          Dim InfoHead() As INFOHEADER
          Dim RequiredSpace As Currency
          
10    On Local Error GoTo ErrHandler
20        If Modo = 0 Then
30            ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
40        ElseIf Modo = 1 Then
50            ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
60        End If
          
          'Open the binary file
70        ResourceFile = FreeFile()
80        Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
              'Extract the FILEHEADER
90            Get ResourceFile, 1, FileHead
          
              'Check the file for validity
100           If LOF(ResourceFile) <> FileHead.lngFileSize Then
110               Call MsgBox("Archivo de recursos dañado. " & ResourceFilePath, , _
                      "Error")
120               Close ResourceFile
130               Exit Function
140           End If
              
              'Size the InfoHead array
150           ReDim InfoHead(FileHead.lngNumFiles - 1)
              
              'Extract the INFOHEADER
160           Get ResourceFile, , InfoHead
              
              'Check if there is enough hard drive space to extract all files
170           For LoopC = 0 To UBound(InfoHead)
                  
180               RequiredSpace = RequiredSpace + _
                      InfoHead(LoopC).lngFileSizeUncompressed
190           Next LoopC
              
200           If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) _
                  Then
210               Erase InfoHead
220               Close ResourceFile
230               Call _
                      MsgBox("No hay suficiente espacio en el disco para extraer los archivos.", _
                      , "Error")
240               Exit Function
250           End If
260       Close ResourceFile
          
          'Update progress bar
270       If Not prgBar Is Nothing Then
280           prgBar.value = 0
290           prgBar.max = FileHead.lngNumFiles + 1
300       End If
          
          'Extract all of the files from the binary file
310       For LoopC = 0 To UBound(InfoHead)
              'Extract this file
320           If Extract_File(ResourcePath, InfoHead(LoopC), SourceData) Then
                  'Destroy file if it previuosly existed
330               If FileExist(OutputPath & InfoHead(LoopC).strFileName, vbNormal) _
                      Then
340                   Call Kill(OutputPath & InfoHead(LoopC).strFileName)
350               End If
                  
                  'Save it!
360               OutputFile = FreeFile()
370               Open OutputPath & InfoHead(LoopC).strFileName For Binary As _
                      OutputFile
380                   Put OutputFile, , SourceData
390               Close OutputFile
                  
400               Erase SourceData
410           Else
420               Erase SourceData
430               Erase InfoHead
                  
440               Call MsgBox("No se pudo extraer el archivo " & _
                      InfoHead(LoopC).strFileName, vbOKOnly, "Error")
450               Exit Function
460           End If
                  
              'Update progress bar
470           If Not prgBar Is Nothing Then prgBar.value = prgBar.value + 1
480           DoEvents
490       Next LoopC
          
500       Erase InfoHead
510       Extract_Files = True
520   Exit Function

ErrHandler:
530       Close ResourceFile
540       Erase SourceData
550       Erase InfoHead
          
560       Call MsgBox("No se pudo extraer el archivo binario correctamente. Razón: " _
              & Err.number & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef ResourcePath As String, ByRef FileName As _
    String, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 16/07/2012 - ^[GS]^
      'Retrieves a byte array with the specified file data
      '*****************************************************************
          Dim InfoHead As INFOHEADER
          
10        If Get_InfoHeader(ResourcePath, FileName, InfoHead, Modo) Then
              'Extract!
20            Get_File_Data = Extract_File(ResourcePath, InfoHead, data, Modo)
30        Else
40            Call MsgBox("No se se encontro el recurso " & FileName)
50        End If
End Function

''
' Retrieves bitmap file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Bitmap(ByRef ResourcePath As String, ByRef FileName As _
    String, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 11/30/2007
      'Retrieves bitmap file data
      '*****************************************************************
          Dim InfoHead As INFOHEADER
          Dim rawData() As Byte
          Dim offBits As Long
          Dim bitmapSize As Long
          Dim colorCount As Long
          
10        If Get_InfoHeader(ResourcePath, FileName, InfoHead) Then
              'Extract the file and create the bitmap data from it.
20            If Extract_File(ResourcePath, InfoHead, rawData) Then
30                Call CopyMemory(offBits, rawData(10), 4)
40                Call CopyMemory(bmpInfo.bmiHeader, rawData(14), 40)
                  
50                With bmpInfo.bmiHeader
60                    bitmapSize = AlignScan(.biWidth, .biBitCount) * Abs(.biHeight)
                      
70                    If .biBitCount < 24 Or .biCompression = BI_BITFIELDS Or _
                          (.biCompression <> BI_RGB And .biBitCount = 32) Then
80                        If .biClrUsed < 1 Then
90                            colorCount = 2 ^ .biBitCount
100                       Else
110                           colorCount = .biClrUsed
120                       End If
                          
                          ' When using bitfields on 16 or 32 bits images, bmiColors has a 3-longs mask.
130                       If .biBitCount >= 16 And .biCompression = BI_BITFIELDS Then _
                              colorCount = 3
                          
140                       Call CopyMemory(bmpInfo.bmiColors(0), rawData(54), _
                              colorCount * 4)
150                   End If
160               End With
                  
170               ReDim data(bitmapSize - 1) As Byte
180               Call CopyMemory(data(0), rawData(offBits), bitmapSize)
                  
190               Get_Bitmap = True
200           End If
210       Else
220           Call MsgBox("No se encontro el recurso " & FileName)
230       End If
End Function


''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As _
    Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 02/11/2007
      'Compare two byte arrays to detect any difference
      '*****************************************************************
          Dim Length As Long
          Dim act As Long
          
10        Length = UBound(data1) + 1
          
20        If (UBound(data2) + 1) = Length Then
30            While act < Length
40                If data1(act) Xor data2(act) Then Exit Function
                  
50                act = act + 1
60            Wend
              
70            Compare_Datas = True
80        End If
End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, ByRef _
    FileHead As FILEHEADER, ByRef InfoHead As INFOHEADER, ByRef ReadFiles As Long) _
    As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 08/24/2007
      'Reads the next InfoHeader
      '*****************************************************************

10        If ReadFiles < FileHead.lngNumFiles Then
              'Read header
20            Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, _
                  InfoHead
              
              'Update
30            ReadNext_InfoHead = True
40        End If
          
50        ReadFiles = ReadFiles + 1
End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNext_Bitmap(ByRef ResourcePath As String, ByRef ReadFiles As _
    Long, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte, ByRef fileIndex As _
    Long) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 12/02/2007
      'Reads the next InfoHeader
      '*****************************************************************
10    On Error Resume Next

          Dim ResourceFile As Integer
          Dim FileHead As FILEHEADER
          Dim InfoHead As INFOHEADER
          Dim FileName As String
          
20        ResourceFile = FreeFile
30        Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As _
              ResourceFile
40        Get ResourceFile, 1, FileHead
          
50        If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
60            Call Get_Bitmap(ResourcePath, InfoHead.strFileName, bmpInfo, data())
70            FileName = Trim$(InfoHead.strFileName)
80            fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
              
90            GetNext_Bitmap = True
100       End If
          
110       Close ResourceFile
End Function

''
' Compares two resource versions and makes a patch file.
'
' @param    NewResourcePath The actual reource file folder.
' @param    OldResourcePath The previous reource file folder.
' @param    OutputPath The patchs file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.

Public Function Make_Patch(ByRef NewResourcePath As String, ByRef _
    OldResourcePath As String, ByRef OutputPath As String, ByRef prgBar As _
    ProgressBar, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 17/07/2012 - ^[GS]^
      'Compares two resource versions and make a patch file
      '*****************************************************************
          Dim NewResourceFile As Integer
          Dim NewResourceFilePath As String
          Dim NewFileHead As FILEHEADER
          Dim NewInfoHead As INFOHEADER
          Dim NewReadFiles As Long
          Dim NewReadNext As Boolean
          
          Dim OldResourceFile As Integer
          Dim OldResourceFilePath As String
          Dim OldFileHead As FILEHEADER
          Dim OldInfoHead As INFOHEADER
          Dim OldReadFiles As Long
          Dim OldReadNext As Boolean
          
          Dim OutputFile As Integer
          Dim OutputFilePath As String
          Dim data() As Byte
          Dim auxData() As Byte
          Dim Instruction As Byte
          
      'Set up the error handler
10    On Local Error GoTo ErrHandler

20        If Modo = 0 Then
30            NewResourceFilePath = NewResourcePath & GRH_RESOURCE_FILE
40            OldResourceFilePath = OldResourcePath & GRH_RESOURCE_FILE
50            OutputFilePath = OutputPath & GRH_PATCH_FILE
60        ElseIf Modo = 1 Then
70            NewResourceFilePath = NewResourcePath & MAPS_RESOURCE_FILE
80            OldResourceFilePath = OldResourcePath & MAPS_RESOURCE_FILE
90            OutputFilePath = OutputPath & MAPS_PATCH_FILE
100       End If
          
          'Open the old binary file
110       OldResourceFile = FreeFile
120       Open OldResourceFilePath For Binary Access Read Lock Write As _
              OldResourceFile
              
              'Get the old FileHeader
130           Get OldResourceFile, 1, OldFileHead
              'Check the file for validity
140           If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
150               Call MsgBox("Archivo de recursos anterior dañado. " & _
                      OldResourceFilePath, , "Error")
160               Close OldResourceFile
170               Exit Function
180           End If
              
              'Open the new binary file
190           NewResourceFile = FreeFile()
200           Open NewResourceFilePath For Binary Access Read Lock Write As _
                  NewResourceFile
                  
                  'Get the new FileHeader
210               Get NewResourceFile, 1, NewFileHead
                  'Check the file for validity
220               If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
230                   Call MsgBox("Archivo de recursos anterior dañado. " & _
                          NewResourceFilePath, , "Error")
240                   Close NewResourceFile
250                   Close OldResourceFile
260                   Exit Function
270               End If
                  
                  'Destroy file if it previuosly existed
280               If LenB(Dir(OutputFilePath, vbNormal)) <> 0 Then Kill OutputFilePath
                  
                  'Open the patch file
290               OutputFile = FreeFile()
300               Open OutputFilePath For Binary Access Read Write As OutputFile
                      
310                   If Not prgBar Is Nothing Then
320                       prgBar.value = 0
330                       prgBar.max = (OldFileHead.lngNumFiles + _
                              NewFileHead.lngNumFiles) + 1
340                   End If
                      
                      'put previous file version (unencrypted)
350                   Put OutputFile, , OldFileHead.lngFileVersion
                      
                      'Put the new file header
360                   Put OutputFile, , NewFileHead
                      'Try to read old and new first files
370                   If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, _
                          OldReadFiles) And ReadNext_InfoHead(NewResourceFile, _
                          NewFileHead, NewInfoHead, NewReadFiles) Then
                          
                          'Update
380                       prgBar.value = prgBar.value + 2
                          
390                       Do 'Main loop
                              'Comparisons are between encrypted names, for ordering issues
400                           If OldInfoHead.strFileName = NewInfoHead.strFileName _
                                  Then

                                  'Get old file data
410                               Call Get_File_RawData(OldResourcePath, OldInfoHead, _
                                      auxData, Modo)
                                  
                                  'Get new file data
420                               Call Get_File_RawData(NewResourcePath, NewInfoHead, _
                                      data, Modo)
                                  
430                               If Not Compare_Datas(data, auxData) Then
                                      'File was modified
440                                   Instruction = PatchInstruction.Modify_File
450                                   Put OutputFile, , Instruction
                                      
                                      'Write header
460                                   Put OutputFile, , NewInfoHead
                                      
                                      'Write data
470                                   Put OutputFile, , data
480                               End If
                                  
                                  'Read next OldResource
490                               If Not ReadNext_InfoHead(OldResourceFile, _
                                      OldFileHead, OldInfoHead, OldReadFiles) Then
500                                   Exit Do
510                               End If
                                  
                                  'Read next NewResource
520                               If Not ReadNext_InfoHead(NewResourceFile, _
                                      NewFileHead, NewInfoHead, NewReadFiles) Then
                                      'Reread last OldInfoHead
530                                   OldReadFiles = OldReadFiles - 1
540                                   Exit Do
550                               End If
                                  
                                  'Update
560                               If Not prgBar Is Nothing Then prgBar.value = _
                                      prgBar.value + 2
                              
570                           ElseIf OldInfoHead.strFileName < _
                                  NewInfoHead.strFileName Then
                                  
                                  'File was deleted
580                               Instruction = PatchInstruction.Delete_File
590                               Put OutputFile, , Instruction
600                               Put OutputFile, , OldInfoHead
                                  
                                  'Read next OldResource
610                               If Not ReadNext_InfoHead(OldResourceFile, _
                                      OldFileHead, OldInfoHead, OldReadFiles) Then
                                      'Reread last NewInfoHead
620                                   NewReadFiles = NewReadFiles - 1
630                                   Exit Do
640                               End If
                                  
                                  'Update
650                               If Not prgBar Is Nothing Then prgBar.value = _
                                      prgBar.value + 1
                              
660                           Else
                                  
                                  'New file
670                               Instruction = PatchInstruction.Create_File
680                               Put OutputFile, , Instruction
690                               Put OutputFile, , NewInfoHead
                                           
                                  'Get file data
700                               Call Get_File_RawData(NewResourcePath, NewInfoHead, _
                                      data, Modo)
                                  
                                  'Write data
710                               Put OutputFile, , data
                                  
                                  'Read next NewResource
720                               If Not ReadNext_InfoHead(NewResourceFile, _
                                      NewFileHead, NewInfoHead, NewReadFiles) Then
                                      'Reread last OldInfoHead
730                                   OldReadFiles = OldReadFiles - 1
740                                   Exit Do
750                               End If
                                  
                                  'Update
760                               If Not prgBar Is Nothing Then prgBar.value = _
                                      prgBar.value + 1
770                           End If
                              
780                           DoEvents
790                       Loop
                      
800                   Else
                          'if at least one is empty
810                       OldReadFiles = 0
820                       NewReadFiles = 0
830                   End If
                      
                      'Read everything?
840                   While ReadNext_InfoHead(OldResourceFile, OldFileHead, _
                          OldInfoHead, OldReadFiles)
                          'Delete file
850                       Instruction = PatchInstruction.Delete_File
860                       Put OutputFile, , Instruction
870                       Put OutputFile, , OldInfoHead
                          
                          'Update
880                       If Not prgBar Is Nothing Then prgBar.value = prgBar.value + _
                              1
890                       DoEvents
900                   Wend
                      
                      'Read everything?
910                   While ReadNext_InfoHead(NewResourceFile, NewFileHead, _
                          NewInfoHead, NewReadFiles)
                          'Create file
920                       Instruction = PatchInstruction.Create_File
930                       Put OutputFile, , Instruction
940                       Put OutputFile, , NewInfoHead
                          
                          'Get file data
950                       Call Get_File_RawData(NewResourcePath, NewInfoHead, data, _
                              Modo)
                          'Write data
960                       Put OutputFile, , data
                          
                          'Update
970                       If Not prgBar Is Nothing Then prgBar.value = prgBar.value + _
                              1
980                       DoEvents
990                   Wend
                  
                  'Close the patch file
1000              Close OutputFile
              
              'Close the new binary file
1010          Close NewResourceFile
          
          'Close the old binary file
1020      Close OldResourceFile
          
1030      Make_Patch = True
1040  Exit Function

ErrHandler:
1050      Close OutputFile
1060      Close NewResourceFile
1070      Close OldResourceFile
          
1080      Call MsgBox("No se pudo terminar de crear el parche. Razón: " & Err.number _
              & " : " & Err.Description, vbOKOnly, "Error")
End Function

''
' Follows patches instructions to update a resource file.
'
' @param    ResourcePath The reource file folder.
' @param    PatchPath The patch file folder.
' @param    PrgBar The control that shows the process state.
'
' @return   True if no error occurred.
Public Function Apply_Patch(ByRef ResourcePath As String, ByRef PatchPath As _
    String, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 17/07/2012 - ^[GS]^
      'Follows patches instructions to update a resource file
      '*****************************************************************
          Dim ResourceFile As Integer
          Dim ResourceFilePath As String
          Dim FileHead As FILEHEADER
          Dim InfoHead As INFOHEADER
          Dim ResourceReadFiles As Long
          Dim EOResource As Boolean

          Dim PatchFile As Integer
          Dim PatchFilePath As String
          Dim PatchFileHead As FILEHEADER
          Dim PatchInfoHead As INFOHEADER
          Dim Instruction As Byte
          Dim OldResourceVersion As Long

          Dim OutputFile As Integer
          Dim OutputFilePath As String
          Dim data() As Byte
          Dim WrittenFiles As Long
          Dim DataOutputPos As Long

10    On Local Error GoTo ErrHandler

20        If Modo = 0 Then
30            ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
40            PatchFilePath = PatchPath & GRH_PATCH_FILE
50            OutputFilePath = ResourcePath & GRH_RESOURCE_FILE & "tmp"
60        ElseIf Modo = 1 Then
70            ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
80            PatchFilePath = PatchPath & MAPS_PATCH_FILE
90            OutputFilePath = ResourcePath & MAPS_RESOURCE_FILE & "tmp"
100       End If
          
          'Open the old binary file
110       ResourceFile = FreeFile()
120       Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
              
              'Read the old FileHeader
130           Get ResourceFile, , FileHead
              'Check the file for validity
140           If LOF(ResourceFile) <> FileHead.lngFileSize Then
150               Call MsgBox("Archivo de recursos anterior dañado. " & _
                      ResourceFilePath, , "Error")
160               Close ResourceFile
170               Exit Function
180           End If
              
              'Open the patch file
190           PatchFile = FreeFile()
200           Open PatchFilePath For Binary Access Read Lock Write As PatchFile
                  
                  'Get previous file version
210               Get PatchFile, , OldResourceVersion
                  
                  'Check the file version
220               If OldResourceVersion <> FileHead.lngFileVersion Then
230                   Call MsgBox("Incongruencia en versiones.", , "Error")
240                   Close ResourceFile
250                   Close PatchFile
260                   Exit Function
270               End If
                  
                  'Read the new FileHeader
280               Get PatchFile, , PatchFileHead
                  
                  'Destroy file if it previuosly existed
290               If FileExist(OutputFilePath, vbNormal) Then Call _
                      Kill(OutputFilePath)
                  
                  'Open the patch file
300               OutputFile = FreeFile()
310               Open OutputFilePath For Binary Access Read Write As OutputFile
                      
                      'Save the file header
320                   Put OutputFile, , PatchFileHead
        
330                   If Not prgBar Is Nothing Then
340                       prgBar.value = 0
350                       prgBar.max = PatchFileHead.lngNumFiles + 1
360                   End If
                      
                      'Update
370                   DataOutputPos = Len(FileHead) + Len(InfoHead) * _
                          PatchFileHead.lngNumFiles + 1
                      
                      'Process loop
380                   While Loc(PatchFile) < LOF(PatchFile)
                          
                          'Get the instruction
390                       Get PatchFile, , Instruction
                          'Get the InfoHead
400                       Get PatchFile, , PatchInfoHead
                          
410                       Do
420                           EOResource = Not ReadNext_InfoHead(ResourceFile, _
                                  FileHead, InfoHead, ResourceReadFiles)
                              
                              'Comparison is performed among encrypted names for ordering issues
430                           If Not EOResource And InfoHead.strFileName < _
                                  PatchInfoHead.strFileName Then
          
                                  'GetData and update InfoHead
440                               Call Get_File_RawData(ResourcePath, InfoHead, data, _
                                      Modo)
450                               InfoHead.lngFileStart = DataOutputPos
                                                 
                                  'Save file!
460                               Put OutputFile, Len(FileHead) + Len(InfoHead) * _
                                      WrittenFiles + 1, InfoHead
470                               Put OutputFile, DataOutputPos, data
                                  
                                  'Update
480                               DataOutputPos = DataOutputPos + UBound(data) + 1
490                               WrittenFiles = WrittenFiles + 1
500                               If Not prgBar Is Nothing Then prgBar.value = _
                                      WrittenFiles
510                           Else
520                               Exit Do
530                           End If
540                       Loop
                          
550                       Select Case Instruction
                              'Delete
                              Case PatchInstruction.Delete_File
560                               If InfoHead.strFileName <> _
                                      PatchInfoHead.strFileName Then
570                                   Err.Description = _
                                          "Incongruencia en archivos de recurso"
580                                   GoTo ErrHandler
590                               End If
                              
                              'Create
600                           Case PatchInstruction.Create_File
610                               If (InfoHead.strFileName > _
                                      PatchInfoHead.strFileName) Or EOResource Then
                                      
                                      'Get file data
620                                   ReDim data(PatchInfoHead.lngFileSize - 1)
630                                   Get PatchFile, , data
                                      
                                      'Save it
640                                   Put OutputFile, Len(FileHead) + Len(InfoHead) * _
                                          WrittenFiles + 1, PatchInfoHead
650                                   Put OutputFile, DataOutputPos, data
                                      
                                      'Reanalize last Resource InfoHead
660                                   EOResource = False
670                                   ResourceReadFiles = ResourceReadFiles - 1
                                      
                                      'Update
680                                   DataOutputPos = DataOutputPos + UBound(data) + 1
690                                   WrittenFiles = WrittenFiles + 1
700                                   If Not prgBar Is Nothing Then prgBar.value = _
                                          WrittenFiles
710                               Else
720                                   Err.Description = _
                                          "Incongruencia en archivos de recurso"
730                                   GoTo ErrHandler
740                               End If
                              
                              'Modify
750                           Case PatchInstruction.Modify_File
760                               If InfoHead.strFileName = PatchInfoHead.strFileName _
                                      Then
                                  

                                      'Get file data
770                                   ReDim data(PatchInfoHead.lngFileSize - 1)
780                                   Get PatchFile, , data
                                                   
                                      'Save it
790                                   Put OutputFile, Len(FileHead) + Len(InfoHead) * _
                                          WrittenFiles + 1, PatchInfoHead
800                                   Put OutputFile, DataOutputPos, data
                                      
                                      'Update
810                                   DataOutputPos = DataOutputPos + UBound(data) + 1
820                                   WrittenFiles = WrittenFiles + 1
830                                   If Not prgBar Is Nothing Then prgBar.value = _
                                          WrittenFiles
840                               Else
850                                   Err.Description = _
                                          "Incongruencia en archivos de recurso"
860                                   GoTo ErrHandler
870                               End If
880                       End Select
                          
890                       DoEvents
900                   Wend
                      
                      'Read everything?
910                   While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, _
                          ResourceReadFiles)

                          'GetData and update InfoHeader
920                       Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
930                       InfoHead.lngFileStart = DataOutputPos
                          
                          'Save file!
940                       Put OutputFile, Len(FileHead) + Len(InfoHead) * _
                              WrittenFiles + 1, InfoHead
950                       Put OutputFile, DataOutputPos, data
                          
                          'Update
960                       DataOutputPos = DataOutputPos + UBound(data) + 1
970                       WrittenFiles = WrittenFiles + 1
980                       If Not prgBar Is Nothing Then prgBar.value = WrittenFiles
990                       DoEvents
1000                  Wend
                  
                  'Close the patch file
1010              Close OutputFile
              
              'Close the new binary file
1020          Close PatchFile
          
          'Close the old binary file
1030      Close ResourceFile
          
          'Check integrity
1040      If (PatchFileHead.lngNumFiles = WrittenFiles) Then
              'Replace File
1050          Call Kill(ResourceFilePath)
1060          Name OutputFilePath As ResourceFilePath
1070      Else
1080          Err.Description = "Falla al procesar parche"
1090          GoTo ErrHandler
1100      End If
          
1110      Apply_Patch = True
1120  Exit Function

ErrHandler:
1130      Close OutputFile
1140      Close PatchFile
1150      Close ResourceFile
          'Destroy file if created
1160      If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
          
1170      Call MsgBox("No se pudo parchear. Razón: " & Err.number & " : " & _
              Err.Description, vbOKOnly, "Error")
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As _
    Long
      '*****************************************************************
      'Author: Unknown
      'Last Modify Date: Unknown
      '*****************************************************************
10        AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
      '*****************************************************************
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/23/2008
      '
      '*****************************************************************
          Dim ResourceFile As Integer
          Dim FileHead As FILEHEADER
          
10        ResourceFile = FreeFile()
20        Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
              'Extract the FILEHEADER
30            Get ResourceFile, 1, FileHead
              
40        Close ResourceFile
          
50        GetVersion = FileHead.lngFileVersion
End Function

Public Function ArrayToPicture(inArray() As Byte, offset As Long, Size As Long) _
    As IPicture
          
          Dim o_hMem  As Long
          Dim o_lpMem  As Long
          Dim aGUID(0 To 3) As Long
          Dim IIStream As IUnknown
          
10        aGUID(0) = &H7BF80980
20        aGUID(1) = &H101ABF32
30        aGUID(2) = &HAA00BB8B
40        aGUID(3) = &HAB0C3000
          
50        o_hMem = GlobalAlloc(&H2&, Size)
60        If Not o_hMem = 0& Then
70            o_lpMem = GlobalLock(o_hMem)
80            If Not o_lpMem = 0& Then
90                CopyMemory ByVal o_lpMem, inArray(offset), Size
100               Call GlobalUnlock(o_hMem)
110               If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
120                     Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), _
                            ArrayToPicture)
130               End If
140           End If
150       End If
End Function
Public Function Get_Bitmapp(ByRef ResourcePath As String, ByRef FileName As _
    String, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte) As Boolean
      '*****************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 11/30/2007
      'Retrieves bitmap file data
      '*****************************************************************
          Dim InfoHead As INFOHEADER
          Dim offBits As Long
          Dim bitmapSize As Long
          Dim colorCount As Long
          
10        If Get_InfoHeader(ResourcePath, FileName, InfoHead, 0) Then
20            If Extract_File(ResourcePath, InfoHead, data, 0) Then Get_Bitmapp = True
30        Else
40            Call MsgBox("No se encontro el recurso " & FileName)
50        End If
End Function




