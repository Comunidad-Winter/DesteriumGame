VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If Wgl = 0 Then
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Desterium AO is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As DirectDrawSurface7
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private DirectDraw As DirectDraw7

Private maxBytesToUse As Long
Private usedBytes As Long

Private useVideoMemory As Boolean

Private ResourcePath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function SetDIBitsToDevice Lib "GDI32.dll" (ByVal hdc As Long, _
    ByVal XDest As Long, ByVal YDest As Long, ByVal dwWidth As Long, ByVal dwHeight _
    As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal uStartScan As Long, _
    ByVal CScanLine As Long, lpBits As Any, lpbmi As BITMAPINFO, ByVal fuColorUse _
    As Long) As Long

Private Sub Class_Initialize()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      '
      '**************************************************************
10        usedBytes = 0
20        maxBytesToUse = MIN_MEMORY_TO_USE
End Sub

Private Sub Class_Terminate()
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      'Clean up
      '**************************************************************
          Dim i As Long
          Dim j As Long
          
          'Destroy every surface in memory
10        For i = 0 To HASH_TABLE_SIZE - 1
20            With surfaceList(i)
30                For j = 1 To .surfaceCount
40                    Set .SurfaceEntry(j).Surface = Nothing
50                Next j
                  
                  'Destroy the arrays
60                Erase .SurfaceEntry
70            End With
80        Next i
End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, ByVal _
    videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal _
    maxMemoryUsageInMb As Long = -1)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      'Initializes the manager
      '**************************************************************
10        Set DirectDraw = DD
          
20        useVideoMemory = videoMemory
          
30        ResourcePath = graphicPath
          
40        If maxMemoryUsageInMb = -1 Then
50            maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
60        ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
70            maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
80        Else
90            maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
100       End If
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As _
    DirectDrawSurface7
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      'Retrieves the requested texture
      '**************************************************************
          Dim i As Long
          
          ' Search the index on the list
10        With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
20            For i = 1 To .surfaceCount
30                If .SurfaceEntry(i).fileIndex = fileIndex Then
40                    .SurfaceEntry(i).lastAccess = GetTickCount
50                    Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
60                    Exit Property
70                End If
80            Next i
90        End With
          
          'Not in memory, load it!
100       Set clsSurfaceManager_Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As DirectDrawSurface7
      '**************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 05/11/2009
      'Loads the surface named fileIndex + ".bmp" and inserts it to the
      'surface list in the listIndex position
      '**************************************************************
10    On Error GoTo ErrHandler

          Dim newSurface As SURFACE_ENTRY_DYN
          Dim ddsd As DDSURFACEDESC2
          Dim ddck As DDCOLORKEY
          Dim data() As Byte
          Dim bmpInfo As BITMAPINFO
          Dim sDC As Long
          
          'get Bitmap
20        Call Get_Bitmap(ResourcePath, CStr(fileIndex) & ".BMP", bmpInfo, data)

          'Set up the surface desc
30        ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
40        If useVideoMemory Then
50            ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
60        Else
70            ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
80        End If
90        ddsd.lWidth = bmpInfo.bmiHeader.biWidth
100       ddsd.lHeight = bmpInfo.bmiHeader.biHeight
              
110       With newSurface
120           .fileIndex = fileIndex
              
              'Set last access time (if we didn't we would reckon this texture as the one lru)
130           .lastAccess = GetTickCount
              
              'Load surface
140           Set .Surface = DirectDraw.CreateSurface(ddsd)
              
150           sDC = .Surface.GetDC
              
160           Call SetDIBitsToDevice(sDC, 0, 0, bmpInfo.bmiHeader.biWidth, _
                  bmpInfo.bmiHeader.biHeight, 0, 0, 0, bmpInfo.bmiHeader.biHeight, _
                  data(0), bmpInfo, DIB_RGB_COLORS)
              
170           Call .Surface.ReleaseDC(sDC)
              
              'Set colorkey
180           ddck.high = 0
190           ddck.low = 0
200           Call .Surface.SetColorKey(DDCKEY_SRCBLT, ddck)
              
              'Retrieve the updated surface desc
210           Call .Surface.GetSurfaceDesc(ddsd)
220       End With
          
          'Insert surface to the list
230       With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
240           .surfaceCount = .surfaceCount + 1
              
250           ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
              
260           .SurfaceEntry(.surfaceCount) = newSurface
              
270           Set LoadSurface = newSurface.Surface
280       End With
          
          'Update used bytes
290       usedBytes = usedBytes + ddsd.lHeight * ddsd.lPitch
          
          'Check if we have exceeded our allowed share of memory usage
300       Do While usedBytes > maxBytesToUse
              'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
310           If Not RemoveLRU() Then
320               Exit Do
330           End If
340       Loop
350   Exit Function

ErrHandler:
360       If Err.number = DDERR_OUTOFMEMORY Or Err.number = DDERR_OUTOFVIDEOMEMORY _
              Then
              'Remove a surface and try again
370           If RemoveLRU() Then
380               Resume
390           Else
400               MsgBox _
                      "No hay memoria disponible! El programa abortará. Cierra algunos programas e intenta de nuevo"
410               End
420           End If
430       Else
440           MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & _
                  CStr(fileIndex) & ".BMP" & ". " & vbCrLf & "El código de error es " & _
                  CStr(Err.number) & " - " & Err.Description & vbCrLf & vbCrLf & _
                  "Copia este mensaje y notifica a los administradores.", vbOKOnly Or _
                  vbCritical Or vbExclamation, "Error"
450           End
460       End If
End Function

Private Function RemoveLRU() As Boolean
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      'Removes the Least Recently Used surface to make some room for new ones
      '**************************************************************
          Dim LRUi As Long
          Dim LRUj As Long
          Dim LRUtime As Long
          Dim i As Long
          Dim j As Long
          Dim ddsd As DDSURFACEDESC2
          
10        LRUtime = GetTickCount
          
          'Check out through the whole list for the least recently used
20        For i = 0 To HASH_TABLE_SIZE - 1
30            With surfaceList(i)
40                For j = 1 To .surfaceCount
50                    If LRUtime > .SurfaceEntry(j).lastAccess Then
60                        LRUi = i
70                        LRUj = j
80                        LRUtime = .SurfaceEntry(j).lastAccess
90                    End If
100               Next j
110           End With
120       Next i
          
130       If LRUj Then
140           RemoveLRU = True
              'Retrieve the surface desc
150           Call surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetSurfaceDesc(ddsd)
              
              'Remove it
160           Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
170           surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
              
              'Move back the list (if necessary)
180           With surfaceList(LRUi)
190               For j = LRUj To .surfaceCount - 1
200                   .SurfaceEntry(j) = .SurfaceEntry(j + 1)
210               Next j
                  
220               .surfaceCount = .surfaceCount - 1
230               If .surfaceCount Then
240                   ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As _
                          SURFACE_ENTRY_DYN
250               Else
260                   Erase .SurfaceEntry
270               End If
280           End With
              
              'Update the used bytes
290           usedBytes = usedBytes - ddsd.lHeight * ddsd.lPitch
300       End If
End Function


#End If
