VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

#If Wgl = 0 Then
'**************************************************************
' clsSurfaceManStatic.cls - Inherits from clsSurfaceManager. Is designed to load
' surfaces at startup, and never deallocating them.
' This grants high performance can use a lot of RAM.
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

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 541

Private Type SURFACE_ENTRY_STATIC
    fileIndex As Long
    Surface As DirectDrawSurface7
End Type

Private Type HashBucket
    sufraceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_STATIC
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashBucket

Private DirectDraw As DirectDraw7

Private useVideoMemory As Boolean

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
20            For j = 1 To surfaceList(i).sufraceCount
30                Set surfaceList(i).SurfaceEntry(j).Surface = Nothing
40            Next j
              
              'Destroy the array
50            Erase surfaceList(i).SurfaceEntry
60        Next i
End Sub

Private Sub clsSurfaceManager_Initialize(ByRef DD As DirectDraw7, ByVal _
    videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal _
    maxMemoryUsageInMb As Long = -1&)
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      '
      '**************************************************************
10        Set DirectDraw = DD
          
20        useVideoMemory = videoMemory
          
30        Call LoadSurfaces(graphicPath)
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As _
    DirectDrawSurface7
      '**************************************************************
      'Author: Juan Martín Sotuyo Dodero
      'Last Modify Date: 3/06/2006
      'Retrieves the requested texture
      '**************************************************************
On Error GoTo ErrHandler:
          Dim i As Long
          
          ' Search for it and return it
10        With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
20            For i = 1 To .sufraceCount
30                If .SurfaceEntry(i).fileIndex = fileIndex Then
40                    Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
50                    Exit Property
60                End If
70            Next i
80        End With
          
          'We have a missing bmp!
90        MsgBox "No se encuentra el archivo " & CStr(fileIndex) & _
              ".bmp. Reinstale el juego, " & _
              "y si el problema persiste contactese con los adminsitradores", vbOKOnly Or _
              vbCritical Or vbExclamation, "Error"
          
100       End
110   Exit Property

ErrHandler:
120       MsgBox "Un error inesperado ocurrió a intentar cargar el archivo " & _
              CStr(fileIndex) & ".bmp. & vbcrlf" & "Error: " & CStr(Err.number), vbOKOnly _
              Or vbCritical Or vbExclamation, "Error"
          
130       End
End Property

Private Sub LoadSurfaces(ByVal ResourcePath As String)
      '**************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 12/02/2007
      'Loads all surfaces in random order and then sorts them
      '**************************************************************
10    On Error Resume Next
          Dim surfaceIndex As Long
          Dim bmpInfo As BITMAPINFO
          Dim data() As Byte
          Dim i As Long
          
20        While GetNext_Bitmap(ResourcePath, i, bmpInfo, data(), surfaceIndex)
30            Call LoadSurface(bmpInfo, data(), surfaceIndex)
40        Wend
End Sub

Private Sub LoadSurface(ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte, _
    ByVal fileIndex As Long)
      '**************************************************************
      'Author: Nicolas Matias Gonzalez (NIGO)
      'Last Modify Date: 11/30/2007
      'Loads the surface named fileIndex + ".bmp" and inserts it to the
      'surface list in the listIndex position
      '**************************************************************
10    On Error GoTo ErrHandler

          Dim newSurface As SURFACE_ENTRY_STATIC
          Dim ddsd As DDSURFACEDESC2
          Dim ddck As DDCOLORKEY
          Dim sDC As Long
          
          'Set up the surface desc
20        ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
30        If useVideoMemory Then
40            ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
50        Else
60            ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
70        End If
80        ddsd.lWidth = bmpInfo.bmiHeader.biWidth
90        ddsd.lHeight = bmpInfo.bmiHeader.biHeight
          
100       With newSurface
110           .fileIndex = fileIndex
              
              'Load surface
120           Set .Surface = DirectDraw.CreateSurface(ddsd)
130           sDC = .Surface.GetDC
              
140           Call SetDIBitsToDevice(sDC, 0, 0, bmpInfo.bmiHeader.biWidth, _
                  bmpInfo.bmiHeader.biHeight, 0, 0, 0, bmpInfo.bmiHeader.biHeight, _
                  data(0), bmpInfo, DIB_RGB_COLORS)
              
150           Call .Surface.ReleaseDC(sDC)
              
              'Set colorkey
160           ddck.high = 0
170           ddck.low = 0
180           Call .Surface.SetColorKey(DDCKEY_SRCBLT, ddck)
              
              'Retrieve the updated surface desc
190           Call .Surface.GetSurfaceDesc(ddsd)
200       End With
          
          'Store the surface in the given index (it MUST be empty or data will be lost)
210       With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
220           .sufraceCount = .sufraceCount + 1
230           ReDim Preserve .SurfaceEntry(1 To .sufraceCount) As SURFACE_ENTRY_STATIC
              
240           .SurfaceEntry(.sufraceCount) = newSurface
250       End With
260   Exit Sub

ErrHandler:
270       If Err.number = DDERR_OUTOFVIDEOMEMORY Then
              ' No video memory? Try system memory! We want all surfaces loaded!
280           ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
290           Resume
300       End If
          
310       MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & _
              CStr(fileIndex) & ".BMP" & ". " & vbCrLf & "El código de error es " & _
              CStr(Err.number) & " - " & Err.Description & vbCrLf & vbCrLf & _
              "Copia este mensaje y notifica a los administradores.", vbOKOnly Or _
              vbCritical Or vbExclamation, "Error"
320           End
End Sub

 
#End If

