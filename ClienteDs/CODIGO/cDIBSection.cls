VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As _
    Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() _
    As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As _
    Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As _
    Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As _
    Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal _
    hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As _
    Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal _
    hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As _
    Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc _
    As Long) As Long
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal _
    hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, _
    ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Const CF_BITMAP = 2

' Handle to the current DIBSection:
Private m_hDIb As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

Public Function CopyToClipboard(Optional ByVal bAsDIB As Boolean = True) As _
    Boolean
      Dim lhDCDesktop As Long
      Dim lhDC As Long
      Dim lhBmpOld As Long
      Dim hObj As Long
      Dim lFmt As Long

10        lhDCDesktop = GetDC(GetDesktopWindow())
20        If (lhDCDesktop <> 0) Then
30            lhDC = CreateCompatibleDC(lhDCDesktop)
40            If (lhDC <> 0) Then
50                If (bAsDIB) Then
60                   MsgBox _
                         "I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!"
                      ' Create a duplicate DIBSection and copy
                      ' to the clipboard:
                      'LSet tBI = m_tBI
                      'hDibCopy = CreateDIBSection( '        lhDC, '        m_tBI, '        DIB_RGB_COLORS, '        lPtr, '        0, 0)
                      'If (hDibCopy <> 0) Then
                      '    lhBmpOld = SelectObject(lhDC, hObj)
                      '    BitBlt lhDC, 0, 0, Width, Height, m_hDC, 0, 0, vbSrcCopy
                      '    SelectObject lhDC, lhBmpOld
                      '    lFmt = CF_DIB
                      '
                      '     '....
                                          
                      'Else
                      '    hObj = 0
                      'End If
70                Else
                      ' Create a compatible bitmap and copy to
                      ' the clipboard:
80                    hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
90                    If (hObj <> 0) Then
100                       lhBmpOld = SelectObject(lhDC, hObj)
110                       PaintPicture lhDC
120                       SelectObject lhDC, lhBmpOld
130                       lFmt = CF_BITMAP
                          ' Now set the clipboard to the bitmap:
140                       If (OpenClipboard(0) <> 0) Then
150                           EmptyClipboard
160                           If (SetClipboardData(lFmt, hObj) <> 0) Then
170                               CopyToClipboard = True
180                           End If
190                           CloseClipboard
200                       End If
210                   End If
220               End If
230               DeleteDC lhDC
240           End If
250           DeleteDC lhDCDesktop
260       End If
End Function

Public Function CreateDIB(ByVal lhDC As Long, ByVal lWidth As Long, ByVal _
    lHeight As Long, ByRef hDib As Long) As Boolean
10        With m_tBI.bmiHeader
20            .biSize = Len(m_tBI.bmiHeader)
30            .biWidth = lWidth
40            .biHeight = lHeight
50            .biPlanes = 1
60            .biBitCount = 24
70            .biCompression = BI_RGB
80            .biSizeImage = BytesPerScanLine * .biHeight
90        End With
100       hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
110       CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture(ByRef picThis As StdPicture)
      Dim lhDC As Long
      Dim lhDCDesktop As Long
      Dim lhBmpOld As Long
      Dim tBMP As BITMAP
      Dim lhWnd As Long
          
10        GetObjectAPI picThis.handle, Len(tBMP), tBMP
20        If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
30            lhWnd = GetDesktopWindow()
40            lhDCDesktop = GetDC(lhWnd)
50            If (lhDCDesktop <> 0) Then
60                lhDC = CreateCompatibleDC(lhDCDesktop)
70                ReleaseDC lhWnd, lhDCDesktop
80                If (lhDC <> 0) Then
90                    lhBmpOld = SelectObject(lhDC, picThis.handle)
100                   LoadPictureBlt lhDC
110                   SelectObject lhDC, lhBmpOld
120                   DeleteDC lhDC
130               End If
140           End If
150       End If
End Function
Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
10        ClearUp
20        m_hDC = CreateCompatibleDC(0)
30        If (m_hDC <> 0) Then
40            If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
50                m_hBmpOld = SelectObject(m_hDC, m_hDIb)
60                Create = True
70            Else
80                DeleteDC m_hDC
90                m_hDC = 0
100           End If
110       End If
End Function
Public Property Get BytesPerScanLine() As Long
          ' Scans must align on dword boundaries:
10        BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
10        Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
10        Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt(ByVal lhDC As Long, Optional ByVal lSrcLeft As Long = _
    0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, _
    Optional ByVal lSrcHeight As Long = -1, Optional ByVal eRop As _
    RasterOpConstants = vbSrcCopy)
10        If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
20        If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
30        BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub


Public Sub PaintPicture(ByVal lhDC As Long, Optional ByVal lDestLeft As Long = _
    0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, _
    Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, _
    Optional ByVal lSrcTop As Long = 0, Optional ByVal eRop As RasterOpConstants = _
    vbSrcCopy)
10        If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
20        If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
30        BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, _
              lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
10        hdc = m_hDC
End Property
Public Property Get hDib() As Long
10        hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
10        DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits(Optional ByVal bGray As Boolean = False)
      Dim bDib() As Byte
      Dim X As Long, Y As Long
      Dim lC As Long
      Dim tSA As SAFEARRAY2D
      Dim xEnd As Long
          
          ' Get the bits in the from DIB section:
10        With tSA
20            .cbElements = 1
30            .cDims = 2
40            .Bounds(0).lLbound = 0
50            .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
60            .Bounds(1).lLbound = 0
70            .Bounds(1).cElements = BytesPerScanLine()
80            .pvData = m_lPtr
90        End With
100       CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

          ' random:
110       Randomize Timer
          
120       xEnd = (Width - 1) * 3
130       If (bGray) Then
140           For Y = 0 To m_tBI.bmiHeader.biHeight - 1
150               For X = 0 To xEnd Step 3
160                   lC = Rnd * 255
170                   bDib(X, Y) = lC
180                   bDib(X + 1, Y) = lC
190                   bDib(X + 2, Y) = lC
200               Next X
210           Next Y
220       Else
230           For X = 0 To xEnd Step 3
240               For Y = 0 To m_tBI.bmiHeader.biHeight - 1
250                   bDib(X, Y) = 0
260                   bDib(X + 1, Y) = Rnd * 255
270                   bDib(X + 2, Y) = Rnd * 255
280               Next Y
290           Next X
300       End If
          
          ' Clear the temporary array descriptor
          ' (This does not appear to be necessary, but
          ' for safety do it anyway)
310       CopyMemory ByVal VarPtrArray(bDib), 0&, 4
          
End Sub

Public Sub ClearUp()
10        If (m_hDC <> 0) Then
20            If (m_hDIb <> 0) Then
30                SelectObject m_hDC, m_hBmpOld
40                DeleteObject m_hDIb
50            End If
60            DeleteDC m_hDC
70        End If
80        m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Function Resample(ByVal lNewHeight As Long, ByVal lNewWidth As Long) As _
    cDIBSection
      Dim cDib As cDIBSection
10        Set cDib = New cDIBSection
20        If cDib.Create(lNewWidth, lNewHeight) Then
30            If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> _
                  m_tBI.bmiHeader.biHeight) Then
                  ' Change in size, do resample:
40                ResampleDib cDib
50            Else
                  ' No size change so just return a copy:
60                cDib.LoadPictureBlt m_hDC
70            End If
80            Set Resample = cDib
90        End If
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
      Dim bDibFrom() As Byte
      Dim bDibTo() As Byte

      Dim tSAFrom As SAFEARRAY2D
      Dim tSATo As SAFEARRAY2D

          ' Get the bits in the from DIB section:
10        With tSAFrom
20            .cbElements = 1
30            .cDims = 2
40            .Bounds(0).lLbound = 0
50            .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
60            .Bounds(1).lLbound = 0
70            .Bounds(1).cElements = BytesPerScanLine()
80            .pvData = m_lPtr
90        End With
100       CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

          ' Get the bits in the to DIB section:
110       With tSATo
120           .cbElements = 1
130           .cDims = 2
140           .Bounds(0).lLbound = 0
150           .Bounds(0).cElements = cDibTo.Height
160           .Bounds(1).lLbound = 0
170           .Bounds(1).cElements = cDibTo.BytesPerScanLine()
180           .pvData = cDibTo.DIBSectionBitsPtr
190       End With
200       CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

      Dim xScale As Single
      Dim yScale As Single

      Dim X As Long, Y As Long, xEnd As Long, xOut As Long

      Dim fX As Single, fY As Single
      Dim ifY As Long, ifX As Long
      Dim dX As Single, dy As Single
      Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
      Dim g As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
      Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
      Dim ir1 As Long, ig1 As Long, ib1 As Long
      Dim ir2 As Long, ig2 As Long, ib2 As Long

210       xScale = (Width - 1) / cDibTo.Width
220       yScale = (Height - 1) / cDibTo.Height
          
230       xEnd = cDibTo.Width - 1
              
240       For Y = 0 To cDibTo.Height - 1
              
250           fY = Y * yScale
260           ifY = Int(fY)
270           dy = fY - ifY
              
280           For X = 0 To xEnd
290               fX = X * xScale
300               ifX = Int(fX)
310               dX = fX - ifX
                  
320               ifX = ifX * 3
                  ' Interpolate using the four nearest pixels in the source
330               b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): r1 = _
                      bDibFrom(ifX + 2, ifY)
340               b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = _
                      bDibFrom(ifX + 5, ifY)
350               b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = _
                      bDibFrom(ifX + 2, ifY + 1)
360               b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): _
                      r4 = bDibFrom(ifX + 5, ifY + 1)
                  
                  ' Interplate in x direction:
370               ir1 = r1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = _
                      b1 * (1 - dy) + b3 * dy
380               ir2 = r2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = _
                      b2 * (1 - dy) + b4 * dy
                  ' Interpolate in y:
390               r = ir1 * (1 - dX) + ir2 * dX: g = ig1 * (1 - dX) + ig2 * dX: b = _
                      ib1 * (1 - dX) + ib2 * dX
                  
                  ' Set output:
400               If (r < 0) Then r = 0
410               If (r > 255) Then r = 255
420               If (g < 0) Then g = 0
430               If (g > 255) Then g = 255
440               If (b < 0) Then b = 0
450               If (b > 255) Then
460                   b = 255
470               End If
480               xOut = X * 3
490               bDibTo(xOut, Y) = b
500               bDibTo(xOut + 1, Y) = g
510               bDibTo(xOut + 2, Y) = r
                  
520           Next X
              
530       Next Y

          ' Clear the temporary array descriptor
          ' (This does not appear to be necessary, but
          ' for safety do it anyway)
540       CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
550       CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function

Private Sub Class_Terminate()
10        ClearUp
End Sub




