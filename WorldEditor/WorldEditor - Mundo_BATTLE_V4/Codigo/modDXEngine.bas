Attribute VB_Name = "modDXEngine"
Option Explicit


 
Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private Type tGraphicChar
    Src_X As Integer
    Src_Y As Integer
End Type

Private Type tGraphicFont
    texture_index As Long
    Caracteres(0 To 255) As tGraphicChar 'Ascii Chars
    Char_Size As Byte 'In pixels
End Type

Private Type DXFont
    dFont As D3DXFont
    Size As Integer
End Type


Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'***************************
'Variables
'***************************
'Major DX Objects
Public dX As DirectX8
Public d3d As Direct3D8
Public ddevice As Direct3DDevice8
Public d3dx As D3DX8

Dim d3dpp As D3DPRESENT_PARAMETERS

'Texture Manager for Dinamic Textures
Public DXPool As clsTextureManager

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim FPS As Long 'What the current frame rate is.....

Dim engine_render_started As Boolean

'Graphic Font List
Dim gfont_list() As tGraphicFont
Dim gfont_count As Long

'Font List
Private font_list() As DXFont
Private font_count As Integer


'***************************
'Constants
'***************************
'Engine
Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'PI
Private Const PI As Single = 3.14159265358979

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'Initialization
Public Function DXEngine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean)
'On Error GoTo errhandler
    Dim d3dcaps As D3DCAPS8
    Dim d3ddm As D3DDISPLAYMODE
    
    DXEngine_Initialize = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dX = New DirectX8
    'Create the Direct3D object
    Set d3d = dX.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        .windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = d3ddm.Format 'current display depth
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

    DeviceRenderStates
    
    '****************************************************
    'Inicializamos el manager de texturas
    '****************************************************
    Call DXPool.Texture_Initialize(500)
    
    '****************************************************
    'Clears the buffer to start rendering
    '****************************************************
    Device_Clear
    '****************************************************
    'Load Misc
    '****************************************************
    LoadGraphicFonts
    LoadFonts
    'CargarParticulas
    Exit Function
ErrHandler:
    DXEngine_Initialize = False
End Function

Public Function DXEngine_BeginRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_BeginRender = True
    
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then
        Do
            DoEvents
        Loop While ddevice.TestCooperativeLevel = D3DERR_DEVICELOST
        
        DXPool.Texture_Remove_All
        Fonts_Destroy
        Device_Reset
        
        DeviceRenderStates
        LoadFonts
        LoadGraphicFonts
    End If
    
    '****************************************************
    'Render
    '****************************************************
    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    DXEngine_BeginRender = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function DXEngine_EndRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_EndRender = True

    If engine_render_started = False Then
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        FPS = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    

    
    
    engine_render_started = False
Exit Function
ErrorHandler:
    DXEngine_EndRender = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Private Function Device_Reset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler:
'On Error Resume Next

    'Be sure the scene is finished
    ddevice.EndScene
    'Reset device
    ddevice.Reset d3dpp
    
    DeviceRenderStates
       
Exit Function
ErrHandler:
    Device_Reset = Err.Number
End Function
Public Sub DXEngine_TextureRenderAdvance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal Src_X As Long, ByVal Src_Y As Long, _
                                             ByVal dest_width As Long, ByVal dest_height As Long, ByVal src_width As Long, ByVal src_height As Long, ByRef rgb_list() As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'This sub allow texture resizing
'
'**************************************************************

    
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim Texture As Direct3DTexture8
    Dim texture_width As Long
    Dim texture_height As Long

    'rgb_list(0) = RGB(255, 255, 255)
    'rgb_list(1) = RGB(255, 255, 255)
    'rgb_list(2) = RGB(255, 255, 255)
    'rgb_list(3) = RGB(255, 255, 255)
    
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    With src_rect
        .Bottom = Src_Y + src_height
        .Right = Src_X + src_width
        .Top = Src_Y
        .Left = Src_X
    End With
    
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    'Geometry_Create_Box temp_verts(), dest_rect, src_rect, texture_width, texture_height, angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Public Sub DXEngine_TextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Last Modify Date: 25/08/2012 - ^[GS]^
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Long
    Dim texture_width As Long
    Dim Texture As Direct3DTexture8
    
    'Set up the source rectangle
    With src_rect
        .Bottom = Src_Y + src_height - 1
        .Left = Src_X
        .Right = Src_X + src_width - 1
        .Top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Geometry_Create_Box temp_verts(), dest_rect, src_rect, texture_width, texture_height, angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function


Public Sub DXEngine_GraphicTextRender(Font_Index As Integer, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                  ByVal Color As Long)

    If Len(Text) > 255 Then Exit Sub
    
    Dim i As Byte
    Dim X As Integer
    Dim rgb_list(3) As Long
    
    For i = 0 To 3
        rgb_list(i) = Color
    Next i
    
    X = -1
    Dim Char As Integer
    For i = 1 To Len(Text)
        Char = AscB(mid$(Text, i, 1)) - 32
        
        If Char = 0 Then
            X = X + 1
        Else
            X = X + 1
            Call DXEngine_TextureRenderAdvance(gfont_list(Font_Index).texture_index, Left + X * gfont_list(Font_Index).Char_Size, _
                                                        Top, gfont_list(Font_Index).Caracteres(Char).Src_X, gfont_list(Font_Index).Caracteres(Char).Src_Y, _
                                                            gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, _
                                                                rgb_list(), False)
        End If
    Next i
    
    
    
End Sub

Public Sub DXEngine_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    'El manager de texturas es ahora independiente del engine.
    Call DXPool.Texture_Remove_All
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dX = Nothing
    Set DXPool = Nothing
End Sub

Private Sub LoadChars(ByVal Font_Index As Integer)
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For i = 0 To 255
        With gfont_list(Font_Index).Caracteres(i)
            X = (i Mod 16) * gfont_list(Font_Index).Char_Size
            If X = 0 Then '16 chars per line
                Y = Y + 1
            End If
            .Src_X = X
            .Src_Y = (Y * gfont_list(Font_Index).Char_Size) - gfont_list(Font_Index).Char_Size
        End With
    Next i
End Sub
Public Sub LoadGraphicFonts()
    Dim i As Byte
    Dim file_path As String

    file_path = DirIndex & "GUIFonts.ini"

    If General_File_Exist(file_path, vbArchive) Then
        gfont_count = general_var_get(file_path, "INIT", "FontCount")
        If gfont_count > 0 Then
            ReDim gfont_list(1 To gfont_count) As tGraphicFont
            For i = 1 To gfont_count
                With gfont_list(i)
                    .Char_Size = general_var_get(file_path, "FONT" & i, "Size")
                    .texture_index = general_var_get(file_path, "FONT" & i, "Graphic")
                    If .texture_index > 0 Then Call DXPool.Texture_Load(.texture_index, 0)
                    LoadChars (i)
                End With
            Next i
        End If
    End If
End Sub

Public Sub DXEngine_StatsRender()
    'fps
    Call DXEngine_TextRender(1, FPS & " FPS", 0, 0, D3DColorXRGB(255, 255, 255))
End Sub

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, screen_hwnd, ByVal 0&
End Sub

Private Sub DeviceRenderStates()
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True


    End With
End Sub

Private Sub Font_Make(ByVal Style As String, ByVal Size As Long, ByVal italic As Boolean, ByVal bold As Boolean)
    font_count = font_count + 1
    ReDim Preserve font_list(1 To font_count)
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.name = Style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    font_list(font_count).Size = Size
    Set font_list(font_count).dFont = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Sub LoadFonts()
    Dim num_fonts As Integer
    Dim i As Integer
    Dim file_path As String
    
    file_path = DirIndex & "fonts.ini"
    
    If Not General_File_Exist(file_path, vbArchive) Then Exit Sub
    
    num_fonts = general_var_get(file_path, "INIT", "FontCount")
    
    For i = 1 To num_fonts
        Call Font_Make(general_var_get(file_path, "FONT" & i, "Name"), general_var_get(file_path, "FONT" & i, "Size"), general_var_get(file_path, "FONT" & i, "Cursiva"), general_var_get(file_path, "FONT" & i, "Negrita"))
    Next i
End Sub
Public Sub DXEngine_TextRender(ByVal Font_Index As Integer, ByVal Text As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Color As Long, Optional ByVal Alingment As Byte = DT_LEFT, Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0)
    If Not Font_Check(Font_Index) Then Exit Sub
    
    Dim TextRect As RECT 'This defines where it will be
    'Dim BorderColor As Long
    
    'Set width and height if no specified
    If Width = 0 Then Width = Len(Text) * (font_list(Font_Index).Size + 1)
    If Height = 0 Then Height = font_list(Font_Index).Size * 2
    
    'DrawBorder
    
    'BorderColor = D3DColorXRGB(0, 0, 0)
    
    'TextRect.top = top - 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left - 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top + 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left + 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.Bottom = Top + Height
    TextRect.Right = Left + Width
    d3dx.DrawText font_list(Font_Index).dFont, Color, Text, TextRect, Alingment

End Sub
Private Function Font_Check(ByVal Font_Index As Long) As Boolean
    If Font_Index > 0 And Font_Index <= font_count Then
        Font_Check = True
    End If
End Function

Private Sub Fonts_Destroy()
    Dim i As Integer
    
    For i = 1 To font_count
        Set font_list(i).dFont = Nothing
        font_list(i).Size = 0
    Next i
    font_count = 0
End Sub

Public Function D3DColorValueGet(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As D3DCOLORVALUE
    D3DColorValueGet.A = A
    D3DColorValueGet.R = R
    D3DColorValueGet.G = G
    D3DColorValueGet.B = B
End Function
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Public Sub DXEngine_TextureToHdcRender(ByVal texture_index As Long, desthdc As Long, ByVal screen_x As Long, ByVal screen_Y As Long, ByVal SX As Integer, ByVal SY As Integer, ByVal sW As Integer, ByVal sH As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************

    Dim file_path As String
    Dim Src_X As Long
    Dim Src_Y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long

    file_path = DirGraficos & texture_index & ".bmp"
    If FileExist(file_path, vbArchive) Then
        If frmMain.picTemp.Tag <> file_path Then
            frmMain.picTemp.Picture = LoadPicture(file_path)
            frmMain.picTemp.Tag = file_path
        End If
    Else
        file_path = DirGraficos & texture_index & ".png"
        If FileExist(file_path, vbArchive) Then
            If frmMain.picTemp.Tag <> file_path Then
                Call modPngGDI.PngPictureLoad(file_path, frmMain.picTemp, False)
                frmMain.picTemp.Tag = file_path
            End If
        End If
    End If
    
    Src_X = SX
    Src_Y = SY
    src_width = sW
    src_height = sH

    hdcsrc = CreateCompatibleDC(desthdc)
    
    SelectObject hdcsrc, frmMain.picTemp.Picture
    
    If transparent = False Then
        BitBlt desthdc, screen_x, screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, SRCCOPY
    Else
        TransparentBlt desthdc, screen_x, screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, COLOR_KEY
    End If
        
    DeleteDC hdcsrc
End Sub

Public Sub DXEngine_BeginSecondaryRender()
    Device_Clear
    ddevice.BeginScene
End Sub
Public Sub DXEngine_EndSecondaryRender(ByVal hWnd As Long, ByVal Width As Integer, ByVal Height As Integer)
    Dim DR As RECT
    DR.Left = 0
    DR.Top = 0
    DR.Bottom = Height
    DR.Right = Width
    
    ddevice.EndScene
    ddevice.Present DR, ByVal 0&, hWnd, ByVal 0&
End Sub

Public Sub DXEngine_DrawBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color As Long, Optional ByVal border_width = 1)
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
    ddevice.SetTexture 0, Nothing
    
    'Upper Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Top, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Top + border_width, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Top + border_width, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Left Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Left + border_width, box_rect.Top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Left + border_width, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Top, 0, 2, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Bottom, 0, 2, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Right Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.Top, 0, 3, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.Bottom, 0, 3, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Left, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.A, Color.R, Color.G, Color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

'**************************************
' Name: Take Screenshot (DirectX 8)
' Description:This small function call takes a screenshot of DirectX's front buffer and saves it in the given path. The given path is the path where you want all the screenshots kept. This is so that you can simply call the function and it will automatically number the file for you.
' By: Rob Loach
'
' Inputs:'Direct3DDevice - The Direct3DDevice8 that you are using.
' D3DX - The D3DX8 that you are using.
' FilePath - The path of which you want all screenshots saved.
' ScreenWidth - The X resolution.
' ScreenHeight - The Y resolution.
'
' Returns:True or False depending on if it worked.
'
' Assumes:Assumes basic surface knowledge of DirectX.
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/sc ... gWId=1'for details.'**************************************
 
Public Function TakeScreenShot(Direct3DDevice As Direct3DDevice8, d3dx As D3DX8, ByVal ScreenHeight As Long, ByVal ScreenWidth As Long, Optional ByVal FilePath As String) As Boolean
 ' This function takes a screenshot of DirectX's front buffer.
 ' Returns true or false depending on if it worked.
 ' By: Rob Loach
 
 ' Declare variables
 Dim ScreenShot As Direct3DSurface8 ' The pointer to our new screen buffer which will be saved to a file
 Dim FileName As String ' The file name of the screen shot
 Dim X As Long ' The loop holder
 Dim PalEntry As PALETTEENTRY ' The pallette entry used to save the BMP
 Dim SourceRect As RECT ' The source rectangle used when saving the BMP
 
 ' See if the path exists
 If Len(FilePath) = 0 Then FilePath = App.Path
 If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
 If Len(Dir(FilePath, vbDirectory)) = 0 Then Exit Function
 
 ' Find a file path that doesn't exist
 Do
 X = X + 1
 FileName = FilePath & "screen" & Format(X, "0000") & ".bmp"
 Loop Until Len(Dir(FileName)) = 0
 
 ' 1. Create the image surface
 Set ScreenShot = Direct3DDevice.CreateImageSurface(ScreenWidth, ScreenHeight, D3DFMT_A8R8G8B8)
 If ScreenShot Is Nothing Then Exit Function ' Check if there was an error
 
 ' 2. Put the front buffer into the screen shot surface
 'Direct3DDevice.GetFrontBuffer ScreenShot
 Direct3DDevice.GetFrontBuffer ScreenShot
 
 ' 3. Save the whole screen as a BMP
 SourceRect.Right = ScreenWidth
 SourceRect.Bottom = ScreenHeight
 d3dx.SaveSurfaceToFile FileName, D3DXIFF_BMP, ScreenShot, PalEntry, SourceRect
 
 ' Delete allocated memory and return the result
 Set ScreenShot = Nothing
 TakeScreenShot = True
 
End Function
