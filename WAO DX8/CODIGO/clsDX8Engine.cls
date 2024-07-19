VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim font_count As Long
Dim font_last As Long

Private font_list() As D3DXFont

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

Dim texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private particletimer As Single
Public engineBaseSpeed As Single
Private lFrameTimer As Long
Private lFrameLimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte
Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
Private Const GrhFogata As Integer = 1521

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

Private Enum PARTICLE_STATUS
    Alive = 0
    Dead = 1
End Enum

Private Type PARTICLE
    X As Single     'World Space Coordinates
    Y As Single
    Z As Single
    vX As Single    'Speed and Direction
    vY As Single
    vZ As Single
    StartColor As D3DCOLORVALUE
    EndColor As D3DCOLORVALUE
    CurrentColor As D3DCOLORVALUE
    lifeTime As Long    'How long Mr. Particle Exists
    created As Long 'When this particle was created...
    status As PARTICLE_STATUS 'Does he even exist?
End Type

Private Type pa_gro
    PrtData() As PARTICLE
    PrtVertList() As TLVERTEX
    Position As D3DVECTOR
    Light As D3DLIGHT8
    type As Integer
    nParticles As Long
    ParticleSize As Single
    gravity As Single
    XWind As Single
    ZWind As Single
    YWind As Single
    XVariation As Single
    YVariation As Single
    ZVariation As Single
    X As Single
    Y As Single
    Z As Single
    activated As Boolean
    texture As Integer
    size As Single
    life As Integer
End Type

Dim particle_group_list() As pa_gro
Dim particle_group_count As Integer
Dim particle_group_last As Integer

'Dim fxs_list() As New clsFXS
'Dim fxs_count As Integer
'Dim fxs_last As Integer


Dim valoreBlur As Long
Dim dimeTex As Long
Dim tex As Direct3DTexture8
Dim D3DbackBuffer As Direct3DSurface8
Dim zTarget As Direct3DSurface8
Dim stencil As Direct3DSurface8
Dim superTex As Direct3DSurface8
Dim blur As Boolean
Public blur_factor As Byte

Dim bump_map_texture As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported As Boolean
Dim bump_map_powa As Boolean
'Sets a Grh animation to loop indefinitely.







#Const HARDCODED = False 'True ' == MÁS FPS ^^






Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal ob As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
Dim r As Long
Dim g As Long
Dim b As Long
r = o_r
g = og
b = ob
    
    If o_r > r_ Then
        r = r - 1
    ElseIf o_r < r_ Then
        r = r + 1
    End If

    If og > g_ Then
        g = g - 1
    ElseIf og < r Then
        g = g + 1
    End If

    If ob > b_ Then
        b = b - 1
    ElseIf og < r Then
        b = b + 1
    End If
If r > 255 Then r = 255
If g > 255 Then g = 255
If b > 255 Then b = 255
day_r_old = r
day_g_old = g
day_b_old = b

If r < 0 Or g < 0 Or b < 0 Then r = 1: g = 1: b = 1
change_day_effect = ARGB(r, g, b, 255)

End Function

Public Sub setup_ambient()

'Noche 87, 61, 43
luz_dia(0).r = 87
luz_dia(0).g = 61
luz_dia(0).b = 43
luz_dia(1).r = 87
luz_dia(1).g = 61
luz_dia(1).b = 43
luz_dia(2).r = 87
luz_dia(2).g = 61
luz_dia(2).b = 43
luz_dia(3).r = 87
luz_dia(3).g = 61
luz_dia(3).b = 43
'4 am 124,117,91
luz_dia(4).r = 124
luz_dia(4).g = 127
luz_dia(4).b = 91
'5,6 am 143,137,135
luz_dia(5).r = 143
luz_dia(5).g = 137
luz_dia(5).b = 135
luz_dia(6).r = 143
luz_dia(6).g = 137
luz_dia(6).b = 135
'7 am 212,205,207
luz_dia(7).r = 212
luz_dia(7).g = 205
luz_dia(7).b = 207
luz_dia(8).r = 212
luz_dia(8).g = 205
luz_dia(8).b = 207
luz_dia(9).r = 212
luz_dia(9).g = 205
luz_dia(9).b = 207
luz_dia(10).r = 212
luz_dia(10).g = 205
luz_dia(10).b = 207
luz_dia(11).r = 212
luz_dia(11).g = 205
luz_dia(11).b = 207
luz_dia(12).r = 212
luz_dia(12).g = 205
luz_dia(12).b = 207
'Dia 255, 255, 255
luz_dia(12).r = 255
luz_dia(12).g = 255
luz_dia(12).b = 255
luz_dia(13).r = 255
luz_dia(13).g = 255
luz_dia(13).b = 255
'Medio Dia 255, 200, 255
luz_dia(14).r = 255
luz_dia(14).g = 250
luz_dia(14).b = 255
luz_dia(15).r = 255
luz_dia(15).g = 240
luz_dia(15).b = 255
luz_dia(16).r = 255
luz_dia(16).g = 230
luz_dia(16).b = 255
'17/18 0, 100, 255
luz_dia(17).r = 230
luz_dia(17).g = 230
luz_dia(17).b = 255
'18/19 0, 100, 255
luz_dia(18).r = 230
luz_dia(18).g = 230
luz_dia(18).b = 255
'19/20 156, 142, 83
luz_dia(19).r = 156
luz_dia(19).g = 142
luz_dia(19).b = 83
luz_dia(20).r = 87
luz_dia(20).g = 61
luz_dia(20).b = 43
luz_dia(21).r = 87
luz_dia(21).g = 61
luz_dia(21).b = 43
luz_dia(22).r = 87
luz_dia(22).g = 61
luz_dia(22).b = 43
luz_dia(23).r = 87
luz_dia(23).g = 61
luz_dia(23).b = 43
luz_dia(24).r = 87
luz_dia(24).g = 61
luz_dia(24).b = 43
            
End Sub


Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
   
    Dim TextRect As RECT
    Dim ShadowRect As RECT
   
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.bottom = Top + height
    TextRect.Right = Left + width
   
    If shadow Then
        ShadowRect.Top = Top
        ShadowRect.Left = Left - 1
        ShadowRect.bottom = (Top + height)
        ShadowRect.Right = (Left + width) - 1
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left
        ShadowRect.bottom = (Top + height) - 1
        ShadowRect.Right = (Left + width)
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
        ShadowRect.Top = Top
        ShadowRect.Left = Left + 1
        ShadowRect.bottom = (Top + height)
        ShadowRect.Right = (Left + width) + 1
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
        ShadowRect.Top = Top + 1
        ShadowRect.Left = Left
        ShadowRect.bottom = (Top + height) + 1
        ShadowRect.Right = (Left + width)
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
    End If
   
    D3DX.DrawText font, color, Text, TextRect, format
End Sub
 
Public Sub Text_Render_ext(ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal width As Long, ByVal height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)
    If center = True Then
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top, Left, width, height, color, DT_TOP Or DT_LEFT, shadow)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.Z = Z
End Function

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With
    DispMode.format = D3DFMT_X8R8G8B8
    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then
        Dim Caps8 As D3DCAPS8
        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8
        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format
        End If
    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format
    End If
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.02
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set FPS value to 60 for startup
    FPS = 100
    FramesPerSecCounter = 100
    
    ScrollPixelsPerFrameX = 9
    ScrollPixelsPerFrameY = 9
    
    UserPos.X = 50
    UserPos.Y = 50
    
    MinXBorder = XMinMapSize
    MaxXBorder = XMaxMapSize
    MinYBorder = YMinMapSize
    MaxYBorder = YMaxMapSize

    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    'motion blur
    Set D3DbackBuffer = D3DDevice.GetRenderTarget
    Set zTarget = D3DDevice.GetDepthStencilSurface
    Set stencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set tex = D3DX.CreateTexture(D3DDevice, dimeTex, dimeTex, 1, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
    Set superTex = tex.GetSurfaceLevel(0)
    blur_factor = 10
    'bump mapping
    
    Font_Create "Tahoma", 8, True, 0
    'Font_Create "Verdana", 8, False, 0
    
bRunning = True
Exit Sub
ErrHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub Engine_Deinit()
Erase MapData
Erase charlist
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing
End
End Sub


Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub Engine_ActFPS()
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
End Sub

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub
Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, Optional ByVal alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)
    Dim CurrentGrhIndex As Integer
    If Grh.grhindex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If
    


    Device_Box_Textured_Render CurrentGrhIndex, _
        X, Y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value, _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        alpha _
        , angle
'exits:

End Sub

Public Sub Engine_Blur_Set(ByVal bool As Boolean)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = bool
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Engine_Blur_Toggle()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
blur = Not blur
If blur = True Then
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
Else
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
End If
End Sub

Public Sub Render()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************

D3DDevice.BeginScene

If blur Then
    D3DDevice.SetRenderTarget superTex, stencil, 0
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    D3DDevice.SetRenderTarget D3DbackBuffer, zTarget, 0
Else
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
End If
    ShowNextFrame
    
    If blur Then
        D3DDevice.SetTexture 0, tex
        D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
        D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(blur_factor, 255, 255, 255)
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertList(0), Len(vertList(0))
    End If

    Text_Render font_list(1), "MZEngine", 0, 470, 70, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    Text_Render font_list(1), "FPS: " & FPS, 15, 470, 60, 20, &HFFFFFFFF, DT_TOP Or DT_LEFT, True
    Text_Render font_list(1), "Coord: " & UserMap & "," & UserPos.X & "," & UserPos.Y, 398, 0, 100, 100, &HFFFFFFFF, DT_TOP Or DT_LEFT, False
    Text_Render font_list(1), "Usuarios: " & NumUsers, 384, 0, 100, 100, &HFFFFFFFF, DT_TOP Or DT_LEFT, False
    Call Dialogos.Render
    Call DialogosClanes.Draw
D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

lFrameLimiter = GetTickCount
FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
particletimer = timerElapsedTime * 0.05
End Sub



Sub ShowNextFrame()
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    

    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
            
#If HARDCODED = True Then
                    If MapData(X, Y).Graphic(1).Started = 1 Then
                        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames / MapData(X, Y).Graphic(1).Speed)
                        If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames Then
                            MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).grhindex).NumFrames) + 1
                            
                            If MapData(X, Y).Graphic(1).Loops <> -1 Then
                                If MapData(X, Y).Graphic(1).Loops > 0 Then
                                    MapData(X, Y).Graphic(1).Loops = MapData(X, Y).Graphic(1).Loops - 1
                                Else
                                    MapData(X, Y).Graphic(1).Started = 0
                                End If
                            End If
                        End If
                    End If

                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(1).grhindex).Frames(MapData(X, Y).Graphic(1).FrameCounter)

                Device_Box_Textured_Render CurrentGrhIndex, _
                    (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, _
                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    MapData(X, Y).light_value, _
                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    False _
                    , 0
#Else
                Call Draw_Grh(MapData(X, Y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, , X, Y)
                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, , X, Y)
                End If
#End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw floor layer 2
'    ScreenY = minYOffset
'    For Y = screenminY To screenmaxY
'        ScreenX = minXOffset
'        For X = screenminX To screenmaxX
                'Layer 2 **********************************
'                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
'#If HARDCODED = True Then
'                    If MapData(X, Y).Graphic(2).Started = 1 Then
'                        MapData(X, Y).Graphic(2).FrameCounter = MapData(X, Y).Graphic(2).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames / MapData(X, Y).Graphic(2).Speed)
'                        If MapData(X, Y).Graphic(2).FrameCounter > GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames Then
'                            MapData(X, Y).Graphic(2).FrameCounter = (MapData(X, Y).Graphic(2).FrameCounter Mod GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames) + 1
'
'                            If MapData(X, Y).Graphic(2).Loops <> -1 Then
'                                If MapData(X, Y).Graphic(2).Loops > 0 Then
'                                    MapData(X, Y).Graphic(2).Loops = MapData(X, Y).Graphic(2).Loops - 1
'                                Else
'                                    MapData(X, Y).Graphic(2).Started = 0
'                                End If
'                            End If
'                        End If
'                    End If
'
'                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(MapData(X, Y).Graphic(2).FrameCounter)
'
'                offx = 0
'                offy = 0
'                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                If GrhData(MapData(X, Y).Graphic(2).grhindex).TileHeight <> 1 Then _
'                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                Device_Box_Textured_Render CurrentGrhIndex, _
'                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
'                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                    MapData(X, Y).light_value, _
'                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                    False _
'                    , 0
'#Else
''                    Call Draw_Grh(MapData(X, Y).Graphic(2), _
'                            (ScreenX - 1) * 32 + PixelOffsetX, _
'                            (ScreenY - 1) * 32 + PixelOffsetY, _
'                            1, 1, , X, Y)
'#End If
''                End If
'
''            ScreenX = ScreenX + 1
''        Next X'
'
'        'Reset ScreenX to original value and increment ScreenY
'        'ScreenX = ScreenX - X + screenminX
'        'ScreenY = ScreenY + 1
'    'Next Y
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                'Object Layer **********************************
                If .ObjGrh.grhindex <> 0 Then
'#If HARDCODED = True Then
'                    CurrentGrhIndex = GrhData(.ObjGrh.grhindex).Frames(.ObjGrh.FrameCounter)
'                    offx = 0
'                    offy = 0
'                    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                        offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                    If GrhData(.ObjGrh.grhindex).TileHeight <> 1 Then _
'                        offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                    Device_Box_Textured_Render CurrentGrhIndex, _
'                        PixelOffsetXTemp + offx, PixelOffsetYTemp + offx, _
'                        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                        MapData(X, Y).light_value, _
'                        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                        False _
'                        , 0
'#Else
                    Call Draw_Grh(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
'#End If
                End If
                '***********************************************
                'If .particle_group > 0 Then
                '    Particles_Update .particle_group, PixelOffsetYTemp, PixelOffsetXTemp
       '
       '             D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
       '             D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
       '
       '             Dim tmplng As Long
       '             D3DDevice.SetTexture 0, SurfaceDB.GetTexture(GrhData(particle_group_list(.particle_group).texture).FileNum, tmplng) 'Textures(particle_group_list(.particle_group).Texture)
       '             D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(particle_group_list(.particle_group).size)
       '             D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, particle_group_list(.particle_group).nParticles, particle_group_list(.particle_group).PrtVertList(0), Len(particle_group_list(.particle_group).PrtVertList(0))
       '             D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
       '             D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
       '
       '         End If
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If .Graphic(3).grhindex <> 0 Then
'#If HARDCODED = True Then
'                    If .Graphic(3).Started = 1 Then
'                        .Graphic(3).FrameCounter = .Graphic(3).FrameCounter + (timerElapsedTime * GrhData(.Graphic(3).grhindex).NumFrames / .Graphic(3).Speed)
'                        If .Graphic(3).FrameCounter > GrhData(.Graphic(3).grhindex).NumFrames Then
'                            .Graphic(3).FrameCounter = (.Graphic(3).FrameCounter Mod GrhData(.Graphic(3).grhindex).NumFrames) + 1
'
'                            If .Graphic(3).Loops <> -1 Then
'                                If .Graphic(3).Loops > 0 Then
'                                    .Graphic(3).Loops = .Graphic(3).Loops - 1
'                                Else
'                                    .Graphic(3).Started = 0
'                                End If
'                            End If
'                        End If
'                    End If
'
'                CurrentGrhIndex = GrhData(.Graphic(3).grhindex).Frames(.Graphic(3).FrameCounter)
'
'                offx = 0
'                offy = 0
'                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                If GrhData(.Graphic(3).grhindex).TileHeight <> 1 Then _
'                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                Device_Box_Textured_Render CurrentGrhIndex, _
'                    PixelOffsetXTemp + offx, PixelOffsetYTemp + offy, _
'                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                    MapData(X, Y).light_value, _
'                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                    False _
'                    , 0
'#Else
                    Call Draw_Grh(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
'#End If
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
                    
                    
                    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    Dim tmplng As Long
    Dim tmblng2 As Long
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            With MapData(X, Y)
                '***********************************************
                If .particle_group > 0 Then
                    Particles_Update .particle_group, ScreenY * 32 + PixelOffsetY, ScreenX * 32 + PixelOffsetX
        
                    
                    D3DDevice.SetTexture 0, SurfaceDB.GetTexture(GrhData(particle_group_list(.particle_group).texture).FileNum, tmplng, tmblng2) 'Textures(particle_group_list(.particle_group).Texture)
                    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(particle_group_list(.particle_group).size)
                    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, particle_group_list(.particle_group).nParticles, particle_group_list(.particle_group).PrtVertList(0), Len(particle_group_list(.particle_group).PrtVertList(0))

                End If
                'If .FXS_group <> 0 Then Call fxs_list(.FXS_group).render(PixelOffsetXTemp, PixelOffsetYTemp)
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
                    
                    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
If Not bTecho Then
        'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).grhindex Then
'#If HARDCODED = True Then
'                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(4).grhindex).Frames(MapData(X, Y).Graphic(4).FrameCounter)
'                offx = 0
'                offy = 0
'                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
'                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
'                If GrhData(MapData(X, Y).Graphic(4).grhindex).TileHeight <> 1 Then _
'                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
'
'                Device_Box_Textured_Render CurrentGrhIndex, _
'                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
'                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
'                    MapData(X, Y).light_value, _
'                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
'                    False _
'                    , 0
'#Else
                    Call Draw_Grh(MapData(X, Y).Graphic(4), _
                        ScreenX * 32 + PixelOffsetX, _
                        ScreenY * 32 + PixelOffsetY, _
                        1, 1, , X, Y)
'#End If
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    Else
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).grhindex Then
                    Call Draw_Grh(MapData(X, Y).Graphic(4), _
                        ScreenX * 32 + PixelOffsetX, _
                        ScreenY * 32 + PixelOffsetY, _
                        1, 1, True, X, Y)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Function Geometry_Create_TLVertex2(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
'mz
Geometry_Create_TLVertex2.X = X
Geometry_Create_TLVertex2.Y = Y
Geometry_Create_TLVertex2.Z = Z
Geometry_Create_TLVertex2.rhw = rhw
Geometry_Create_TLVertex2.color = color
Geometry_Create_TLVertex2.Specular = Specular
Geometry_Create_TLVertex2.tu1 = tu1
Geometry_Create_TLVertex2.tv1 = tv1
Geometry_Create_TLVertex2.tu2 = tu2
Geometry_Create_TLVertex2.tv2 = tv2
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
    End If

End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    'If Not char_current_blind Then
    '    If (light_value(0) = 0) Then light_value(0) = 0
    '    If (light_value(1) = 0) Then light_value(1) = 0
    '    If (light_value(2) = 0) Then light_value(2) = 0
    '    If (light_value(3) = 0) Then light_value(3) = 0
    'Else
        light_value(0) = &HFFFFFFFF 'blind_color
        light_value(1) = &HFFFFFFFF 'blind_color
        light_value(2) = &HFFFFFFFF 'blind_color
        light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'If Not char_current_blind Then
        If (light_value(0) = 0) Then light_value(0) = base_light
        If (light_value(1) = 0) Then light_value(1) = base_light
        If (light_value(2) = 0) Then light_value(2) = base_light
        If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
   
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
               
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .body.Walk(.Heading).Speed > 0 Then _
                    .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
               
                'Char moved
                moved = True
               
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
           
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
               
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .body.Walk(.Heading).Speed > 0 Then _
                    .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
               
                'Char moved
                moved = True
               
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
       
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .body.Walk(.Heading).Started = 0
            .body.Walk(.Heading).FrameCounter = 1
           
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
           
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
           
            .Moving = False
        End If
       
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
       
        If .head.head(.Heading).grhindex Then
            If Not .invisible Then
                'Draw Body
 
                            If .body.Walk(.Heading).grhindex Then _
                                Call Draw_Grh(.body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                               
                            If .head.head(.Heading).grhindex Then
                                Call Draw_Grh(.head.head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, , X, Y)
                                If .Casco.head(.Heading).grhindex Then _
                                    Call Draw_Grh(.Casco.head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, , X, Y)
                                If .Arma.WeaponWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                                If .Escudo.ShieldWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                            End If
                           
                    'Draw name over head
                    If Nombres Then
                        If Len(.Nombre) > 0 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2
                           
                            If .priv = 0 Then
                                If .Criminal Then
                                    color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                Else
                                    color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                End If
                            Else
                                color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                           
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call Text_Render(font_list(1), line, PixelOffsetY + 30, PixelOffsetX - 20, 72, 20, color, fa_center, True)
                           
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call Text_Render(font_list(1), line, PixelOffsetY + 45, PixelOffsetX - 35, 100, 20, color, fa_center, True)
                        End If
                    End If
             Else
            Dim sclan As String
                Dim UserClan As String
                Pos = InStr(.Nombre, "<")
                If Pos <> 0 Then
                sclan = Trim(mid$(charlist(UserCharIndex).Nombre, InStr(charlist(UserCharIndex).Nombre, "<")))
                UserClan = mid$(.Nombre, InStr(.Nombre, "<"))
                Else
                sclan = "a"
                UserClan = "b"
                End If
                If (UCase$(.Nombre) = UCase$(frmMain.Label8.Caption)) Or sclan = UserClan Then
                   
                If .body.Walk(.Heading).grhindex Then _
                                Call Draw_Grh(.body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, X, Y)
                               
                            If .head.head(.Heading).grhindex Then
                                Call Draw_Grh(.head.head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, True, X, Y)
                                If .Casco.head(.Heading).grhindex Then _
                                    Call Draw_Grh(.Casco.head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, True, X, Y)
                                If .Arma.WeaponWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, X, Y)
                                If .Escudo.ShieldWalk(.Heading).grhindex Then _
                                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, True, X, Y)
                            End If
                           
                    'Draw name over head
                    If Nombres Then
                        If Len(.Nombre) > 0 Then
                            Pos = InStr(.Nombre, "<")
                            If Pos = 0 Then Pos = Len(.Nombre) + 2
                           
                            If .priv = 0 Then
                                If .Criminal Then
                                    color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                Else
                                    color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                End If
                            Else
                                color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                           
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call Text_Render(font_list(1), line, PixelOffsetY + 30, PixelOffsetX - 20, 72, 20, color, fa_center, True)
                           
                            'Clan
                            line = mid$(.Nombre, Pos)
                            Call Text_Render(font_list(1), line, PixelOffsetY + 45, PixelOffsetX + 21, 100, 20, color, fa_center, True)
                        End If
                    End If
               
                End If
            End If
        Else
            'Draw Body
            If .body.Walk(.Heading).grhindex Then _
                Call Draw_Grh(.body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
        End If
 
        ''Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .body.HeadOffset.X - 155, PixelOffsetY + .body.HeadOffset.Y, CharIndex)
       
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.Fx, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, True)
 
            If .Fx.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub
 
Public Sub Start()

Dim f As Boolean
Dim ulttick As Long, esttick As Long
Dim timers(1 To 2) As Integer
Dim loopc As Integer


    day_r_old = luz_dia(Hour(Time)).r
    day_g_old = luz_dia(Hour(Time)).g
    day_b_old = luz_dia(Hour(Time)).b
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
    
DoEvents
vertList(0).X = 0: vertList(0).Y = 0
vertList(1).X = 800: vertList(1).Y = 0
vertList(2).X = 0: vertList(2).Y = 600
vertList(3).X = 800: vertList(3).Y = 600
vertList(0).rhw = 1
vertList(1).rhw = 1
vertList(2).rhw = 1
vertList(3).rhw = 1
vertList(0).color = D3DColorXRGB(255, 255, 255)
vertList(1).color = D3DColorXRGB(255, 255, 255)
vertList(2).color = D3DColorXRGB(255, 255, 255)
vertList(3).color = D3DColorXRGB(255, 255, 255)
vertList(0).tu = 0
vertList(0).tv = 0
vertList(1).tu = 1
vertList(1).tv = 0
vertList(2).tu = 0
vertList(2).tv = 1
vertList(3).tu = 1
vertList(3).tv = 1
On Error Resume Next ' putos timers!!!!!
Do While prgRun

If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
    CheckKeys
    engine.Render
    engine.DrawInv
Else
    Sleep 10&
End If
DoEvents
    Rem Limitar FPS
'    While (GetTickCount - lFrameLimiter) < FramesPerSecCounter
'        Sleep 1
'    Wend
'    While GetTickCount - lFrameLimiter < 55
'       Sleep 5
'    Wend

        esttick = GetTickCount
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)
            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        Next loopc
        ulttick = GetTickCount

        DoEvents
Loop
engine.Engine_Deinit

    EngineRun = False
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    
End
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal Fx As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = Fx
        
        If .FxIndex > 0 Then
            Call InitGrh(.Fx, FxData(Fx).Animacion)
        
            .Fx.Loops = Loops
        End If
    End With
End Sub

Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then
    Call Dialogos.RemoveDialog(CharIndex)
    'Call Hits.RemoveHit(CharIndex)
    End If
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Public Sub Particle_Group_Make( _
ByVal particle_group_index As Integer, _
ByVal map_z As Integer, _
ByVal map_x As Integer, _
ByVal map_y As Integer, _
ByVal var_z As Single, _
ByVal var_x As Single, _
ByVal var_y As Single, _
ByVal start_a As Integer, _
ByVal start_r As Integer, _
ByVal start_g As Integer, _
ByVal start_b As Integer, _
ByVal end_a As Integer, _
ByVal end_r As Integer, _
ByVal end_g As Integer, _
ByVal end_b As Integer, _
ByVal n_particles As Integer, _
ByVal gravity As Single, _
ByVal texture As Integer, _
ByVal size As Single, _
ByVal life As Integer)
    'If particle_group_index > particle_group_last Then
    '    particle_group_last = particle_group_index
    
    ReDim Preserve particle_group_list(0 To particle_group_index)
    With particle_group_list(particle_group_index)
        ReDim Preserve .PrtData(0 To n_particles)
        ReDim Preserve .PrtVertList(0 To n_particles)
        'Map pos
        .X = map_x
        .Y = map_y
        .Z = map_z
        .life = life
        .size = size
        .gravity = gravity '-0.5
        'Sets alive counter
        .nParticles = n_particles
        .texture = texture
        Dim i As Integer
        .XVariation = var_x '0.5
        .YVariation = var_y '0.85
        .ZVariation = var_z '0.5
        Dim nowa As Long
        nowa = GetTickCount
        For i = 0 To .nParticles
            .PrtData(i).status = Alive
            .PrtData(i).lifeTime = 1000 + (Rnd * 500)
            .PrtData(i).created = nowa
            .PrtData(i).X = map_x
            .PrtData(i).Y = map_y
            '.PrtData(I).Z = map_z
            .PrtData(i).vX = (Rnd * var_x) - (var_x / 2)
            .PrtData(i).vY = (Rnd * var_y) - (var_y / 3)
            '.PrtData(I).vZ = (Rnd * var_z) - (var_z / 2)
            'Randomize
            .PrtData(i).StartColor = CreateColorVal(start_a, start_r, start_g, start_g)
            .PrtData(i).EndColor = CreateColorVal(end_a, end_r, end_g, end_b)
            .PrtData(i).CurrentColor = .PrtData(i).StartColor
            .PrtVertList(i).X = .PrtData(i).X
            .PrtVertList(i).rhw = 1
            .PrtVertList(i).Y = .PrtData(i).Y ', 0) '.PrtData(I).Z)
            .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.A, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.g, .PrtData(i).CurrentColor.b)
        Next i
        MapData(map_x, map_y).particle_group = particle_group_index
    End With
    
    particle_group_count = particle_group_count + 1
End Sub

'Public Sub FXMake( _
ByVal Index As Integer, _
ByVal map_y As Integer, _
ByVal map_x As Integer, _
ByVal tipo As fxs_type)
'    Dim datos(2) As Byte
'    ReDim Preserve fxs_list(0 To Index)
    
'    MapData(map_x, map_y).FXS_group = Index
'    fxs_list(Index).Set_fx tipo, 0, datos()
'End Sub

Private Function CreateColorVal(A As Integer, r As Integer, g As Integer, b As Integer) As D3DCOLORVALUE
    CreateColorVal.A = A
    CreateColorVal.r = r
    CreateColorVal.g = g
    CreateColorVal.b = b
End Function

Private Sub Particles_Update(ii As Integer, offy As Integer, offx As Integer)

    Dim i As Long
    Dim nowa As Long
    nowa = GetTickCount
    With particle_group_list(ii)
        For i = 0 To .nParticles - 1
            If nowa - .PrtData(i).created >= .PrtData(i).lifeTime Then .PrtData(i).status = Dead
            If .PrtData(i).status = Alive Then
                .PrtData(i).X = .PrtData(i).X + ((.PrtData(i).vX / 5)) * particletimer
                .PrtData(i).Y = .PrtData(i).Y + ((.PrtData(i).vY / 5)) * particletimer
                .PrtData(i).vX = .PrtData(i).vX + (.XWind / 5) * particletimer
                .PrtData(i).vY = .PrtData(i).vY + (.gravity / 5) * particletimer
                D3DXColorLerp .PrtData(i).CurrentColor, .PrtData(i).StartColor, .PrtData(i).EndColor, _
                    (nowa - .PrtData(i).created) / .PrtData(i).lifeTime
                .PrtVertList(i).X = .PrtData(i).X + offx
                .PrtVertList(i).Y = .PrtData(i).Y + offy
                .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.A, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.g, .PrtData(i).CurrentColor.b)
                'If .PrtData(I).y < -1 Then .PrtData(I).Status = Dead
            Else
                .PrtData(i).status = Alive
                .PrtData(i).lifeTime = .life + (Rnd * .life / 2)
                .PrtData(i).created = nowa
                .PrtData(i).X = .X
                .PrtData(i).Y = .Y
                .PrtData(i).vX = (Rnd * .XVariation) - (.XVariation * Rnd)
                .PrtData(i).vY = (Rnd * .YVariation) - (.YVariation * Rnd)
                .PrtData(i).CurrentColor = .PrtData(i).StartColor
                .PrtVertList(i).X = .PrtData(i).X + offx
                .PrtVertList(i).Y = .PrtData(i).Y + offy
                .PrtVertList(i).color = D3DColorARGB(.PrtData(i).CurrentColor.A, .PrtData(i).CurrentColor.r, .PrtData(i).CurrentColor.g, .PrtData(i).CurrentColor.b)
            End If
        Next i
    End With
End Sub

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
Dim r As Integer, g As Integer, b As Integer, A As Integer
    r = 127 * Vec.X + 128
    g = 127 * Vec.Y + 128
    b = 127 * Vec.Z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, g, b)
End Function



Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addy As Byte
    Dim addx As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        


        Select Case Heading
            Case NORTH
                addy = -1
        
            Case EAST
                addx = 1
        
            Case SOUTH
                addy = 1
            
            Case WEST
                addx = -1
        End Select
        
        nX = map_x + addx
        nY = map_y + addy
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'menduz
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(0) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(1) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(2) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(3) = color
                
    'For ia = 0 To 6.5 Step 0.05
    '    For i = 1 To light_list(light_index).range / 2
    '        x = (Sin(ia) * i) + light_list(light_index).map_x
    '        y = (Cos(ia) * i) + light_list(light_index).map_y
    '        'Debug.Print x; y
    '        If InMapBounds(x, y) Then
    '            If i = (light_list(light_index).range / 2) Then
    '                If MapData(x - 1, y - 1).light_value(2) <> color Then MapData(x, y).light_value(0) = 0
    '                If MapData(x + 1, y + 1).light_value(0) <> color Then MapData(x, y).light_value(2) = 0
    '                If MapData(x + 1, y - 1).light_value(1) <> color Then MapData(x, y).light_value(1) = 0
    '                If MapData(x - 1, y + 1).light_value(3) <> color Then MapData(x, y).light_value(3) = 0
    '                If MapData(x - 1, y).light_value(0) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x + 1, y).light_value(1) = color And MapData(x + 1, y).light_value(2) = color Then
    '                    MapData(x, y).light_value(0) = &HFFFF0FFF
    '                    MapData(x, y).light_value(3) = &HFFFF0FFF
    '                End If
    '                If MapData(x, y - 1).light_value(2) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x, y + 1).light_value(3) = color Then MapData(x, y).light_value(1) = color''

    '            Else
    '                MapData(x, y).light_value(0) = &HFFFFFFFF
    '                MapData(x, y).light_value(1) = &HFFFFFFFF
    '                MapData(x, y).light_value(2) = &HFFFFFFFF
    '                MapData(x, y).light_value(3) = &HFFFFFFFF
    '            End If
    '        End If
    '    Next i
    'Next ia
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = color
            MapData(X, min_y).light_value(2) = color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = color
            MapData(X, max_y).light_value(3) = color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = color
            MapData(min_x, Y).light_value(3) = color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = color
            MapData(max_x, Y).light_value(1) = color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = color
                MapData(X, Y).light_value(1) = color
                MapData(X, Y).light_value(2) = color
                MapData(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub

Public Function RemoveCharAparence(ByVal CharIndex As Integer, ByVal muertO As Integer, ByVal body As Integer, ByVal head As Integer, ByVal Heading As E_Heading, _
                        ByVal tempint As Integer, ByVal tempweap As Integer, ByVal tempesc As Integer, _
                        ByVal Fx As Integer, ByVal Loops As Integer)
                        
            CharIndex = CharIndex
            charlist(CharIndex).muertO = muertO = 500
            charlist(CharIndex).body = BodyData(body)
            charlist(CharIndex).head = HeadData(head)
            charlist(CharIndex).Heading = Heading
            charlist(CharIndex).FxIndex = Fx
            charlist(CharIndex).Fx.Loops = Loops
            tempint = tempint
            If tempint <> 0 Then charlist(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = tempweap
            If tempint <> 0 Then charlist(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = tempesc
            If tempint <> 0 Then charlist(CharIndex).Casco = CascoAnimData(tempint)
        Call SetCharacterFx(CharIndex, charlist(CharIndex).FxIndex, charlist(CharIndex).Fx.Loops)
            
            Call RefreshAllChars
End Function


Public Sub DrawInv()
    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
        Inventario.DrawInventory
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.picInv.hWnd, ByVal 0
End Sub

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = Y + height
        .Left = X
        .Right = X + width
        .Top = Y
    End With
    
    With Outline
        .bottom = Y + height + 2
        .Left = X - 2
        .Right = X + width + 2
        .Top = Y - 2
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub

