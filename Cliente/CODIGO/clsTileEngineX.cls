VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsTileEngineX - ORE Direct3D 8 Tile Engine - v0.5.0
'
'Graphics Engine.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'18/1/2005 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'   - Add: Particle_Type_Get method
'   - Add: Chats over head
'   - Add: Engine_Resolution_Change method
'   - Add: Particle_Data_Load_All method and related type
'   - Add: Grh_Render_Advance which allows grh resizing
'   - Add: Light_Move_By_Head method
'   - Change: Map_Render nolonger attempts to render anything outside map bounds, so it's faster, several FPS gained
'   - Change: Removed the clipping around edges code at the end of Map_Render since it was useless. Related vars and methods were removed as well
'   - Change: Engine now renders VSYNC since it's FPS have been really improved
'   - Change: Removed Engine_Deinitialize method, now it's all done when setting the engine object to Nothing
'   - Change: Map format was redesigned
'   - Change: Map_Base_Light_Fill was edited to store the base light color to the map structure
'   - Change: Font_Make now saves the StdFont so it can be recreated after changing resolution
'   - Change: Geometry_Create_Box method was slightly modified to improve speed
'   - Change: Light_Render method was recoded so now it's MUCH faster
'   - Change: Alpha blending is enabled only once at start up (or on reset) instead of every time something is rendered
'   - Change: the alpha rendering mode is only changed when alpha blending is required, and then set back to normal (less code executed, faster code)
'   - Change: GUI_Grh_Render was enhaced to support further customization on the way the grh is displayed
'   - Change: Almost every variable type was changed to save memory. Some variables were directly removed
'   - Change: Removed layer 5 since it wasn't used at all
'   - Change: When placing a NPC (map editor) it creates a char to be rendered, no need to set show special tile  to true anymore
'   - Change: Items are rendered after layer 2, no need to set show special tile  to true anymore
'   - Change: show_special_tiles is now show_exit tiles, related methods were renamed appropiately
'   - Change: Particle_Render was completely recoded, now it's faster and simpler
'   - Change: FontAlignment was improved to support properly the 9 possible alignments
'   - Change: Got rid of the header in the grh.dat file
'   - Change: The array of events in Input_Mouse_Poll was set to 1000 instead of 5000 since that's the limit set when creating the device
'   - Change: Several more small changes all over
'   - Fix: The mouse map pos is now updated even if the mouse is not being moved (since we may be scrolling the map)
'   - Fix: Engine_Render_Start now is able to avoid being called without having ended the scene
'   - Fix: Fonts are now destroyed to avoid memory leak
'   - Fix: Input_Mouse_Poll can now tell if a mouse button is pressed or not properly
'   - Fix: Memory-leak issue when deiitializing
'   - Fix: Memory-leak issue when rendering to a hDC
'   - Fix: The device is now properly reset
'   - Fix: The engine no longer crushes every now and then when you hit Ctrl-Alt-Del
'
'05/15/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: GUI_Grh_Render
'   -Change: Text_Render is now called GUI_Text_Render
'       -Added word_wrap option
'   Sub Release Contributors:
'       05/15/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: GUI_Box_Filled_Render
'           - Add: GUI_Box_Outline_Render
'
'05/12/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: Char_Data_Grh_Index_Get
'
'05/08/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Moved all the General_X methods to a external modGeneral.bas
'
'05/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Fix: Map_Save and Map_Load format changed. Added angle and alpha values.
'   Sub Release Contributors:
'       05/07/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Fix: Geometry_Create_Sqaure - fixed rotation
'
'05/02/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Fix: bug with Light_Move and base light values
'   - Fix: Grh_Get_Info and Grh_Render_To_Hdc when using animated grhs
'   - Add: Base light values can now be specfied for each corner
'      - Changed .map format. Old maps will not work anymore
'      - Modified Map_Base_Light_Set
'   - Remove: Light_Remove_XY (Use Light_Remove Map_Light_Get(X,Y) instead)
'
'04/29/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Add: Grh_Render_To_Hdc
'   - Add: Map_Save_To_File and Map_Load_From_File
'   - Add: Text_Render method
'   - Change: file name to clsTileEngineX.cls
'   - Change: Split Engine_Render into Engine_Render_Start and Engine_Render_End
'   - Change: Method Names
'       -Engine_Stats_Render
'       -Engine_Blocked_Tiles_Show_Toggle
'   Sub Release Contributors:
'       04/28/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: Grh_Add_GrhList_To_ListBox method
'           - Change: Particle_Render (added 3 new stream types)
'
'04/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - First release
'   Sub Release Contributors:
'       03/31/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Light_Erase method
'*****************************************************************

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************
Option Explicit

'***************************
'Constants
'***************************
Public Enum FontAlignment
    fa_center = DT_CENTER Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
    fa_top = DT_TOP Or DT_CENTER Or DT_NOPREFIX Or DT_NOCLIP
    fa_left = DT_LEFT Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
    fa_topleft = DT_TOP Or DT_LEFT Or DT_NOPREFIX Or DT_NOCLIP
    fa_bottomleft = DT_BOTTOM Or DT_LEFT Or DT_NOPREFIX Or DT_NOCLIP
    fa_bottom = DT_BOTTOM Or DT_CENTER Or DT_NOPREFIX Or DT_NOCLIP
    fa_right = DT_RIGHT Or DT_VCENTER Or DT_NOPREFIX Or DT_NOCLIP
    fa_bottomright = DT_BOTTOM Or DT_RIGHT Or DT_NOPREFIX Or DT_NOCLIP
    fa_topright = DT_TOP Or DT_RIGHT Or DT_NOPREFIX Or DT_NOCLIP
End Enum

Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Paths
Private Const PATH_GRAPHICS As String = "\graphics"
Private Const PATH_MAPS As String = "\maps"
Private Const PATH_SOUNDS As String = "\sounds"
Private Const PATH_SCRIPTS As String = "\scripts"

'***************************
'Types
'***************************

'Structure to hold all the graphics for the game
Private Type DX_Texture
    inuse As Boolean
    dimension As Long
    d3dtexture As Direct3DTexture8
End Type

'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'Holds data about where a bmp can be found,
'How big it is and animation info
Private Type Grh_Data
    texture_index As Integer
    src_x As Integer
    src_y As Integer
    src_width As Integer
    src_height As Integer
    
    frame_count As Integer
    frame_list() As Integer
    frame_speed As Single
End Type

'Points to a Grh_Data and keeps animation info
Private Type grh
    grh_index As Integer
    alpha_blend As Boolean
    angle As Single
    frame_speed As Single
    frame_counter As Single
    started As Boolean
    noloop As Boolean
    h_centered As Boolean
    v_centered As Boolean
End Type

'Char Body
Private Type Char_Data_Body
    body(1 To 8) As grh
End Type

'Char Data
Private Type Char_Data
    body_list() As Char_Data_Body
    body_count As Integer
    label_offset_x As Integer
    label_offset_y As Integer
End Type

'Hold info about a character
Private Type char
    Heading As Byte
    id As Integer
    map_x As Integer
    map_y As Integer
    
    chr_data As Char_Data
    chr_data_body_index As Integer
    
    label As String
    label_font_index As Integer
    label_offset_x As Integer
    label_offset_y As Integer
    
    Chat As String
    chat_death_time As Long
    chat_color As Long
    chat_font_index As Integer
    
    scroll_on As Boolean
    scroll_offset_counter_x As Single
    scroll_offset_counter_y As Single
    scroll_direction_x As Integer
    scroll_direction_y As Integer
End Type

Private Type Map_Exit
    exit_map_name As String
    exit_map_x As Integer
    exit_map_y As Integer
    
    c_map_x As Integer
    c_map_y As Integer
End Type

Private Type Map_NPC
    npc_data_index As Integer
    
    c_char_data_index As Integer
    c_map_x As Integer
    c_map_y As Integer
End Type

Private Type Map_Item
    item_data_index As Integer
    item_amount As Integer
    
    c_grh_index As Integer
    c_map_x As Integer
    c_map_y As Integer
End Type

'Map Tile structure
Private Type Map_Tile
    grh(1 To 4) As grh
    blocked As Boolean
    particle_group_index As Integer
    char_index As Integer
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
    
    exit_index As Integer
    npc_index As Integer
    item_index As Integer
End Type

'Map structure
Private Type Map
    map_grid() As Map_Tile
    map_x_max As Integer
    map_x_min As Integer
    map_y_max As Integer
    map_y_min As Integer
    map_description As String
    base_light_color As Long
End Type

'The light structure
Private Type Light
    id As Integer
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Particle types
Private Type Particle
    X As Integer
    Y As Integer
    vector_x As Integer
    vector_y As Integer
    angle As Single
    grh As grh
    alive_counter As Integer
End Type

Private Type Particle_Group
    id As Integer
    map_x As Integer
    map_y As Integer
    
    frame_counter As Single
    
    stream_type As Integer
    
    Particle_Stream() As Particle
    
    alive_counter As Integer
End Type

Private Type Particle_Stream
    frame_speed As Single
    
    particle_count As Integer
    
    grh_index_list() As Integer
    grh_index_count As Integer
    
    alpha_blend As Boolean
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Integer
    life2 As Integer
    friction As Single
    spin_speedL As Single
    spin_speedH As Single
    grav_strength As Integer
    bounce_strength As Integer
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    speed As Single
    life_counter As Integer
    resize_x As Integer
    resize_y As Integer
End Type

'***************************
'Variables
'***************************
'Major DX Objects
Dim DX As DirectX8
Dim d3d As Direct3D8
Dim ddevice As Direct3DDevice8
Dim d3dx As D3DX8
Dim d3dpp As D3DPRESENT_PARAMETERS
Dim d3dcaps As D3DCAPS8
Dim d3ddm As D3DDISPLAYMODE
Dim di As DirectInput8
Dim ddevice_mouse As DirectInputDevice8

'The app path
Dim resource_path As String

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long
Dim screen_width As Integer
Dim screen_height As Integer

'Map view area (where the game is played)
Dim view_screen_top As Integer 'In pixels
Dim view_screen_left As Integer 'In pixels
Dim view_screen_bottom As Integer
Dim view_screen_right As Integer
Dim view_screen_tile_width As Integer 'In tiles
Dim view_screen_tile_height As Integer 'In tiles
Dim view_screen_width As Integer
Dim view_screen_height As Integer

'Buffer area (used to draw object outside the map area but may still show up on the screen)
Dim view_tile_buffer As Integer 'In tiles

'Base tile size (smallest possible tile size: must be square)
Dim base_tile_size As Integer 'In pixels

'View position: In tiles
Dim view_pos_x As Integer
Dim view_pos_y As Integer

'Scrolling stuff
Dim scroll_on As Boolean
Dim scroll_direction_x As Integer
Dim scroll_direction_y As Integer
Dim scroll_offset_counter_x As Single
Dim scroll_offset_counter_y As Single
Dim scroll_pixels_per_frame As Integer

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim fps As Integer 'What the current frame rate is.....

'time between frames
Dim timer_elapsed_time As Single
'ticks per frame
Dim timer_ticks_per_frame As Single

'base speed for the engine
Dim engine_base_speed As Single

'true if the rendering phase was started
Dim engine_render_started As Boolean

'windowed or not windowed
Dim engine_windowed As Boolean

'show engine stats
Dim engine_show_stats As Boolean

'show blocked tiles on map
Dim engine_show_blocked_tiles As Boolean
Dim engine_show_exit_tiles As Boolean

'mouse input
Dim input_mouse_button_left As Boolean
Dim input_mouse_button_right As Boolean
Dim input_mouse_moved As Boolean
Dim input_mouse_screen_x As Integer
Dim input_mouse_screen_y As Integer
Dim input_mouse_view_x As Integer
Dim input_mouse_view_y As Integer
Dim input_mouse_map_x As Integer
Dim input_mouse_map_y As Integer
Dim input_mouse_heading As Byte

'***************************
'Arrays
'***************************
'Array of surfaces
Dim texture_list() As DX_Texture
Dim texture_count As Integer

'Grh Data Array
Dim grh_list() As Grh_Data
Dim grh_count As Integer

'Char list
Dim char_list() As char
Dim char_count As Integer
Dim char_last As Integer

'Char data list
Dim char_data_list() As Char_Data
Dim char_data_count As Integer

'Current Map
Dim map_current As Map

'Light list
Dim light_list() As Light
Dim light_count As Integer
Dim light_last As Integer

'Particle system
Dim particle_group_list() As Particle_Group
Dim particle_group_count As Integer
Dim particle_group_last As Integer
'Stream data
Dim stream_data() As Particle_Stream
Dim stream_count As Integer

'Font List
Dim font_primitive_list() As StdFont  'This contains all basic data on fonts
Dim font_list() As D3DXFont
Dim font_count As Integer
Dim font_last As Integer

Dim exit_list() As Map_Exit
Dim npc_list() As Map_NPC
Dim item_list() As Map_Item

'***************************
'External Functions
'***************************
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

'MouseInput
Private Type PointAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'For getting the display size in windowed mode
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Old fashion BitBlt functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Class_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub Class_Terminate()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 01/16/2005
'
'*****************************************************************
On Error Resume Next
    Dim i As Long
    
    'Set no texture in the device to avoid memory leaks
    If Not ddevice Is Nothing Then
        ddevice.SetTexture 0, Nothing
    End If
    
    'Destroy all textures
    For i = 1 To UBound(texture_list)
        Set texture_list(i).d3dtexture = Nothing
    Next i
    
    'Destroy all fonts
    For i = 1 To font_last
        Set font_list(i) = Nothing
        Set font_primitive_list(i) = Nothing
    Next i
    
    'Deallocate memory used for arrays
    Erase map_current.map_grid
    Erase font_list
    Erase font_primitive_list
    Erase light_list
    Erase npc_list
    Erase exit_list
    Erase item_list
    Erase char_list
    Erase char_data_list
    Erase particle_group_list
    Erase texture_list
    Erase grh_list
    Erase stream_data
    
    'Release the mouse
    ddevice_mouse.Unacquire
    
    'Detroy DX objects
    Set ddevice_mouse = Nothing
    Set di = Nothing
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set DX = Nothing
End Sub

Private Function Convert_Tile_To_View_Y(ByVal Y As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_Y = ((Y - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_Y = view_screen_top + ((Y - 1) * base_tile_size)
    End If
End Function

Private Sub Convert_Screen_To_View(ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef view_x As Integer, ByRef view_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = screen_x - view_screen_left
    view_y = screen_y - view_screen_top
End Sub

Private Sub Convert_View_To_Map(ByVal view_x As Integer, ByVal view_y As Integer, ByRef map_x As Integer, ByRef map_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim half_view_tile_width As Integer
    Dim half_view_tile_height As Integer
    
    half_view_tile_width = (view_screen_tile_width \ 2)
    half_view_tile_height = (view_screen_tile_height \ 2)
    
    map_x = (view_x \ base_tile_size)
    map_y = (view_y \ base_tile_size)
    
    If map_x > half_view_tile_width Then
        map_x = (map_x - half_view_tile_width)
    Else
        If map_x < half_view_tile_width Then
            map_x = (0 - (half_view_tile_width - map_x))
        Else
            map_x = 0
        End If
    End If
    
    If map_y > half_view_tile_height Then
        map_y = (0 - (half_view_tile_height - map_y))
    Else
        If map_y < half_view_tile_height Then
            map_y = (map_y - half_view_tile_height)
        Else
            map_y = 0
        End If
    End If
    
    map_x = view_pos_x + map_x
    map_y = view_pos_y + map_y
End Sub

Private Function Convert_View_To_Heading(ByVal view_x As Integer, ByVal view_y As Integer) As Byte
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/27/2005
'
'Cuadrants (we consider them this way, since 0 is north):
'       |
'   4   |   1
'       |
'-------+-------
'       |
'   3   |   2
'       |
'**************************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim temp As Single
    Dim radious As Single
    
    temp_x = view_x - view_screen_width / 2
    temp_y = -view_y + view_screen_height / 2
    
    'Avoid trouble at 90 and 270 degrees
    If temp_y = 0 Then temp_y = 1
    
    radious = Sqr(temp_x ^ 2 + temp_y ^ 2)
    
    'Avoid divizion by zero
    If radious = 0 Then
        Convert_View_To_Heading = 0
        Exit Function
    End If
    
    'Get angle
    temp = Abs(temp_x) / radious
    temp = Atn(temp / Sqr(-temp * temp + 1))
    
    If temp_y > 0 Then
        If temp_x < 0 Then
            'Correction for cuadrant 4
            temp = 2 * PI - temp
        End If
    Else
        If temp_x > 0 Then
            'Correction for cuadrant 2
            temp = PI - temp
        Else
            'Correction for cuadrant 3
            temp = PI + temp
        End If
    End If
    
    Convert_View_To_Heading = Round((temp - PI / 4) / (PI / 4)) + 2
    
    If Convert_View_To_Heading = 9 Then Convert_View_To_Heading = 1
End Function

Private Sub Convert_Heading_to_Direction(ByVal Heading As Byte, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addY As Integer
    Dim addX As Integer
    
    'Figure out which way to move
    Select Case Heading
        Case 1
            addY = -1
        
        Case 2
            addY = -1
            addX = 1
        
        Case 3
            addX = 1
        
        Case 4
            addX = 1
            addY = 1
        
        Case 5
            addY = 1
        
        Case 6
            addX = -1
            addY = 1
        
        Case 7
            addX = -1
        
        Case 8
            addX = -1
            addY = -1
    End Select
    
    direction_x = direction_x + addX
    direction_y = direction_y + addY
End Sub

Private Function Convert_Tile_To_View_X(ByVal X As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_X = ((X - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_X = view_screen_left + ((X - 1) * base_tile_size)
    End If
End Function

Private Sub Engine_Stats_Render()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan MArtín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim offset_x As Integer
    
    offset_x = screen_width - 110
    
    'fps
    Device_Text_Render font_list(1), fps & " FPS", 0, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
    'delay
    Device_Text_Render font_list(1), timer_ticks_per_frame & " tpf", 20, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
    
    'cursor screen pos
    Device_Text_Render font_list(1), input_mouse_screen_x & " sx, " & input_mouse_screen_y & " sy", 40, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
    
    Device_Text_Render font_list(1), input_mouse_view_x & " vx, " & input_mouse_view_y & " vy", 60, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
    
    Device_Text_Render font_list(1), input_mouse_map_x & " mx, " & input_mouse_map_y & " my", 80, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
    
    Device_Text_Render font_list(1), input_mouse_heading & " heading", 100, offset_x, 100, 20, &HFFFFFFFF, fa_topleft
End Sub

Public Sub Engine_Stats_Show_Toggle()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Toggles engine stats
'*****************************************************************
    engine_show_stats = Not engine_show_stats
End Sub

Public Sub Engine_Blocked_Tiles_Show_Toggle()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/17/2003

'*****************************************************************
    engine_show_blocked_tiles = Not engine_show_blocked_tiles
End Sub

Public Sub Engine_Special_Tiles_Show_Toggle()
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003

'*****************************************************************
    engine_show_exit_tiles = Not engine_show_exit_tiles
End Sub

Public Function Engine_View_Pos_Set(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Edited by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets the user postion. If valid, returns True, else False
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        view_pos_x = map_x
        view_pos_y = map_y
        Engine_View_Pos_Set = True
    End If
End Function

Public Sub Engine_View_Pos_Get(ByRef map_x As Integer, ByRef map_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the user postion
'*****************************************************************
    map_x = view_pos_x
    map_y = view_pos_y
End Sub

Public Sub Engine_Base_Speed_Set(ByVal speed As Single)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'*****************************************************************
    engine_base_speed = speed
End Sub

Public Function Engine_Base_Speed_Get() As Single
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_Base_Speed_Get = engine_base_speed
End Function

Public Function Engine_FPS_Get() As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_FPS_Get = fps
End Function

Public Function Engine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean, _
                                    ByVal r_path As String, Optional ByVal s_width As Integer, Optional ByVal s_height As Integer, _
                                    Optional ByVal v_left As Integer = 0, Optional ByVal v_top As Integer = 0, Optional ByVal v_width_in_tiles As Integer = 0, _
                                    Optional ByVal v_height_in_tiles As Integer = 0, Optional ByVal tile_size As Integer = 32) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    '****************************************************
    'Fill in global variables
    '****************************************************
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    'App path
    resource_path = r_path
    
    'Fill in base tile size (must always be square)
    base_tile_size = tile_size 'In pixels
    
    'Fill in viewport sizes. How large the game area is.
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 20 'In tiles
        view_screen_tile_height = 15 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    If windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
        screen_width = view_screen_width
        screen_height = view_screen_height
    Else
        screen_width = s_width
        screen_height = s_height
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
    End If
    engine_windowed = windowed
    
    '****************************************************
    'Get external data
    '****************************************************
    'Load Grh List
    Grh_Load_All
    
    'Load body data for characters
    Char_Load_Char_Data
    
    'Load Particle Stream's data
    Particle_Data_Load_All
    
    '****************************************************
    'Setup Map
    '****************************************************
    'Buffer area
    view_tile_buffer = 6
    
    'How many pixels to move per frame when scrolling
    scroll_pixels_per_frame = 4
    
    'User start position
    view_pos_x = 1
    view_pos_y = 1
    
    'Create default map
    Map_Create 50, 50
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set DX = New DirectX8
    'Create the Direct3D object
    Set d3d = DX.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    'Input object
    Set di = DX.DirectInputCreate
    
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
        If engine_windowed Then
            .windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            .BackBufferFormat = d3ddm.format 'current display depth
        Else
            .SwapEffect = D3DSWAPEFFECT_FLIP
            'Set refresh rate
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            'Turn off vsync
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            'set color depth
            .BackBufferFormat = D3DFMT_R5G6B5 '16 bit
            'Back buffer size
            .BackBufferWidth = screen_width
            .BackBufferHeight = screen_height
        End If
        .hDeviceWindow = screen_hwnd
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    'setup device
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        'Enable alphablending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
    End With
    
    '*******************************
    'Initialize input devices
    '*******************************
    Set ddevice_mouse = di.CreateDevice("guid_SysMouse")
    Call ddevice_mouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call ddevice_mouse.SetCooperativeLevel(form_hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    ' Set the buffer size
    Dim diProp As DIPROPLONG
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 1000
    Call ddevice_mouse.SetProperty("DIPROP_BUFFERSIZE", diProp)
    ddevice_mouse.Acquire
    
    '****************************************************
    'Load textures
    '****************************************************
    Texture_Load_All
    
    '****************************************************
    'Misc stuff
    '****************************************************
    'Set up default fonts
    Font_Create "Arial", 11, False, False
    
    Font_Create "Arial", 14, True, False
    
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
    
    Engine_Initialize = True
Exit Function

ErrorHandler:
    MsgBox "Error in Engine_Initialization: " & Err.Number & ": " & Err.Description
End Function

Public Function Engine_View_Move(ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim nx As Integer
    Dim ny As Integer
    
    'Don't move if we are already moving....
    If scroll_on Then Exit Function
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    X = view_pos_x
    Y = view_pos_y
    nx = X
    ny = Y
    Convert_Heading_to_Direction Heading, nx, ny
    
    'See if out new position is legal
    If Map_In_Bounds(nx, ny) Then
        'start the scrolling process
        view_pos_x = nx
        view_pos_y = ny
        
        scroll_offset_counter_x = (base_tile_size * (X - nx))
        scroll_offset_counter_y = (base_tile_size * (Y - ny))
        scroll_direction_x = nx - X
        scroll_direction_y = ny - Y
        scroll_on = True
        
        Engine_View_Move = True
    Else
        'not legal don't move
        scroll_direction_x = 0
        scroll_direction_y = 0
    End If
End Function

Public Function Engine_Render_Start() As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 4/29/2003
'
'**************************************************************
'On Error GoTo ErrorHandler
    'Exit if rendering was already started
    If engine_render_started Then Exit Function
    
    'Test Cooperative Level
    If Not Device_Test_Cooperative_Level Then Exit Function
    
    '****************************************************
    'Render
    '****************************************************
    
    '*******************************
    'get the screen_rect if windowed
    If engine_windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
    End If
    '*******************************

    '*******************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
    '*******************************
    
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    '*******************************
    'Render lights
    Light_Render_All
    '*******************************
    
    '*******************************
    'Draw Map
    Map_Render
    '*******************************
    
    engine_render_started = True
    
    Engine_Render_Start = True
Exit Function

ErrorHandler:
    Engine_Render_Start = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function Engine_Render_End() As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 4/29/2003
'
'**************************************************************
On Error GoTo ErrorHandler
    If Not engine_render_started Then Exit Function
    
    '*******************************
    'Draw engine stats
    If engine_show_stats Then Engine_Stats_Render
    '*******************************
    
    'Test Cooperative Level
    If Not Device_Test_Cooperative_Level Then
        'If we had to reset or the device is lost we restart the loop
        engine_render_started = False
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    ddevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= fps_last_time + 1000 Then
        fps = fps_frame_counter
        fps_frame_counter = 1
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    
    'Get timing info
    timer_elapsed_time = General_Get_Elapsed_Time()
    timer_ticks_per_frame = timer_elapsed_time * engine_base_speed
    
    '*******************************
    'Poll Inputs
    Input_Mouse_Poll
    '*******************************
    
    engine_render_started = False
    
    Engine_Render_End = True
Exit Function

ErrorHandler:
    Engine_Render_End = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub Device_Box_Textured_Render(ByVal texture_index As Integer, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Static src_rect As RECT         'Statc just because of performance
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height - 1
        .left = src_x
        .Right = src_x + src_width - 1
        .top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), texture_list(texture_index).dimension, angle
    
    'Set Texture
    ddevice.SetTexture 0, texture_list(texture_index).d3dtexture
    
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

Private Sub Device_Reset()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler
    
    'Destroy all fonts
    Dim loopc As Long
    For loopc = 1 To font_last
        Set font_list(loopc) = Nothing
    Next loopc
    
    font_last = 0
    
    'Reset the device
    ddevice.Reset d3dpp
    
    'Reset the device's rendering state
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        Call .SetVertexShader(FVF)
        
        'Turn off lighting
        Call .SetRenderState(D3DRS_LIGHTING, 0)
                                    
        'Set the render state that uses the alpha component as the source for blending.
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            
        'Set the render state that uses the inverse alpha component as the destination blend.
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        
        'Enable alphablending
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
    End With
    
    'Recreate all fonts from their primitives
    For loopc = 1 To UBound(font_primitive_list())
        Font_Create font_primitive_list(loopc).name, font_primitive_list(loopc).size, font_primitive_list(loopc).bold, font_primitive_list(loopc).italic
    Next loopc
Exit Sub

ErrHandler:
End Sub

Private Sub Device_Text_Render(ByVal font As D3DXFont, ByVal text As String, ByVal top As Long, ByVal left As Integer, _
                                ByVal width As Integer, ByVal height As Integer, ByVal color As Long, ByVal format As Long)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim TextRect As RECT 'This defines where it will be
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + height
    TextRect.Right = left + width
    
    d3dx.DrawText font, color, text, TextRect, format
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'**************************************************************
    With Geometry_Create_TLVertex
        .X = X
        .Y = Y
        .z = z
        .rhw = rhw
        .color = color
        .specular = specular
        .tu = tu
        .tv = tv
    End With
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_size As Long, Optional ByVal angle As Single)
'**************************************************************
'Authors: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' | \        |
' |   \      |
' |     \    |
' |       \  |
' |         \|
' * v0 *    v2
'**************************************************************
    Dim x_Cor As Single
    Dim y_Cor As Single
    
    If angle = 0 Then
        'Calculate screen coordinates of sprite
        x_Cor = dest.left
        y_Cor = dest.bottom
        
        '0 - Bottom left vertex
        If texture_size Then
            verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_size, (src.bottom + 1) / texture_size)
        Else
            verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = dest.left
        y_Cor = dest.top
        
        '1 - Top left vertex
        If texture_size Then
            verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_size, src.top / texture_size)
        Else
            verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = dest.Right
        y_Cor = dest.bottom
        
        '2 - Bottom right vertex
        If texture_size Then
            verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / texture_size, (src.bottom + 1) / texture_size)
        Else
            verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = dest.Right
        y_Cor = dest.top
        
        '3 - Top right vertex
        If texture_size Then
            verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / texture_size, src.top / texture_size)
        Else
            verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
        End If
    Else
        Dim x_center As Single
        Dim y_center As Single
        Dim radius As Single
        Dim left_point As Single
        Dim right_point As Single
        Dim temp As Single
        
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left - 1) / 2
        y_center = dest.top + (dest.bottom - dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
        
        'Calculate screen coordinates of sprite
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
        
        '0 - Bottom left vertex
        If texture_size Then
            verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_size, (src.bottom + 1) / texture_size)
        Else
            verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
        
        '1 - Top left vertex
        If texture_size Then
            verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_size, src.top / texture_size)
        Else
            verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
        
        '2 - Bottom right vertex
        If texture_size Then
            verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / texture_size, (src.bottom + 1) / texture_size)
        Else
            verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
        End If
        
        'Calculate screen coordinates of sprite
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
        
        '3 - Top right vertex
        If texture_size Then
            verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / texture_size, src.top / texture_size)
        Else
            verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 0, 0)
        End If
    End If
End Sub

Private Sub Texture_Load_From_File(ByVal texture_index As Long, ByVal file As String)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim surface_desc As D3DSURFACE_DESC
    
    Set texture_list(texture_index).d3dtexture = d3dx.CreateTextureFromFileEx(ddevice, file, D3DX_DEFAULT, _
                                                    D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                    D3DX_FILTER_POINT, D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)
    
    'Get the width and height of the texture
    texture_list(texture_index).d3dtexture.GetLevelDesc 0, surface_desc
    texture_list(texture_index).dimension = surface_desc.width - 1
End Sub

Private Sub Texture_Load_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Loads all the graphics into memory
'**************************************************************
    Dim i As Long
    Dim inipath As String
    Dim filepath As String
    
    inipath = resource_path & PATH_GRAPHICS & "\"
    
    'Get Number of Graphics
    texture_count = General_Var_Get(inipath & "grh.ini", "INIT", "texture_count")
    
    ReDim Preserve texture_list(1 To texture_count)
    
    For i = 1 To texture_count
        filepath = resource_path & PATH_GRAPHICS & "\grh" & i & ".bmp"
        If Extract_File(grh, resource_path, "grh" & i & ".bmp") Then
            If General_File_Exists(filepath, vbNormal) Then
                Texture_Load_From_File i, filepath
                Delete_File filepath
            End If
        End If
    Next i
End Sub

Private Sub Grh_Initialize(ByRef grh As grh, ByVal grh_index As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal h_centered As Boolean, Optional ByVal v_centered As Boolean)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    'Copy of parameters
    grh.grh_index = grh_index
    grh.alpha_blend = alpha_blend
    grh.angle = angle
    grh.h_centered = h_centered
    grh.v_centered = v_centered
    
    'Start it if it's a animated grh
    grh.started = (grh_list(grh.grh_index).frame_count > 1)
    
    'Set frame counters
    grh.frame_counter = 1
    grh.frame_speed = grh_list(grh.grh_index).frame_speed
End Sub

Private Sub Grh_Load_All()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'Loads Grh.dat
'**************************************************************
On Error GoTo ErrorHandler
    Dim grh As Integer
    Dim Frame As Long
    
    Dim inipath As String
    inipath = resource_path & PATH_GRAPHICS & "\"
    
    'Get number of grhs
    grh_count = Val(General_Var_Get(inipath & "grh.ini", "INIT", "grh_count"))
    
    'Resize arrays
    ReDim grh_list(1 To grh_count) As Grh_Data
    
    'Open files
    Open inipath & "grh.dat" For Binary As #1
    Seek #1, 1
    
    'Get Header
    
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , grh
    
    Do Until grh = 0
        
        'Get number of frames
        Get #1, , grh_list(grh).frame_count
        If grh_list(grh).frame_count <= 0 Then GoTo ErrorHandler
        
        'Resize array
        ReDim grh_list(grh).frame_list(1 To grh_list(grh).frame_count)
        
        If grh_list(grh).frame_count > 1 Then
            'Read a animation GRH set
            For Frame = 1 To grh_list(grh).frame_count
                Get #1, , grh_list(grh).frame_list(Frame)
                If grh_list(grh).frame_list(Frame) <= 0 Or grh_list(grh).frame_list(Frame) > grh_count Then GoTo ErrorHandler
            Next Frame
        
            Get #1, , grh_list(grh).frame_speed
            If grh_list(grh).frame_speed = 0 Then GoTo ErrorHandler
            
            'Compute width and height
            grh_list(grh).src_height = grh_list(grh_list(grh).frame_list(1)).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).src_width = grh_list(grh_list(grh).frame_list(1)).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , grh_list(grh).texture_index
            If grh_list(grh).texture_index <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_x
            If grh_list(grh).src_x < 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_y
            If grh_list(grh).src_y < 0 Then GoTo ErrorHandler
                
            Get #1, , grh_list(grh).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).frame_list(1) = grh
        End If
    
        'Get Next Grh Number
        Get #1, , grh
    Loop
    '************************************************
    
    Close #1
Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & grh
End Sub

Public Sub Grh_Add_GrhList_To_ListBox(ByRef ListboxName As ListBox)
'*****************************************************************
'Author: Fredrik Alexandersson - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 4/25/2003
'Add Grh List To Listbox
'*****************************************************************
On Local Error GoTo ErrHandler
    Dim grh As Long
    If grh_count Then
        For grh = 1 To grh_count
            If grh_list(grh).frame_count Then
                ListboxName.AddItem grh
            End If
        Next grh
    End If
Exit Sub

ErrHandler:
    MsgBox "Some sort of error", vbCritical
End Sub

Public Function Grh_Count_Get() As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 1/04/2003
'Gets the total number of grhs
'*****************************************************************
    Grh_Count_Get = grh_count
End Function

Public Sub Grh_Info_Get(ByVal grh_index As Integer, ByRef file_path As String, ByRef src_x As Integer, ByRef src_y As Integer, ByRef src_width As Integer, ByRef src_height As Integer, ByRef frame_count As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/02/2003
'Gets information about a grh
'*****************************************************************
    If Not Grh_Check(grh_index) Then Exit Sub
    
    frame_count = grh_list(grh_index).frame_count
    
    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If
    
    'Get grh file and coords
    file_path = resource_path & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
End Sub

Public Sub Grh_Info_Get_Frame(ByVal grh_index As Integer, ByVal Frame As Integer, ByRef frame_grh_index As Integer)
'*****************************************************************
'Author: Juan Martín Sotuyo dodero
'Last Modify Date: 5/02/2003
'Gets information about a grh's frame
'*****************************************************************
    If Not Grh_Check(grh_index) Then Exit Sub
    
    'Avoid subscripts out of range
    If Frame > grh_list(grh_index).frame_count Then _
        Frame = grh_list(grh_index).frame_count
    
    frame_grh_index = grh_list(grh_index).frame_list(Frame)
End Sub

Private Function Grh_Check(ByVal grh_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        Grh_Check = grh_list(grh_index).frame_count
    End If
End Function

Private Sub Grh_Uninitialize(ByRef grh As grh)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Resets a Grh
'*****************************************************************
    'Copy of parameters
    grh.grh_index = 0
    grh.alpha_blend = False
    grh.angle = 0
    grh.started = False
    'Set frame counters
    grh.frame_counter = 0
    grh.frame_speed = 0
End Sub

Private Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Integer
    
    'Animation
    If grh.started Then
        grh.frame_counter = grh.frame_counter + (timer_ticks_per_frame * grh.frame_speed)
        If grh.frame_counter > grh_list(grh.grh_index).frame_count Then
            If grh.noloop Then
                grh.frame_counter = grh_list(grh.grh_index).frame_count
            Else
                grh.frame_counter = 1
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.frame_counter = 0 Then grh.frame_counter = 1
    grh_index = grh_list(grh.grh_index).frame_list(grh.frame_counter)
    
    'Center Grh over X,Y pos
    If h_center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    If v_center Then
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_height <> 1 Then
            screen_y = screen_y - Int(tile_height * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_list(grh_index).texture_index, _
                                screen_x, screen_y, _
                                grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
                                rgb_list, _
                                grh_list(grh_index).src_x, grh_list(grh_index).src_y, _
                                grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
                                grh.alpha_blend, _
                                grh.angle
End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Integer, ByRef desthdc As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'This method is SLOW... Don't use in a loop if you care about speed!
'*************************************************************
    If Not Grh_Check(grh_index) Then Exit Sub
    
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    
    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If
    
    'Get the grh file path and coords
    file_path = resource_path & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    src_x = grh_list(grh_index).src_x
    src_y = grh_list(grh_index).src_y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
    
    'Decompress bmp
    Extract_File grh, resource_path, "grh" & grh_list(grh_index).texture_index & ".bmp"
    
    'Create a DC and load it there (we save the null pointer in it to avoid memory-leaks)
    hdcsrc = CreateCompatibleDC(desthdc)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    If Not transparent Then
        'Render
        BitBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    Else
        'Create a DC and load it there (we save the null pointer in it to avoid memory-leaks)
        MaskDC = CreateCompatibleDC(desthdc)
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
        
        'We create a mask which we will use to render transparently
        Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
        
        'Render tranparently
        BitBlt desthdc, 0, 0, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
        BitBlt desthdc, 0, 0, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
        
        'Set back original null object, and delete the loaded pic
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
        
        'Delete the DC
        DeleteDC MaskDC
    End If
    
    'Delete decompressed file
    Delete_File resource_path & PATH_GRAPHICS & "\grh" & grh_list(grh_index).texture_index & ".bmp"
    
    'Set back original null object, and delete the loaded pic
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    
    'Delete the DC
    DeleteDC hdcsrc
End Sub

Public Function Input_Mouse_Moved_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Moved_Get = input_mouse_moved
    input_mouse_moved = False
End Function

Private Sub Input_Mouse_Poll()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    'Get the mouse event buffer
    Dim DevData(1 To 1000) As DIDEVICEOBJECTDATA
    Dim nEvents As Integer
    nEvents = ddevice_mouse.GetDeviceData(DevData, DIGDD_DEFAULT)
    
    'Check buffer for clicks
    Dim i As Long
    For i = 1 To nEvents
        Select Case DevData(i).lOfs
            Case DIMOFS_BUTTON0
                'Check if mouse button was pressed or released
                input_mouse_button_left = (DevData(i).lData And &H80)
                
            Case DIMOFS_BUTTON1
                'Check if mouse button was pressed or released
                input_mouse_button_right = (DevData(i).lData And &H80)
        End Select
    Next i
    
    'Use a API to get the mouse cordinates.
    Dim temp_point As PointAPI
    GetCursorPos temp_point
    
    'See if we moved
    If input_mouse_screen_x <> temp_point.X Or input_mouse_screen_y <> temp_point.Y Then
        
        'We moved
        'Store new screen position
        input_mouse_screen_x = temp_point.X
        input_mouse_screen_y = temp_point.Y
        
        'See if we are in the view area
        Convert_Screen_To_View input_mouse_screen_x, input_mouse_screen_y, input_mouse_view_x, input_mouse_view_y
        
        input_mouse_moved = True
    End If
    
    'Figure out where on the map the mouse is
    Convert_View_To_Map input_mouse_view_x, input_mouse_view_y, input_mouse_map_x, input_mouse_map_y
    
    'Figure out heading from center of screen
    input_mouse_heading = Convert_View_To_Heading(input_mouse_view_x, input_mouse_view_y)
End Sub

Public Sub Input_Mouse_Screen_Get(ByRef screen_x As Integer, ByRef screen_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    screen_x = input_mouse_screen_x
    screen_y = input_mouse_screen_y
End Sub

Public Sub Input_Mouse_View_Get(ByRef view_x As Integer, ByRef view_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = input_mouse_view_x
    view_y = input_mouse_view_y
End Sub

Public Function Input_Mouse_In_View() As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_In_View = (input_mouse_view_x >= 0 And input_mouse_view_x < view_screen_width And input_mouse_view_y >= 0 And input_mouse_view_y < view_screen_height)
End Function

Public Function Input_Mouse_Button_Left_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Left_Get = input_mouse_button_left
End Function

Public Function Input_Mouse_Button_Right_Get() As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Button_Right_Get = input_mouse_button_right
End Function

Public Function Input_Mouse_Heading_Get() As Byte
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Mouse_Heading_Get = input_mouse_heading
End Function

Public Sub Input_Mouse_Map_Get(ByRef map_x As Integer, ByRef map_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    map_x = input_mouse_map_x
    map_y = input_mouse_map_y
End Sub

Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Input_Key_Get = (GetKeyState(key_code) < 0)
End Function

Private Function NPC_Ini_Char_Data_Index_Get(ByVal s_npc_data_index As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    NPC_Ini_Char_Data_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\npc.ini", "NPC" & s_npc_data_index, "npc_char_data_index"))
End Function

Private Function Item_Ini_Grh_Index_Get(ByVal s_item_data_index As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Item_Ini_Grh_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\item.ini", "ITEM" & s_item_data_index, "item_grh_index"))
End Function

Public Function Map_NPC_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_npc_data_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_Legal_Char_Pos(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index = 0 Then
            ReDim Preserve npc_list(0 To UBound(npc_list) + 1)
            npc_list(UBound(npc_list)).npc_data_index = s_npc_data_index
            npc_list(UBound(npc_list)).c_char_data_index = NPC_Ini_Char_Data_Index_Get(s_npc_data_index)
            npc_list(UBound(npc_list)).c_map_x = s_map_x
            npc_list(UBound(npc_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).npc_index = UBound(npc_list)
            'Create a char and plot in on the map
            Char_Create s_map_x, s_map_y, 5, npc_list(UBound(npc_list)).c_char_data_index
            Map_NPC_Add = True
        End If
    End If
End Function

Public Function Map_NPC_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_char_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_x = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).npc_index = 0
            'Get rid of the char on the map
            Char_Remove Map_Char_Get(s_map_x, s_map_y)
            Map_NPC_Remove = True
        End If
    End If
End Function

Public Function Map_NPC_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'Returns NPC's data index
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            Map_NPC_Get = npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index
        End If
    End If
End Function

Public Function Map_Exit_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_exit_map_name As String, ByVal s_exit_map_x As Integer, ByVal s_exit_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index = 0 Then
            ReDim Preserve exit_list(0 To UBound(exit_list) + 1)
            exit_list(UBound(exit_list)).exit_map_name = s_exit_map_name
            exit_list(UBound(exit_list)).exit_map_x = s_exit_map_x
            exit_list(UBound(exit_list)).exit_map_y = s_exit_map_y
            exit_list(UBound(exit_list)).c_map_x = s_map_x
            exit_list(UBound(exit_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).exit_index = UBound(exit_list)
            Map_Exit_Add = True
        End If
    End If
End Function

Public Function Map_Exit_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name = ""
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).exit_index = 0
            Map_Exit_Remove = True
        End If
    End If
End Function

Public Function Map_Exit_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByRef r_exit_map_name As String, ByRef r_exit_map_x As Integer, ByRef r_exit_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'Returns exit information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            r_exit_map_name = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name
            r_exit_map_x = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x
            r_exit_map_y = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y
            Map_Exit_Get = True
        End If
    End If
End Function

Public Function Map_Item_Add(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByVal s_item_data_index As Integer, ByVal s_item_amount As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index = 0 Then
            ReDim Preserve item_list(0 To UBound(item_list) + 1)
            item_list(UBound(item_list)).item_data_index = s_item_data_index
            item_list(UBound(item_list)).item_amount = s_item_amount
            item_list(UBound(item_list)).c_grh_index = Item_Ini_Grh_Index_Get(s_item_data_index)
            item_list(UBound(item_list)).c_map_x = s_map_x
            item_list(UBound(item_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).item_index = UBound(item_list)
            Map_Item_Add = True
        End If
    End If
End Function

Public Function Map_Item_Remove(ByVal s_map_x As Integer, ByVal s_map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index <> 0 Then
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_data_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_amount = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_grh_index = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_map_x = 0
            item_list(map_current.map_grid(s_map_x, s_map_y).item_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).item_index = 0
            Map_Item_Remove = True
        End If
    End If
End Function

Public Function Map_Item_Get(ByVal s_map_x As Integer, ByVal s_map_y As Integer, ByRef r_item_data_index As Integer, ByRef r_item_amount As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 5/14/2002
'Returns item information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index <> 0 Then
            r_item_data_index = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_data_index
            r_item_amount = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_amount
            Map_Item_Get = True
        End If
    End If
End Function

Public Sub Map_Save_Ini_To_File(ByVal file_path As String)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Dim loopc As Long
    Dim counter As Long

   'If file already exists kill it
    If General_File_Exists(file_path, vbNormal) Then
        Kill file_path
    End If
    
    General_Var_Write file_path, "GENERAL", "map_description", map_current.map_description
    
    'NPCs
    counter = 1
    If UBound(npc_list()) <> 0 Then
        For loopc = 1 To UBound(npc_list())
            If npc_list(loopc).npc_data_index Then
                General_Var_Write file_path, "NPC", CStr(counter), CStr(npc_list(loopc).c_map_x) & "-" & CStr(npc_list(loopc).c_map_y) & "-" & CStr(npc_list(loopc).npc_data_index)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "NPC", "count", CStr(counter - 1)
    
    'Exits
    counter = 1
    If UBound(exit_list()) <> 0 Then
        For loopc = 1 To UBound(exit_list())
            If exit_list(loopc).exit_map_name <> "" Then
                General_Var_Write file_path, "EXIT", CStr(counter), CStr(exit_list(loopc).c_map_x) & "-" & CStr(exit_list(loopc).c_map_y) & "-" & exit_list(loopc).exit_map_name & "-" & CStr(exit_list(loopc).exit_map_x) & "-" & CStr(exit_list(loopc).exit_map_y)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "EXIT", "count", CStr(counter - 1)
    
    'Items
    counter = 1
    If UBound(item_list()) <> 0 Then
        For loopc = 1 To UBound(item_list())
            If item_list(loopc).item_data_index Then
                General_Var_Write file_path, "ITEM", CStr(counter), CStr(item_list(loopc).c_map_x) & "-" & CStr(item_list(loopc).c_map_y) & "-" & CStr(item_list(loopc).item_data_index) & "-" & CStr(item_list(loopc).item_amount)
                counter = counter + 1
            End If
        Next loopc
    End If
    General_Var_Write file_path, "ITEM", "count", CStr(counter - 1)
End Sub

Public Sub Map_Load_Ini_From_File(ByVal file_path As String)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Dim loopc As Long
    Dim t_count As Integer
    Dim temp_string As String
    
    map_current.map_description = General_Var_Get(file_path, "GENERAL", "map_description")
        
    'NPC
    t_count = Val(General_Var_Get(file_path, "NPC", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "NPC", CStr(loopc))
        If temp_string <> "" Then
            Map_NPC_Add CLng(General_Field_Read(1, temp_string, "-")), CLng(General_Field_Read(2, temp_string, "-")), CInt(General_Field_Read(3, temp_string, "-"))
        End If
    Next loopc
    
    'Item
    t_count = Val(General_Var_Get(file_path, "ITEM", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "ITEM", CStr(loopc))
        If temp_string <> "" Then
            Map_Item_Add CLng(General_Field_Read(1, temp_string, "-")), CLng(General_Field_Read(2, temp_string, "-")), CInt(General_Field_Read(3, temp_string, "-")), CInt(General_Field_Read(4, temp_string, "-"))
        End If
    Next loopc
    
    'Exit
    t_count = Val(General_Var_Get(file_path, "EXIT", "count"))
    For loopc = 1 To t_count
        temp_string = General_Var_Get(file_path, "EXIT", CStr(loopc))
        If temp_string <> "" Then
            Map_Exit_Add CLng(General_Field_Read(1, temp_string, "-")), CLng(General_Field_Read(2, temp_string, "-")), General_Field_Read(3, temp_string, "-"), CLng(General_Field_Read(4, temp_string, "-")), CLng(General_Field_Read(5, temp_string, "-"))
        End If
    Next loopc
End Sub

Public Function Map_Save_Map(ByVal map_id As Integer, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 1/04/2003
'Save map data to files
'*****************************************************************
    Dim map_path As String
    
    'Get map file path
    map_path = resource_path & PATH_MAPS & "\map" & map_id & ".map"
    
    Map_Save_Map = Map_Save_Map_To_File(map_path, save_ini)
End Function

Public Function Map_Save_Map_To_File(ByVal file_path As String, Optional ByVal save_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/26/2004
'Save map data to files
'*****************************************************************
    Dim loopc As Long
    Dim Y As Long
    Dim X As Long
    Dim map_path As String
    Dim tile_flags As Byte
    Dim layer_flags As Byte
    Dim deco_flags As Byte
    Dim light_index As Long
    Dim particle_index As Long
    
    'Get map file path
    map_path = file_path
    
    'If file already exists kill it
    If General_File_Exists(map_path, vbNormal) Then
        Kill map_path
    End If
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Put #1, , map_current.map_x_max
    Put #1, , map_current.map_y_max
    Put #1, , map_current.base_light_color
    
    'Write .map file
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            '************************
            'Tile Flags START
            '************************
            tile_flags = 0
            
            'Layer 2
            If map_current.map_grid(X, Y).grh(2).grh_index Then tile_flags = tile_flags Or 1
            'Layer 3
            If map_current.map_grid(X, Y).grh(3).grh_index Then tile_flags = tile_flags Or 2
            'Layer 4
            If map_current.map_grid(X, Y).grh(4).grh_index Then tile_flags = tile_flags Or 4
            'Blocked
            If map_current.map_grid(X, Y).blocked Then tile_flags = tile_flags Or 8
            'Base lights
            For loopc = 0 To 3
                If map_current.map_grid(X, Y).light_base_value(loopc) <> map_current.base_light_color Then
                    tile_flags = tile_flags Or 16
                    Exit For
                End If
            Next loopc
            'Lights
            light_index = Map_Light_Get(X, Y)
            If light_index Then tile_flags = tile_flags Or 32
            'Particle Streams
            particle_index = Map_Particle_Group_Get(X, Y)
            If particle_index Then tile_flags = tile_flags Or 64
            
            '************************
            'Tile Flags END
            '************************
            
            '.map file
            
            'Store flags
            Put #1, , tile_flags
            
            'Layer 1
            Put #1, , map_current.map_grid(X, Y).grh(1).grh_index
            
            'Layers 2, 3 and 4
            For loopc = 2 To 4
                If map_current.map_grid(X, Y).grh(loopc).grh_index Then
                    '************************
                    'Layer Flags START
                    '************************
                    layer_flags = 0
                    If map_current.map_grid(X, Y).grh(loopc).alpha_blend Then layer_flags = layer_flags Or 1
                    If map_current.map_grid(X, Y).grh(loopc).h_centered Then layer_flags = layer_flags Or 2
                    If map_current.map_grid(X, Y).grh(loopc).v_centered Then layer_flags = layer_flags Or 4
                    If map_current.map_grid(X, Y).grh(loopc).angle Then layer_flags = layer_flags Or 8
                    '************************
                    'Layer Flags END
                    '************************
                    
                    'Store grh index
                    Put #1, , map_current.map_grid(X, Y).grh(loopc).grh_index
                    
                    'Store layer flags
                    Put #1, , layer_flags
                    
                    'Store angle (only if needed)
                    If map_current.map_grid(X, Y).grh(loopc).angle Then
                        Put #1, , General_Convert_Radians_To_Degrees(map_current.map_grid(X, Y).grh(loopc).angle)
                    End If
                End If
            Next loopc
            
            'Base lights
            If tile_flags And 16 Then
                For loopc = 0 To 3
                    Put #1, , map_current.map_grid(X, Y).light_base_value(loopc)
                Next loopc
            End If
            
            'Lights
            If light_index Then
                Put #1, , light_list(light_index).color
                Put #1, , light_list(light_index).range
            End If
            
            'Particle groups
            If particle_index Then
                Put #1, , particle_group_list(particle_index).stream_type
            End If
        Next X
    Next Y
    
    'Save ini if needed
    If save_ini Then
        Map_Save_Ini_To_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    'Close .map file
    Close #1
    
    Map_Save_Map_To_File = True
End Function

Public Function Map_Load_Map(ByVal map_id As Integer, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins  - Modified by Juan Martín Sotuyo dodero
'Last Modify Date: 4/25/2003
'Load map data from file using an map id number
'*****************************************************************
    Dim map_path As String

    'Get map file path
    map_path = resource_path & PATH_MAPS & "\map" & map_id & ".map"
    
    Map_Load_Map = Map_Load_Map_From_File(map_path, load_ini)
End Function

Public Function Map_Load_Map_By_Name(ByVal name As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo dodero
'Last Modify Date: 6/16/2005
'Load map data from file using a map name
'*****************************************************************
    Dim map_path As String
    
    'Get map file path
    map_path = resource_path & PATH_MAPS & "\" & name & ".map"
    
    Map_Load_Map_By_Name = Map_Load_Map_From_File(map_path, load_ini)
End Function

Public Function Map_Load_Map_From_File(ByVal file_path As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/26/2004
'Load map data from file using an filepath
'*****************************************************************
    Dim Y As Long
    Dim X As Long
    Dim map_path As String
    Dim tile_flags As Byte
    Dim layer_flags As Byte
    Dim TempInt As Integer
    Dim TempInt2 As Integer
    Dim TempLng As Long
    Dim TempByte As Byte
    Dim loopc As Long
    Dim FilaHandle As Integer
    
    'Get map file path
    map_path = file_path
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_path, vbNormal)) Then Exit Function
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    Particle_Group_Remove_All
    
    map_current.map_x_min = 1
    map_current.map_y_min = 1
    
    FilaHandle = FreeFile
    
    'Open .map file
    Open map_path For Binary As FilaHandle
    Seek #1, 1
    
    'map Header
    Get FilaHandle, , map_current.map_x_max
    Get FilaHandle, , map_current.map_y_max
    Get FilaHandle, , map_current.base_light_color
    
    'Clear out and resize map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, _
                                map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Set map's base light
    Map_Base_Light_Fill map_current.base_light_color
    
    'Read .map file
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            '.map file
            
            'Tile Flags
            Get FilaHandle, , tile_flags
            
            'Blocked
            map_current.map_grid(X, Y).blocked = (tile_flags And 8)
            
            'Layer 1
            Get FilaHandle, , TempInt
            If TempInt Then Grh_Initialize map_current.map_grid(X, Y).grh(1), TempInt
            
            'Layer 2
            If tile_flags And 1 Then
                Get FilaHandle, , TempInt
                Get FilaHandle, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get FilaHandle, , TempInt2
                Else
                    TempInt2 = 0
                End If
                
                Grh_Initialize map_current.map_grid(X, Y).grh(2), TempInt, layer_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), layer_flags And 1, layer_flags And 2
            End If
            
            'Layer 3
            If tile_flags And 2 Then
                Get FilaHandle, , TempInt
                Get FilaHandle, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get FilaHandle, , TempInt2
                Else
                    TempInt2 = 0
                End If
                
                Grh_Initialize map_current.map_grid(X, Y).grh(3), TempInt, layer_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), layer_flags And 1, layer_flags And 2
            End If
            
            'Layer 4
            If tile_flags And 4 Then
                Get FilaHandle, , TempInt
                Get FilaHandle, , layer_flags
                'Check for angle
                If layer_flags And 8 Then
                    Get FilaHandle, , TempInt2
                Else
                    TempInt2 = 0
                End If
                
                Grh_Initialize map_current.map_grid(X, Y).grh(4), TempInt, layer_flags And 1, General_Convert_Degrees_To_Radians(CSng(TempInt2)), layer_flags And 1, layer_flags And 2
            End If
            
            'Base lights
            If tile_flags And 16 Then
                For loopc = 0 To 3
                    Get FilaHandle, , TempLng
                    Map_Base_Light_Set X, Y, TempLng, loopc
                Next loopc
            End If
            
            'Lights
            If tile_flags And 32 Then
                Get FilaHandle, , TempLng
                Get FilaHandle, , TempByte
                Light_Create X, Y, TempLng, TempByte
            End If
            
            'Particle Streams
            If tile_flags And 64 Then
                Get FilaHandle, , TempInt
                Particle_Group_Create X, Y, TempInt
            End If
        Next X
    Next Y
    
    'Close .map file
    Close FilaHandle
    
    'Load Ini file
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    If load_ini Then
        Map_Load_Ini_From_File left$(map_path, Len(map_path) - 3) & "ini"
    End If
    
    Map_Load_Map_From_File = True
End Function

Public Function Map_Fill(ByVal grh_index As Integer, ByVal layer As Byte, Optional ByVal light_base_color As Long = -1, _
                        Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            'Grh
            If Not Map_Grh_Set(X, Y, grh_index, layer, alpha_blend, angle) Then Exit Function
            
            'Base light color
            If light_base_color <> -1 Then
                If Not Map_Base_Light_Set(X, Y, light_base_color) Then Exit Function
            End If
        Next X
    Next Y
    
    Map_Fill = True
End Function

Public Sub Map_Edges_Blocked_Set(ByVal edge_distance_x As Integer, ByVal edge_distance_y As Integer, ByVal blocked As Boolean)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/27/2003
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            If X <= edge_distance_x Or Y <= edge_distance_y Then
                map_current.map_grid(X, Y).blocked = blocked
            End If

            If X > map_current.map_x_max - edge_distance_x Or Y > map_current.map_y_max - edge_distance_y Then
                map_current.map_grid(X, Y).blocked = blocked
            End If
        Next X
    Next Y
End Sub

Public Sub Map_Create(ByVal map_x_max As Integer, ByVal map_y_max As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Set size
    map_current.map_x_min = 1
    map_current.map_x_max = map_x_max
    map_current.map_y_min = 1
    map_current.map_y_max = map_y_max
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    Particle_Group_Remove_All
    
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    
    'Erase map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Fill in the map with grh 1 so ther is something to render
    Map_Fill 1, 1, &HAAAAAA
    
    'Set the map's base light color
    map_current.base_light_color = &HAAAAAA
End Sub

Public Sub Map_Bounds_Get(ByRef map_x_max As Integer, ByRef map_y_max As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Get size
    map_x_max = map_current.map_x_max
    map_y_max = map_current.map_y_max
End Sub

Public Sub Map_Base_Light_Fill(ByVal light_base_color As Long)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim loop_counter As Long
    
    'Store it
    map_current.base_light_color = light_base_color
    
    For Y = map_current.map_y_min To map_current.map_y_max
        For X = map_current.map_x_min To map_current.map_x_max
            'Set all corners
            For loop_counter = 0 To 3
                map_current.map_grid(X, Y).light_base_value(loop_counter) = light_base_color
                map_current.map_grid(X, Y).light_value(loop_counter) = light_base_color
            Next loop_counter
        Next X
    Next Y
End Sub

Public Function Map_Grh_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal grh_index As Integer, _
                            ByVal layer As Byte, Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    'Check
    If layer < 1 Or layer > 5 Then Exit Function
    
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    If Not Grh_Check(grh_index) Then Exit Function
    
    'Do it
    Grh_Initialize map_current.map_grid(map_x, map_y).grh(layer), grh_index, alpha_blend, angle
    
    Map_Grh_Set = True
End Function

Public Function Map_Grh_UnSet(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim grh_index As Integer
    
    'Check
    If layer < 1 Or layer > 5 Then Exit Function
    
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    grh_index = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    If Not Grh_Check(grh_index) Then Exit Function
    
    'Do it
    Grh_Uninitialize map_current.map_grid(map_x, map_y).grh(layer)
    
    Map_Grh_UnSet = True
End Function

Public Function Map_Base_Light_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal light_base_value As Long, Optional corner As Long = -1) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
    
    'Check
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    If corner < -1 Or corner > 3 Then Exit Function
    
    'Do it
    If corner = -1 Then
        'Set all corners
        For loop_counter = 0 To 3
            map_current.map_grid(map_x, map_y).light_base_value(loop_counter) = light_base_value
            map_current.map_grid(map_x, map_y).light_value(loop_counter) = light_base_value
        Next loop_counter
    Else
        'Set just one
        map_current.map_grid(map_x, map_y).light_base_value(corner) = light_base_value
        map_current.map_grid(map_x, map_y).light_value(corner) = light_base_value
    End If
    
    Map_Base_Light_Set = True
End Function

Public Function Map_Base_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer, Optional corner As Byte = 0) As Long
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Base_Light_Get = map_current.map_grid(map_x, map_y).light_base_value(corner)
    End If
End Function

Public Function Map_In_Bounds(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Exit Function
    End If
    
    Map_In_Bounds = True
End Function

Public Function Map_Legal_Char_Pos(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    If Not Map_In_Bounds(map_x, map_y) Then Exit Function
    
    If map_current.map_grid(map_x, map_y).blocked Then Exit Function
    
    If map_current.map_grid(map_x, map_y).char_index Then Exit Function
    
    Map_Legal_Char_Pos = True
End Function

Public Function Map_Legal_Char_Pos_By_Heading(ByVal char_index As Integer, ByVal Heading As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Dim nx As Integer
        Dim ny As Integer
        nx = char_list(char_index).map_x
        ny = char_list(char_index).map_y
        Convert_Heading_to_Direction Heading, nx, ny
            
        Map_Legal_Char_Pos_By_Heading = Map_Legal_Char_Pos(nx, ny)
    End If
End Function

Public Function Map_Char_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a char_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Char_Get = map_current.map_grid(map_x, map_y).char_index
    End If
End Function

Public Function Map_Description_Get() As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    Map_Description_Get = map_current.map_description
End Function

Public Sub Map_Description_Set(ByVal s_map_description As String)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    map_current.map_description = s_map_description
End Sub

Public Function Map_Blocked_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Blocked_Get = map_current.map_grid(map_x, map_y).blocked
    Else
        Map_Blocked_Get = True
    End If
End Function

Public Function Map_Blocked_Set(ByVal map_x As Integer, ByVal map_y As Integer, ByVal blocked As Boolean) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Sets a tile position to blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        map_current.map_grid(map_x, map_y).blocked = blocked
        Map_Blocked_Set = True
    End If
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    End If
End Function

Public Function Map_Grh_Get(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Byte) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a grh and return it
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then Exit Function
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Get = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    End If
End Function

Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = light_last Then Exit Function
        loopc = loopc + 1
    Loop
    
    Map_Light_Get = loopc
Exit Function

ErrorHandler:
End Function

Private Sub Map_Render()
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim view_x As Integer
    Dim view_y As Integer
    Dim screen_x As Integer
    Dim screen_y As Integer
    
    Dim view_min_y As Integer
    Dim view_max_y As Integer
    Dim view_min_x As Integer
    Dim view_max_x As Integer
    
    Dim view_buffer_min_y As Integer
    Dim view_buffer_max_y As Integer
    Dim view_buffer_min_x As Integer
    Dim view_buffer_max_x As Integer
    
    Dim view_buffer_offset_x As Integer
    Dim view_buffer_offset_y As Integer
    
    Dim view_x_start_value As Integer
    Dim view_y_start_value As Integer
    
    Dim grh As grh
    
    Dim temp_array(3) As Long
    temp_array(0) = &HFFFFFF
    temp_array(1) = &HFFFFFF
    temp_array(2) = &HFFFFFF
    temp_array(3) = &HFFFFFF
    
    '*********************
    'Handle scrolling
    'counters
    '*********************
    If scroll_on Then
        '****** Move screen Left and Right if needed ******
        If scroll_direction_x <> 0 Then
            scroll_offset_counter_x = scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_x)
            If Sgn(scroll_offset_counter_x) = scroll_direction_x Then
                scroll_offset_counter_x = 0
                scroll_direction_x = 0
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If scroll_direction_y <> 0 Then
            scroll_offset_counter_y = scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_y)
            If Sgn(scroll_offset_counter_y) = scroll_direction_y Then
                scroll_offset_counter_y = 0
                scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If scroll_direction_x = 0 And scroll_direction_y = 0 Then
            scroll_on = False
        End If
    End If
    
    'Figure out ends and starts of view area
    view_min_y = view_pos_y - view_screen_tile_height \ 2
    view_max_y = view_pos_y + view_screen_tile_height \ 2
    view_min_x = view_pos_x - view_screen_tile_width \ 2
    view_max_x = view_pos_x + view_screen_tile_width \ 2
    
    'Add the buffer
    view_buffer_min_y = view_min_y - view_tile_buffer
    view_buffer_max_y = view_max_y + view_tile_buffer
    view_buffer_min_x = view_min_x - view_tile_buffer
    view_buffer_max_x = view_max_x + view_tile_buffer
    
    'Reduce buffer size to avoid going off limits
    If view_buffer_min_y < map_current.map_y_min Then
        view_buffer_offset_y = map_current.map_y_min - view_buffer_min_y
        view_buffer_min_y = map_current.map_y_min
    End If
    
    If view_buffer_max_y > map_current.map_y_max Then view_buffer_max_y = map_current.map_y_max
    
    If view_buffer_min_x < map_current.map_x_min Then
        view_buffer_offset_x = map_current.map_x_min - view_buffer_min_x
        view_buffer_min_x = map_current.map_x_min
    End If
    
    If view_buffer_max_x > map_current.map_x_max Then view_buffer_max_x = map_current.map_x_max
    
    view_buffer_offset_x = view_buffer_offset_x - view_tile_buffer + 1
    view_buffer_offset_y = view_buffer_offset_y - view_tile_buffer + 1
    
    'Only attempt to render layer floor beyond edges if in map bounds
    If view_min_y < map_current.map_y_min Then
        view_y_start_value = -view_min_y + 2
        view_min_y = map_current.map_y_min
    ElseIf view_min_y > map_current.map_y_min Then
        view_min_y = view_min_y - 1
    Else
        view_y_start_value = 1
    End If
    
    If view_max_y > map_current.map_y_max Then
        view_max_y = map_current.map_y_max
    ElseIf view_max_y < map_current.map_y_max Then
        view_max_y = view_max_y + 1
    End If
    
    If view_min_x < map_current.map_x_min Then
        view_x_start_value = -view_min_x + 2
        view_min_x = map_current.map_x_min
    ElseIf view_min_x > map_current.map_x_min Then
        view_min_x = view_min_x - 1
    Else
        view_x_start_value = 1
    End If
    
    If view_max_x > map_current.map_x_max Then
        view_max_x = map_current.map_x_max
    ElseIf view_max_x < map_current.map_x_max Then
        view_max_x = view_max_x + 1
    End If
    
    '*********************
    'Layer 1
    '*********************
    view_y = view_y_start_value
    For map_y = view_min_y To view_max_y
        view_x = view_x_start_value
        For map_x = view_min_x To view_max_x
            '*** Start Layer 1 ***
            If map_current.map_grid(map_x, map_y).grh(1).grh_index Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render map_current.map_grid(map_x, map_y).grh(1), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value()
            End If
            '*** End Layer 1 ***
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 2
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start Layer 2 ***
            If map_current.map_grid(map_x, map_y).grh(2).grh_index Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render map_current.map_grid(map_x, map_y).grh(2), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), map_current.map_grid(map_x, map_y).grh(2).h_centered, map_current.map_grid(map_x, map_y).grh(2).v_centered
            End If
            '*** End Layer 2 ***
            
            '*** Start Item Layer ***
            If map_current.map_grid(map_x, map_y).item_index Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                grh.grh_index = item_list(map_current.map_grid(map_x, map_y).item_index).c_grh_index
                Grh_Render grh, screen_x, screen_y, temp_array, True, True
            End If
            '*** End Item Layer ***
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Middle layer
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start Layer 3 ***
            If map_current.map_grid(map_x, map_y).grh(3).grh_index Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render map_current.map_grid(map_x, map_y).grh(3), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), map_current.map_grid(map_x, map_y).grh(3).h_centered, map_current.map_grid(map_x, map_y).grh(3).v_centered
            End If
            '*** End Layer 3 ***
            '*** Start Characters ***
            If map_current.map_grid(map_x, map_y).char_index Then
                'Figure out screen position
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Char_Render char_list(map_current.map_grid(map_x, map_y).char_index), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value
            End If
            '*** End Characters ***
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 4
    '*********************
    view_y = view_buffer_offset_y
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = view_buffer_offset_x
        For map_x = view_buffer_min_x To view_buffer_max_x
            '*** Start particle effects ***
            If map_current.map_grid(map_x, map_y).particle_group_index Then
                screen_x = (Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x) + base_tile_size / 2
                screen_y = (Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y) + base_tile_size / 2
                Particle_Group_Render map_current.map_grid(map_x, map_y).particle_group_index, screen_x, screen_y
            End If
            '*** End particle effects ***
            '*** Start Layer 4 ***
            If map_current.map_grid(map_x, map_y).grh(4).grh_index Then
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                Grh_Render map_current.map_grid(map_x, map_y).grh(4), screen_x, screen_y, map_current.map_grid(map_x, map_y).light_value(), map_current.map_grid(map_x, map_y).grh(4).h_centered, map_current.map_grid(map_x, map_y).grh(4).v_centered
            End If
            '*** End Layer 4 ***
            
            view_x = view_x + 1
        Next map_x
        
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Special Layer
    '*********************
    If engine_show_blocked_tiles Or engine_show_exit_tiles Then
        view_y = view_y_start_value
        For map_y = view_min_y To view_max_y
            view_x = view_x_start_value
            For map_x = view_min_x To view_max_x
                screen_x = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                screen_y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                
                If engine_show_blocked_tiles Then
                    '*** Start Blocked Tiles ***
                    If map_current.map_grid(map_x, map_y).blocked Then
                        grh.grh_index = 2
                        Grh_Render grh, screen_x, screen_y, temp_array
                    End If
                    '*** End Blocked Tiles ***
                End If
                
                If engine_show_exit_tiles Then
                    '*** Start Special Tiles ***
                    If map_current.map_grid(map_x, map_y).exit_index Then
                        grh.grh_index = 3
                        Grh_Render grh, screen_x, screen_y, temp_array
                    End If
                    '*** End Special Tiles ***
                End If
                
                view_x = view_x + 1
            Next map_x
            view_y = view_y + 1
        Next map_y
    End If
End Sub

Public Function Font_Create(ByVal style As String, ByVal size As Integer, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/27/2003
'Returns the font_index if successful, else 0
'**************************************************************
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    
    Font_Make Font_Create, style, bold, italic, size
Exit Function

ErrorHandler:
    Font_Create = 0
End Function

Private Sub Font_Make(ByVal font_index As Integer, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont 'We use this temporarily to setup the font
    Dim fnt As New StdFont 'This is also used to describe and setup the font
    fnt.name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    'We save the font primitive in an array, so that when we reset the device we can recreate all fonts
    If font_last = 1 And font_count = 1 Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    ElseIf font_last > UBound(font_primitive_list()) Then
        ReDim Preserve font_primitive_list(1 To font_last)
        Set font_primitive_list(font_last) = fnt
    End If
    
    Set font_desc = fnt
    Set font_list(font_index) = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Function Font_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/27/2003
'
'**************************************************************
    'check font_index
    Font_Check = (font_index > 0 And font_index <= font_last)
End Function

Public Function Char_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByVal Heading As Byte, _
                            ByVal char_data_index As Integer, Optional ByVal id As Integer = 0) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/13/2003
'Returns the char_index if successful, else 0
'**************************************************************
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    If Map_Char_Get(map_x, map_y) = 0 Then
        Char_Create = Char_Next_Open
        Char_Make Char_Create, map_x, map_y, Heading, char_data_index, 1, id
    End If
End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= char_last Then
        Char_Check = (char_list(char_index).Heading > 0)
    End If
End Function

Private Function Char_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).Heading = 0
        If loopc = char_last Then
            Char_Next_Open = char_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Next_Open = loopc
Exit Function

ErrorHandler:
    Char_Next_Open = 1
End Function

Public Function Char_Find(ByVal id As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).id = id
        If loopc = char_last Then Exit Function
        loopc = loopc + 1
    Loop
    
    Char_Find = loopc
Exit Function

ErrorHandler:
End Function

Public Function Char_Move(ByVal char_index As Integer, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    'Invalid heading
    If Heading < 1 Or Heading > 8 Then Exit Function
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Make sure it's a legal move
        temp_x = char_list(char_index).map_x
        temp_y = char_list(char_index).map_y
        Convert_Heading_to_Direction Heading, temp_x, temp_y
        If Map_In_Bounds(temp_x, temp_y) Then
            'check for another char_index
            If map_current.map_grid(temp_x, temp_y).char_index = 0 Then
                'Move it
                Char_Move_By_Head char_index, Heading
                Char_Move = True
            End If
        End If
    End If
End Function

Public Function Char_Remove(ByVal char_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Char_Check(char_index) Then
        Char_Destroy char_index
        Char_Remove = True
    End If
End Function

Public Function Char_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To char_last
        'Make sure it's a legal index
        If Char_Check(index) Then
            Char_Destroy index
        End If
    Next index
    
    Char_Remove_All = True
End Function

Private Sub Char_Make(ByVal char_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal Heading As Byte, _
                        ByVal char_data_index As Integer, ByVal body_index As Integer, ByVal id As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Makes a new character and puts it on the map
'*****************************************************************
    'Update array size
    If char_index > char_last Then
        char_last = char_index
        ReDim Preserve char_list(1 To char_last)
    End If
    char_count = char_count + 1
    
    'Heading
    char_list(char_index).Heading = Heading
    
    'Update char data
    char_list(char_index).chr_data = char_data_list(char_data_index)
    char_list(char_index).chr_data_body_index = body_index

    'Label Offset
    char_list(char_index).label_offset_x = char_data_list(char_data_index).label_offset_x
    char_list(char_index).label_offset_y = char_data_list(char_data_index).label_offset_y
    
    'Reset moving stats
    char_list(char_index).scroll_on = False
    char_list(char_index).scroll_direction_x = 0
    char_list(char_index).scroll_direction_y = 0
    char_list(char_index).scroll_offset_counter_y = 0
    char_list(char_index).scroll_offset_counter_x = 0
    
    'Update position
    char_list(char_index).map_x = map_x
    char_list(char_index).map_y = map_y
    
    'id
    char_list(char_index).id = id
    
    'Plot on map
    map_current.map_grid(map_x, map_y).char_index = char_index
End Sub

Public Function Char_Label_Set(ByVal char_index As Integer, ByVal label As String, ByVal label_font_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/27/2003
'Changes the character label
'*****************************************************************
    'Check font
    If Not Font_Check(label_font_index) Then Exit Function
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Label
        char_list(char_index).label = label
        char_list(char_index).label_font_index = label_font_index
        Char_Label_Set = True
    End If
End Function

Public Function Char_Data_Grh_Index_Get(ByVal char_data_index As Integer, ByVal s_body_index As Integer, ByVal s_heading As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/12/2003
'Returns the grh_index for a character given a body_index and heading
'*****************************************************************
    'Invalid heading
    If s_heading < 1 Or s_heading > 8 Then Exit Function
    
    'Make sure it's a legal char_index
    If char_data_index <= UBound(char_data_list) And char_data_index >= LBound(char_data_list) Then
        Char_Data_Grh_Index_Get = char_data_list(char_data_index).body_list(s_body_index).body(s_heading).grh_index
    End If
End Function

Public Function Char_Data_Set(ByVal char_index As Integer, ByVal char_data_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/28/2003
'Changes the character data
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Update char data
        char_list(char_index).chr_data = char_data_list(char_data_index)
        char_list(char_index).chr_data_body_index = 1
        
        'Label Offset
        char_list(char_index).label_offset_x = char_data_list(char_data_index).label_offset_x
        char_list(char_index).label_offset_y = char_data_list(char_data_index).label_offset_y
        
        Char_Data_Set = True
    End If
End Function

Public Function Char_Data_Body_Set(ByVal char_index As Integer, ByVal body_index As Integer, ByVal noloop As Boolean) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    Dim loopc As Long
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Update char data
        char_list(char_index).chr_data_body_index = body_index
        'Set loop status
        For loopc = 1 To 8
            char_list(char_index).chr_data.body_list(body_index).body(loopc).frame_counter = 1
            char_list(char_index).chr_data.body_list(body_index).body(loopc).noloop = noloop
        Next loopc
        
        Char_Data_Body_Set = True
    End If
End Function

Public Function Char_Heading_Set(ByVal char_index As Integer, ByVal Heading As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/20/2003
'Changes the character heading
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        char_list(char_index).Heading = Heading
        Char_Heading_Set = True
    End If
End Function

Public Function Char_Map_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Get map pos
        map_x = char_list(char_index).map_x
        map_y = char_list(char_index).map_y
        
        Char_Map_Pos_Get = True
    End If
End Function

Public Function Char_Map_Pos_Set(ByVal char_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Check map pos
        If Map_In_Bounds(map_x, map_y) Then
            'Move char
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
             char_list(char_index).map_x = map_x
             char_list(char_index).map_y = map_y
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = char_index
            Char_Map_Pos_Set = True
        End If
    End If
End Function

Private Sub Char_Move_By_Head(ByVal char_index As Integer, ByVal Heading As Byte)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim nx As Integer
    Dim ny As Integer
    
    X = char_list(char_index).map_x
    Y = char_list(char_index).map_y
    
    nx = X
    ny = Y
    Convert_Heading_to_Direction Heading, nx, ny
    
    map_current.map_grid(nx, ny).char_index = char_index
    char_list(char_index).map_x = nx
    char_list(char_index).map_y = ny
    map_current.map_grid(X, Y).char_index = 0
    
    char_list(char_index).scroll_offset_counter_x = (base_tile_size * (X - nx))
    char_list(char_index).scroll_offset_counter_y = (base_tile_size * (Y - ny))
    char_list(char_index).scroll_direction_x = nx - X
    char_list(char_index).scroll_direction_y = ny - Y
    
    char_list(char_index).scroll_on = True
    char_list(char_index).Heading = Heading
    'Set char to walk
    char_list(char_index).chr_data_body_index = 2
End Sub

Private Sub Char_Destroy(ByVal char_index As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As char
    
    map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
    char_list(char_index) = temp
    
    'Update array size
    If char_index = char_last Then
        Do Until char_list(char_last).Heading > 0
            char_last = char_last - 1
            If char_last = 0 Then
                char_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve char_list(1 To char_last)
    End If
    char_count = char_count - 1
End Sub

Private Sub Char_Load_Char_Data()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Loads Body.dat
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim LoopC3 As Long
    Dim inipath As String
    
    inipath = resource_path & PATH_GRAPHICS & "\" & "char.ini"
    
    'Get number of bodies
    char_data_count = Val(General_Var_Get(inipath, "INIT", "char_data_count"))
    
    'Resize array
    ReDim char_data_list(1 To char_data_count) As Char_Data
    
    'Fill list
    For loopc = 1 To char_data_count
        'Get list count
        char_data_list(loopc).body_count = Val(General_Var_Get(inipath, "CharData" & loopc, "body_count"))
        
        'Get label offset
        char_data_list(loopc).label_offset_x = Val(General_Var_Get(inipath, "CharData" & loopc, "label_offset_x"))
        char_data_list(loopc).label_offset_y = Val(General_Var_Get(inipath, "CharData" & loopc, "label_offset_y"))
        
        'Redim body list
        ReDim char_data_list(loopc).body_list(1 To char_data_list(loopc).body_count)
        
        'Fill body list
        For LoopC2 = 1 To char_data_list(loopc).body_count
            For LoopC3 = 1 To 8
                Grh_Initialize char_data_list(loopc).body_list(LoopC2).body(LoopC3), Val(General_Var_Get(inipath, "CharData" & loopc, LoopC2 & "_" & LoopC3))
            Next LoopC3
        Next LoopC2
    Next loopc
End Sub

Private Sub Char_Render(ByRef temp_char As char, ByVal screen_x As Long, ByVal screen_y As Long, ByRef light_value() As Long)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 12/28/2002
'Renders a character at a paticular screen point
'*****************************************************************
    Dim char_moved As Boolean
    
    'Scrolling
    If temp_char.scroll_on Then
        '****** Move Left and Right if needed ******
        If temp_char.scroll_direction_x <> 0 Then
            temp_char.scroll_offset_counter_x = temp_char.scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_x))
            If Sgn(temp_char.scroll_offset_counter_x) = temp_char.scroll_direction_x Then
                temp_char.scroll_offset_counter_x = 0
                temp_char.scroll_direction_x = 0
            End If
        End If
        '****** Move Up and Down if needed ******
        If temp_char.scroll_direction_y <> 0 Then
            temp_char.scroll_offset_counter_y = temp_char.scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * Sgn(temp_char.scroll_direction_y))
            If Sgn(temp_char.scroll_offset_counter_y) = temp_char.scroll_direction_y Then
                temp_char.scroll_offset_counter_y = 0
                temp_char.scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If temp_char.scroll_direction_x = 0 And temp_char.scroll_direction_y = 0 Then
            'Turn off scrolling
            temp_char.scroll_on = False
            'Set char to stand
            temp_char.chr_data_body_index = 1
        End If
    End If
    
    'Find screen position
    screen_x = screen_x + temp_char.scroll_offset_counter_x
    screen_y = screen_y + temp_char.scroll_offset_counter_y
    
    'Render Grh
    Grh_Render temp_char.chr_data.body_list(temp_char.chr_data_body_index).body(temp_char.Heading), screen_x, screen_y, light_value(), True, True
    
    'Do label
    If temp_char.label <> "" Then
        Device_Text_Render font_list(temp_char.label_font_index), _
                         temp_char.label, _
                         screen_y + temp_char.label_offset_y, _
                         screen_x + temp_char.label_offset_x - base_tile_size, _
                         100, 20, _
                         &HFFFFFFFF, fa_top
    End If
    
    'Do overhead chats
    If temp_char.Chat <> "" Then
        'Check life time
        If temp_char.chat_death_time <= GetTickCount Then
            temp_char.Chat = ""
            Exit Sub
        End If
        
        Device_Text_Render font_list(temp_char.chat_font_index), _
                            temp_char.Chat, _
                            screen_y - 120, _
                            screen_x - 134, _
                            300, 20, _
                            &HFFFFFFFF, fa_top
    End If
End Sub

Public Function Light_Remove(ByVal light_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Color_Value_Get = light_list(light_index).color
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal char_index
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Byte, Optional ByVal id As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        Light_Check = (light_list(light_index).range > 0)
    End If
End Function

Private Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        If light_list(loop_counter).range Then
            Light_Render loop_counter
        End If
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Integer)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim X As Long
    Dim Y As Long
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        map_current.map_grid(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        map_current.map_grid(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        map_current.map_grid(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        map_current.map_grid(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            map_current.map_grid(X, min_y).light_value(0) = color
            map_current.map_grid(X, min_y).light_value(2) = color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            map_current.map_grid(X, max_y).light_value(1) = color
            map_current.map_grid(X, max_y).light_value(3) = color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            map_current.map_grid(min_x, Y).light_value(2) = color
            map_current.map_grid(min_x, Y).light_value(3) = color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            map_current.map_grid(max_x, Y).light_value(0) = color
            map_current.map_grid(max_x, Y).light_value(1) = color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                map_current.map_grid(X, Y).light_value(0) = color
                map_current.map_grid(X, Y).light_value(1) = color
                map_current.map_grid(X, Y).light_value(2) = color
                map_current.map_grid(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).range = 0
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

Public Function Light_Find(ByVal id As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then Exit Function
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function

ErrorHandler:
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(index) Then
            Light_Destroy index
        End If
    Next index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).range > 0
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

Private Sub Light_Erase(ByVal light_index As Integer)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim loopc As Long
    Dim min_x As Long
    Dim min_y As Long
    Dim max_x As Long
    Dim max_y As Long
    Dim X As Long
    Dim Y As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        map_current.map_grid(min_x, min_y).light_value(2) = map_current.map_grid(min_x, min_y).light_base_value(2)
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        map_current.map_grid(max_x, min_y).light_value(0) = map_current.map_grid(max_x, min_y).light_base_value(0)
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        map_current.map_grid(max_x, max_y).light_value(1) = map_current.map_grid(max_x, max_y).light_base_value(1)
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        map_current.map_grid(min_x, max_y).light_value(3) = map_current.map_grid(min_x, max_y).light_base_value(3)
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            map_current.map_grid(X, min_y).light_value(0) = map_current.map_grid(X, min_y).light_base_value(0)
            map_current.map_grid(X, min_y).light_value(2) = map_current.map_grid(X, min_y).light_base_value(2)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            map_current.map_grid(X, max_y).light_value(1) = map_current.map_grid(X, max_y).light_base_value(1)
            map_current.map_grid(X, max_y).light_value(3) = map_current.map_grid(X, max_y).light_base_value(3)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            map_current.map_grid(min_x, Y).light_value(2) = map_current.map_grid(min_x, Y).light_base_value(2)
            map_current.map_grid(min_x, Y).light_value(3) = map_current.map_grid(min_x, Y).light_base_value(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            map_current.map_grid(max_x, Y).light_value(0) = map_current.map_grid(max_x, Y).light_base_value(0)
            map_current.map_grid(max_x, Y).light_value(1) = map_current.map_grid(max_x, Y).light_base_value(1)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                For loopc = 0 To 3
                    map_current.map_grid(X, Y).light_value(loopc) = map_current.map_grid(X, Y).light_base_value(loopc)
                Next loopc
            End If
        Next Y
    Next X
End Sub

Private Function Particle_Group_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Integer
    
    loopc = 1
    Do Until particle_group_list(loopc).stream_type = 0
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function

ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        Particle_Group_Check = (particle_group_list(particle_group_index).stream_type > 0)
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByVal stream_type As Integer, Optional ByVal id As Integer = 0) As Integer
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, stream_type, id
    End If
End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index
        End If
    Next index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then Exit Function
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function

ErrorHandler:
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Particle_Group
    
    map_current.map_grid(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).stream_type
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer, ByVal stream_type As Integer, ByVal id As Integer)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Map pos
    particle_group_list(particle_group_index).map_x = map_x
    particle_group_list(particle_group_index).map_y = map_y
    
    'Sets alive counter
    particle_group_list(particle_group_index).alive_counter = stream_data(stream_type).life_counter
    
    'Set particles
    ReDim particle_group_list(particle_group_index).Particle_Stream(1 To stream_data(stream_type).particle_count)
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'plot particle group on map
    map_current.map_grid(map_x, map_y).particle_group_index = particle_group_index
End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Integer, ByVal screen_x As Long, ByVal screen_y As Long)
'*****************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim no_move As Boolean
    
    If particle_group_list(particle_group_index).alive_counter Then
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        If particle_group_list(particle_group_index).frame_counter > stream_data(particle_group_list(particle_group_index).stream_type).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
        
        'If it's still alive render all the particles inside
        For loopc = 1 To stream_data(particle_group_list(particle_group_index).stream_type).particle_count
            'Render particle
            Particle_Render particle_group_index, CInt(loopc), particle_group_list(particle_group_index).stream_type, _
                            screen_x, screen_y, no_move
        Next loopc
        
        If Not no_move Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).alive_counter Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    End If
End Sub

Private Sub Particle_Render(ByVal particle_group_index As Integer, ByVal particle_index As Integer, ByVal particle_type As Integer, _
                            ByVal screen_x As Long, ByVal screen_y As Long, ByVal no_move As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/24/2003
'
'**************************************************************
    With particle_group_list(particle_group_index).Particle_Stream(particle_index)
        If Not no_move Then
            If .alive_counter = 0 Then
                'Start new particle
                Grh_Initialize .grh, stream_data(particle_type).grh_index_list(Int(General_Random_Number(1, stream_data(particle_type).grh_index_count))), stream_data(particle_type).alpha_blend
                .X = CLng(General_Random_Number(stream_data(particle_type).x1, stream_data(particle_type).x2)) - (base_tile_size / 2)
                .Y = CLng(General_Random_Number(stream_data(particle_type).y1, stream_data(particle_type).y2)) - (base_tile_size / 2)
                .vector_x = CLng(General_Random_Number(stream_data(particle_type).vecx1, stream_data(particle_type).vecx2))
                .vector_y = CLng(General_Random_Number(stream_data(particle_type).vecy1, stream_data(particle_type).vecy2))
                .angle = stream_data(particle_type).angle
                .alive_counter = CLng(General_Random_Number(stream_data(particle_type).life1, stream_data(particle_type).life2))
            Else
                'Continue old particle
                'Do gravity
                .vector_y = .vector_y + stream_data(particle_type).grav_strength
                If .Y > 0 And stream_data(particle_type).bounce_strength <> 0 Then
                    'bounce
                     .vector_y = stream_data(particle_type).bounce_strength
                End If
                'Do rotation
                .grh.angle = .grh.angle + General_Random_Number(stream_data(particle_type).spin_speedL, stream_data(particle_type).spin_speedH) / 100
                If .angle >= 2 * PI Then
                    .angle = 0
                End If
                
                'Apply vector changes
                If stream_data(particle_type).move_x2 Then .vector_x = CLng(General_Random_Number(stream_data(particle_type).move_x1, stream_data(particle_type).move_x2))
                If stream_data(particle_type).move_y2 Then .vector_y = CLng(General_Random_Number(stream_data(particle_type).move_y1, stream_data(particle_type).move_y2))
            End If
            
            'Add in vector
            .X = .X + CLng(.vector_x \ stream_data(particle_type).friction)
            .Y = .Y + CLng(.vector_y \ stream_data(particle_type).friction)
            
            'Decrement life counter
            .alive_counter = .alive_counter - 1
        End If
        
        'Draw it
        If .grh.grh_index Then
            If stream_data(particle_type).resize_x <> 0 Then
                Grh_Render_Advance .grh, .X + screen_x, .Y + screen_y, stream_data(particle_type).resize_x, stream_data(particle_type).resize_y, stream_data(particle_type).rgb_list(), False, False
            Else
                Grh_Render .grh, .X + screen_x, .Y + screen_y, stream_data(particle_type).rgb_list(), False, False
            End If
        End If
    End With
End Sub

Public Sub GUI_Box_Filled_Render(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, _
                            ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Fredrik Alexandersson - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Draws a filled box
'**************************************************************
    If Color2 = 0 Then Color2 = Color1
    If Color3 = 0 Then Color3 = Color1
    If Color4 = 0 Then Color4 = Color1
    
    Dim Vertex(3) As TLVERTEX
    Dim box_rect As RECT
    Dim rgb_list(3) As Long
    
    With box_rect
        .bottom = Y + height - 1
        .left = X
        .Right = X + width - 1
        .top = Y
    End With
    
    rgb_list(0) = Color1
    rgb_list(1) = Color2
    rgb_list(2) = Color3
    rgb_list(3) = Color4
    
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), , angle
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub GUI_Box_Outline_Render(ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, _
                            ByVal border_width As Integer, ByVal Color1 As Long, Optional ByVal Color2 As Long, Optional ByVal Color3 As Long, _
                            Optional ByVal Color4 As Long, Optional alpha_blend As Boolean)
'**************************************************************
'Author: Fredrik Alexandersson - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Draws a box outline
'**************************************************************
    If Color2 = 0 Then Color2 = Color1
    If Color3 = 0 Then Color3 = Color1
    If Color4 = 0 Then Color4 = Color1
    
    Dim VertexU(3) As TLVERTEX
    Dim VertexL(3) As TLVERTEX
    Dim VertexR(3) As TLVERTEX
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .bottom = Y + height - 1
        .left = X
        .Right = X + width - 1
        .top = Y
    End With
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    ddevice.SetTexture 0, Nothing
    'Upper Line
    VertexU(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color1, 0, 0, 0)
    VertexU(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    VertexU(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexU(0), Len(VertexU(0))
    'Left Line
    VertexL(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color2, 0, 0, 0)
    VertexL(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.bottom, 0, 1, Color2, 0, 0, 0)
    VertexL(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color1, 0, 0, 0)
    VertexL(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 2, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexL(0), Len(VertexL(0))
    'Right Border
    VertexR(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color3, 0, 0, 0)
    VertexR(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color3, 0, 0, 0)
    VertexR(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color1, 0, 0, 0)
    VertexR(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.bottom, 0, 3, Color1, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexR(0), Len(VertexR(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom - border_width, 0, 1, Color4, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.bottom, 0, 1, Color4, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub

Public Sub GUI_Text_Render(ByVal text As String, ByVal font_index As Integer, ByVal X As Integer, _
                                ByVal Y As Integer, ByVal width As Integer, _
                                ByVal height As Integer, ByVal color As Long, ByVal Alignment As FontAlignment, Optional ByVal word_wrap As Boolean)
'**************************************************************
'Author: Aaron Perkins - Modified by juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    If Not Font_Check(font_index) Then Exit Sub
    
    If word_wrap Then
        Alignment = Alignment Or DT_WORDBREAK Or DT_EDITCONTROL
    End If
    
    Device_Text_Render font_list(font_index), text, Y, X, width, height, color, Alignment
End Sub

Public Sub GUI_Grh_Render(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Color1 As Long = &HFFFFFFFF, _
                        Optional ByVal Color2 As Long = &HFFFFFFFF, Optional ByVal Color3 As Long = &HFFFFFFFF, Optional ByVal Color4 As Long = &HFFFFFFFF, _
                        Optional ByVal width As Integer, Optional ByVal height As Integer, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As grh
    Dim rgb_list(3) As Long
    
    If Not Grh_Check(grh_index) Then Exit Sub
    
    rgb_list(0) = Color1
    rgb_list(1) = Color2
    rgb_list(2) = Color3
    rgb_list(3) = Color4
    
    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    If height = 0 Then height = grh_list(temp_grh.grh_index).src_height
    If width = 0 Then width = grh_list(temp_grh.grh_index).src_width
    
    Grh_Render_Advance temp_grh, X, Y, width, height, rgb_list
End Sub

Public Sub GUI_Grh_Animated_Render(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef Frame As Single, ByVal width As Integer, ByVal height As Integer, _
                        Optional ByVal Color1 As Long = &HFFFFFFFF, Optional ByVal Color2 As Long = &HFFFFFFFF, Optional ByVal Color3 As Long = &HFFFFFFFF, _
                        Optional ByVal Color4 As Long = &HFFFFFFFF, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As grh
    Dim rgb_list(3) As Long
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    If Not Grh_Check(grh_index) Then Exit Sub
    
    rgb_list(0) = Color1
    rgb_list(1) = Color2
    rgb_list(2) = Color3
    rgb_list(3) = Color4
    
    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    'Get proper x, y corrds for each frame
    temp_x = X + width / 2
    temp_y = Y + height / 2
    
    'Set height and width according to grh size
    height = grh_list(grh_list(temp_grh.grh_index).frame_list(Frame)).src_height * (height / grh_list(grh_list(temp_grh.grh_index).frame_list((grh_list(temp_grh.grh_index).frame_count))).src_height)
    width = grh_list(grh_list(temp_grh.grh_index).frame_list(Frame)).src_width * (width / grh_list(grh_list(temp_grh.grh_index).frame_list((grh_list(temp_grh.grh_index).frame_count))).src_width)
    
    'Apply correction to x, y coords
    temp_x = temp_x - width / 2
    temp_y = temp_y - height / 2 + 10
    
    temp_grh.frame_counter = Frame
    
    Grh_Render_Advance temp_grh, temp_x, temp_y, width, height, rgb_list
    
    Frame = temp_grh.frame_counter
End Sub

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
'**************************************************************
    Dim X As Long
    Dim Y As Long
    Dim TransColor As Long
    Dim ColorKey As String
    
    ColorKey = Hex(COLOR_KEY)
    
    'Check if it has an alpha component
    If Len(ColorKey) > 6 Then
        'get rid of alpha
        ColorKey = "&H" & Right$(ColorKey, 6)
    End If
    
    TransColor = Val(ColorKey)
    
    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height - src_y
        For X = src_x To src_width - src_x
            If GetPixel(hdcsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hdcsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If
        Next X
    Next Y
End Sub

Private Function Device_Test_Cooperative_Level() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Handle Alt-Tab and Ctrl-Alt-Del
'**************************************************************
    'Call TestCooperativeLevel to see what state the device is in.
    Dim hr As Long
    hr = ddevice.TestCooperativeLevel
    If hr = D3DERR_DEVICELOST Then
        'Make sure all fonts react to the lsot device
        For hr = 1 To font_count
            font_list(hr).OnLostDevice
        Next hr
        'If the device is lost, exit and wait for it to come back.
        Exit Function
    ElseIf hr = D3DERR_DEVICENOTRESET Then
        'The device became lost for some reason (Alt-Tab or Ctrl-Alt-Del) and now
        'Reset() needs to be called to try and get the device back.
        Device_Reset
        'Return false so that the rendering loop will start over
        Exit Function
    End If
    
    Device_Test_Cooperative_Level = True
End Function

Public Function Light_Move_By_Head(ByVal light_index As Integer, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim nx As Integer
    Dim ny As Integer
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then Exit Function

    'Make sure it's a legal char_index
    If Light_Check(light_index) Then
        nx = light_list(light_index).map_x
        ny = light_list(light_index).map_y
        
        Convert_Heading_to_Direction Heading, nx, ny
        
        'Make sure it's a legal move
        If Map_In_Bounds(nx, ny) Then
            'Move it
            Light_Erase light_index
            
            light_list(light_index).map_x = nx
            light_list(light_index).map_y = ny
            
            Light_Move_By_Head = True
        End If
    End If
End Function

Private Sub Particle_Data_Load_All()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modified: 12/08/03
'Loads the Particles.ini file
'**************************************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim StreamFile As String

    StreamFile = resource_path & "\Particles.ini"
    
    stream_count = Val(General_Var_Get(StreamFile, "INIT", "stream_count"))
    
    'Resize Stream_Data array
    ReDim stream_data(1 To stream_count) As Particle_Stream
    
    'Fill Stream_Data array
    For loopc = 1 To stream_count
        stream_data(loopc).particle_count = Val(General_Var_Get(StreamFile, loopc, "particle_count"))
        stream_data(loopc).x1 = Val(General_Var_Get(StreamFile, loopc, "X1"))
        stream_data(loopc).y1 = Val(General_Var_Get(StreamFile, loopc, "Y1"))
        stream_data(loopc).x2 = Val(General_Var_Get(StreamFile, loopc, "X2"))
        stream_data(loopc).y2 = Val(General_Var_Get(StreamFile, loopc, "Y2"))
        stream_data(loopc).angle = Val(General_Var_Get(StreamFile, loopc, "Angle"))
        stream_data(loopc).vecx1 = Val(General_Var_Get(StreamFile, loopc, "VecX1"))
        stream_data(loopc).vecx2 = Val(General_Var_Get(StreamFile, loopc, "VecX2"))
        stream_data(loopc).vecy1 = Val(General_Var_Get(StreamFile, loopc, "VecY1"))
        stream_data(loopc).vecy2 = Val(General_Var_Get(StreamFile, loopc, "VecY2"))
        stream_data(loopc).life1 = Val(General_Var_Get(StreamFile, loopc, "Life1"))
        stream_data(loopc).life2 = Val(General_Var_Get(StreamFile, loopc, "Life2"))
        stream_data(loopc).friction = Val(General_Var_Get(StreamFile, loopc, "Friction"))
        stream_data(loopc).spin_speedL = Val(General_Var_Get(StreamFile, loopc, "Spin_SpeedL"))
        stream_data(loopc).spin_speedH = Val(General_Var_Get(StreamFile, loopc, "Spin_SpeedH"))
        stream_data(loopc).alpha_blend = CBool(General_Var_Get(StreamFile, loopc, "AlphaBlend"))
        stream_data(loopc).grav_strength = Val(General_Var_Get(StreamFile, loopc, "Grav_Strength"))
        stream_data(loopc).bounce_strength = Val(General_Var_Get(StreamFile, loopc, "Bounce_Strength"))
        stream_data(loopc).move_x1 = Val(General_Var_Get(StreamFile, loopc, "move_x1"))
        stream_data(loopc).move_x2 = Val(General_Var_Get(StreamFile, loopc, "move_x2"))
        stream_data(loopc).move_y1 = Val(General_Var_Get(StreamFile, loopc, "move_y1"))
        stream_data(loopc).move_y2 = Val(General_Var_Get(StreamFile, loopc, "move_y2"))
        stream_data(loopc).life_counter = Val(General_Var_Get(StreamFile, loopc, "life_counter"))
        stream_data(loopc).speed = Val(General_Var_Get(StreamFile, loopc, "Speed"))
        stream_data(loopc).resize_x = Val(General_Var_Get(StreamFile, loopc, "resize_x"))
        stream_data(loopc).resize_y = Val(General_Var_Get(StreamFile, loopc, "resize_y"))
        
        stream_data(loopc).grh_index_count = Val(General_Var_Get(StreamFile, loopc, "grh_count"))
        
        ReDim stream_data(loopc).grh_index_list(1 To stream_data(loopc).grh_index_count)
        GrhListing = General_Var_Get(StreamFile, loopc, "Grh_List")
        
        For i = 1 To stream_data(loopc).grh_index_count
            stream_data(loopc).grh_index_list(i) = Val(General_Field_Read(i, GrhListing, ","))
        Next i
        
        For i = 1 To 4
            stream_data(loopc).rgb_list(i - 1) = Val(General_Var_Get(StreamFile, loopc, "ColorSet" & i))
        Next i
    Next loopc
End Sub

Public Function Engine_Resolution_Change(ByVal windowed As Boolean, Optional ByVal s_width As Integer, Optional ByVal s_height As Integer, _
                                        Optional ByVal v_left As Integer = 0, Optional ByVal v_top As Integer = 0, Optional ByVal v_width_in_tiles As Integer = 0, _
                                        Optional ByVal v_height_in_tiles As Integer = 0) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/29/2004
'Changes the resolution while in-game
'*****************************************************************
On Error GoTo ErrHandler
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 20 'In tiles
        view_screen_tile_height = 15 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    
    If windowed Then
        Dim target As RECT
        GetWindowRect screen_hwnd, target
        view_screen_top = target.top
        view_screen_left = target.left
        view_screen_right = target.Right
        view_screen_bottom = target.bottom
        view_screen_width = target.Right - target.left
        view_screen_height = target.bottom - target.top
        screen_width = view_screen_width
        screen_height = view_screen_height
        engine_windowed = True
    Else
        view_screen_left = v_left 'In pixels
        view_screen_top = v_top 'In pixels
        view_screen_width = view_screen_tile_width * base_tile_size
        view_screen_height = view_screen_tile_height * base_tile_size
        view_screen_right = view_screen_left + view_screen_width - 1
        view_screen_bottom = view_screen_top + view_screen_height - 1
        screen_width = s_width
        screen_height = s_height
        engine_windowed = False
    End If
    
    Dim d3dpp2 As D3DPRESENT_PARAMETERS
    
    With d3dpp2
        If engine_windowed Then
            .windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            .BackBufferFormat = d3ddm.format 'current display depth
        Else
            .SwapEffect = D3DSWAPEFFECT_FLIP
            'Set refresh rate
            .FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
            'Turn off vsync
            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
            'set color depth
            .BackBufferFormat = D3DFMT_R5G6B5 '16 bit
            'Back buffer size
            .BackBufferWidth = screen_width
            .BackBufferHeight = screen_height
        End If
        .hDeviceWindow = screen_hwnd
    End With
    
    d3dpp = d3dpp2
    
    Device_Reset
    
    Engine_Resolution_Change = True
Exit Function

ErrHandler:
    MsgBox "Couldn't change reolution."
End Function

Private Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal width As Integer, ByVal height As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Integer
    
    'Animation
    If grh.started Then
        grh.frame_counter = grh.frame_counter + (timer_ticks_per_frame * grh.frame_speed)
        If grh.frame_counter > grh_list(grh.grh_index).frame_count Then
            If grh.noloop Then
                grh.frame_counter = grh_list(grh.grh_index).frame_count
            Else
                grh.frame_counter = 1
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.frame_counter = 0 Then grh.frame_counter = 1
    grh_index = grh_list(grh.grh_index).frame_list(grh.frame_counter)
    
    'Center Grh over X,Y pos
    If h_center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
    End If
    If v_center Then
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_height <> 1 Then
            screen_y = screen_y - Int(tile_height * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_list(grh_index).texture_index, _
                                screen_x, screen_y, _
                                grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
                                rgb_list, _
                                grh_list(grh_index).src_x, grh_list(grh_index).src_y, _
                                width, height, _
                                grh.alpha_blend, _
                                grh.angle
End Sub

Public Function Char_Chat_Set(ByVal char_index As Integer, ByVal Chat As String, ByVal font_color As Long, ByVal font_index As Integer) As Boolean
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 4/22/2004
'Sets a text as chat over head to the given char index
'*****************************************************************
    If Char_Check(char_index) Then
        char_list(char_index).Chat = Chat
        char_list(char_index).chat_color = font_color
        char_list(char_index).chat_font_index = font_index
        'Give 10 secs of life to message
        char_list(char_index).chat_death_time = GetTickCount + 10000
        
        Char_Chat_Set = True
    End If
End Function

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function

Public Function Map_Resize(ByVal max_x As Long, ByVal max_y As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/29/2003
'Changes the map´s size (to be used by the Map Editor)
'**************************************************************
    Dim old_max_x As Long
    Dim old_max_y As Long
    Dim map_x As Long
    Dim map_y As Long
    Dim loopc As Long
    Dim TempMap As Map
    Dim X As Long
    Dim Y As Long
    
    'Get old map size
    old_max_x = map_current.map_x_max
    old_max_y = map_current.map_y_max
    
    'Check if map was shrinked
    If max_x < old_max_x Then
        For map_x = old_max_x To max_x
            For map_y = 1 To max_y
                'Remove Particle streams
                If map_current.map_grid(map_x, map_y).particle_group_index > 0 Then
                    Particle_Group_Remove map_current.map_grid(map_x, map_y).particle_group_index
                End If
                'Remove exits
                Map_Exit_Remove map_x, map_y
            Next map_y
        Next map_x
    End If
    If max_y < old_max_y Then
        For map_y = old_max_y To max_x
            For map_x = 1 To max_x
                'Remove Particle streams
                If map_current.map_grid(map_x, map_y).particle_group_index > 0 Then
                    Particle_Group_Remove map_current.map_grid(map_x, map_y).particle_group_index
                End If
                'Remove exits
                Map_Exit_Remove map_x, map_y
            Next map_x
        Next map_y
    End If
    'Remove lights which will be left outside of the map
    If max_y < old_max_y Or max_x < old_max_x Then
        loopc = light_last
        If light_last > 0 Then
            Do Until loopc = 0
                If light_list(loopc).map_x > max_x Or light_list(loopc).map_y > max_y Then
                    Light_Remove loopc
                End If
            Loop
        End If
    End If
    
    'Set new map size
    'Store data to a temp map
    TempMap.map_x_min = 1
    TempMap.map_y_min = 1
    TempMap.map_x_max = max_x
    TempMap.map_y_max = max_y
    
    ReDim TempMap.map_grid(1 To max_x, 1 To max_y)
    
    TempMap.map_description = map_current.map_description
    
    If max_x > old_max_x Then
        X = old_max_x
    Else
        X = max_x
    End If
    If max_y > old_max_y Then
        Y = old_max_y
    Else
        Y = max_y
    End If
    For map_x = 1 To X
        For map_y = 1 To Y
            TempMap.map_grid(map_x, map_y) = map_current.map_grid(map_x, map_y)
        Next map_y
    Next map_x
    
    'Resize map and copy evertyhing back
    map_current.map_x_min = 1
    map_current.map_y_min = 1
    map_current.map_x_max = max_x
    map_current.map_y_max = max_y
    
    map_current.map_description = TempMap.map_description
    
    ReDim map_current.map_grid(1 To max_x, 1 To max_y)
    
    For map_x = 1 To X
        For map_y = 1 To Y
            map_current.map_grid(map_x, map_y) = TempMap.map_grid(map_x, map_y)
        Next map_y
    Next map_x
    
    'If map was enlarged set all new tiles new info
    If max_x > old_max_x Then
        For map_x = old_max_x + 1 To max_x
            For map_y = 1 To max_y
                Grh_Initialize map_current.map_grid(map_x, map_y).grh(1), 1
                Map_Base_Light_Set map_x, map_y, &HAAAAAA
            Next map_y
        Next map_x
    End If
    If max_y > old_max_y Then
        For map_y = old_max_y + 1 To max_x
            For map_x = 1 To max_x
                Grh_Initialize map_current.map_grid(map_x, map_y).grh(1), 1
                Map_Base_Light_Set map_x, map_y, &HAAAAAA
            Next map_x
        Next map_y
    End If
End Function

Public Function Particle_Group_Move(ByVal particle_index As Integer, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 01/19/2005
'Moves a particle group
'**************************************************************
    If Not Particle_Group_Check(particle_index) Then Exit Function
    
    'Make sure it's a legal move
    If Map_In_Bounds(map_x, map_y) Then
        If map_current.map_grid(map_x, map_y).particle_group_index = 0 Then
            'Move it
            map_current.map_grid(map_x, map_y).particle_group_index = particle_index
            map_current.map_grid(particle_group_list(particle_index).map_x, particle_group_list(particle_index).map_y).particle_group_index = 0
            
            particle_group_list(particle_index).map_x = map_x
            particle_group_list(particle_index).map_y = map_y
            
            Particle_Group_Move = True
        End If
    End If
End Function

Public Function Char_Name_Get(ByVal char_index As Integer) As String
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 03/25/2005
'
'**************************************************************
    If Not Char_Check(char_index) Then _
        Exit Function
    
    Char_Name_Get = char_list(char_index).label
End Function
