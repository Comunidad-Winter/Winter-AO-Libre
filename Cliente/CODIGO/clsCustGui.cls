VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsGUI - Generates windows and controls using DX8 - v0.5.0
'
'Custom GUI
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
'01/19/2005 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com) and Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'   - First release

Option Explicit

Const CHECKBOX_WIDTH As Byte = 15   'In pixels
Const CHECKBOX_HEIGHT As Byte = 15

Const OPTIONBUTTON_WIDTH As Byte = 10   'In pixels
Const OPTIONBUTTON_HEIGHT As Byte = 10

Private Type GUIOptionButton
    text As String
    font_index As Integer
    font_color As Long
    top As Integer
    left As Integer
    width As Integer
    height As Integer
    value As Boolean
    grh_unselected As Integer
    grh_selected As Integer
    grh_disabled As Integer
    Background_Color As Long
    visible As Boolean
    enabled As Boolean
    group As Integer
    switch_id As Integer
End Type

Private Type GUICheckBox
    text As String
    font_index As Integer
    font_color As Long
    top As Integer
    left As Integer
    width As Integer
    height As Integer
    value As Boolean
    grh_unchecked As Integer
    grh_checked As Integer
    Background_Color As Long
    visible As Boolean
    enabled As Boolean
    'GUMP specyfic
    switch_id As Integer
End Type

Private Type GUITextBox
    text As String
    top As Integer
    left As Integer
    width As Integer
    height As Integer
    Background_Color As Long
    FontColor As Long
    FontColor2 As Long
    background_grh As Integer
    font_index As Integer
    visible As Boolean
    password As Boolean
End Type

Private Type GUIButton
    text As String
    top As Integer
    left As Integer
    width As Integer
    height As Integer
    font_color As Long
    Background_Color As Long
    Grh_Unpressed As Integer
    Grh_Pressed As Integer
    visible As Boolean
    enabled As Boolean
    'GUMP specyfic
    GUMP_type As Integer
    button_id As Integer
End Type

Private Type GUIImage
    left As Integer
    top As Integer
    grh As Integer
    height As Integer
    width As Integer
End Type

Private Type GUILabel
    text As String
    top As Integer
    left As Integer
    height As Integer
    width As Integer
    font As Integer
    font_color As Long
    Background_Color As Long
    Alignment As FontAlignment
End Type

Private Type GUIWindow
    grhs(1 To 9) As Integer
    top As Integer
    left As Integer
    width As Integer
    height As Integer
    buttoncount As Integer
    textscount As Integer
    labelcount As Integer
    checkboxcount As Integer
    optionbuttoncount As Integer
    imagecount As Integer
    buttons() As GUIButton
    texts() As GUITextBox
    labels() As GUILabel
    checkboxes() As GUICheckBox
    optionbuttons() As GUIOptionButton
    images() As GUIImage
    visible As Boolean
    Moveable As Boolean
    Resizeable As Boolean
    Closeable As Boolean
    'For animated grhs
    animate As Boolean
    Frame As Single
    loops As Integer
    'GUMP specyfic
    GUMP As Boolean
    grh_x(1 To 9) As Integer
    grh_y(1 To 9) As Integer
    grh_width(1 To 9) As Integer
    grh_height(1 To 9) As Integer
    grh_alphablending(1 To 9) As Boolean
    handle As Integer
End Type

Private TextWindow As Integer
Private TextActive As Integer

Private ClickedControlwindow As Integer    'Used with the one below
Private ButtonClicked As Integer

Private WindowPressed As Integer

Private WindowsCount As Integer

Private Window_Active As Integer

Private Windows() As GUIWindow

Private ZOrder() As Integer

Private Last_Handle As Integer

Private Engine As clsTileEngineX

Public Function Text_Get_Active() As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Text_Get_Active = TextActive
End Function

Public Function Text_Active_Get_Window() As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Text_Active_Get_Window = TextWindow
End Function

Public Function Text_Set_Active(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        TextWindow = WindowID
        TextActive = id
        Text_Set_Active = True
    End If
End Function

Public Function Windows_Get_Count() As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Windows_Get_Count = WindowsCount + 1
End Function

Public Function Image_Create(ByVal WindowID As Integer, ByVal left As Integer, ByVal top As Integer, ByVal grh As Integer, ByVal height As Integer, ByVal width As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).imagecount = Windows(WindowID).imagecount + 1
        ReDim Preserve Windows(WindowID).images(0 To Windows(WindowID).imagecount)
        With Windows(WindowID).images(Windows(WindowID).imagecount)
            .grh = grh
            .left = left
            .top = top
            .height = height
            .width = width
        End With
        
        Image_Create = Windows(WindowID).imagecount
    Else
        Image_Create = -1
    End If
End Function

Private Function Image_Check(ByVal WindowID As Integer, ByVal ImageID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Image_Check = (ImageID <= Windows(WindowID).imagecount And ImageID > -1)
    End If
End Function

Public Function Image_Get_Grh(ByVal WindowID As Integer, ByVal ImageID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Grh = Windows(WindowID).images(ImageID).grh
    End If
End Function

Public Function Image_Get_Top(ByVal WindowID As Integer, ByVal ImageID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Top = Windows(WindowID).images(ImageID).top
    End If
End Function

Public Function Image_Get_Left(ByVal WindowID As Integer, ByVal ImageID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Left = Windows(WindowID).images(ImageID).left
    End If
End Function

Public Function Image_Get_Width(ByVal WindowID As Integer, ByVal ImageID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Width = Windows(WindowID).images(ImageID).width
    End If
End Function

Public Function Image_Get_Height(ByVal WindowID As Integer, ByVal ImageID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Image_Check(WindowID, ImageID) Then
        Image_Get_Height = Windows(WindowID).images(ImageID).height
    End If
End Function

Public Function Text_Create(ByVal grh As Integer, ByVal WindowID As Integer, ByVal text As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, ByVal font As Integer, _
                            ByVal FontColor As Long, Optional ByVal FontColor2 As Long = -1, Optional ByVal BackgroundColor As Long = -1, Optional ByVal password As Boolean) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).textscount = Windows(WindowID).textscount + 1
        ReDim Preserve Windows(WindowID).texts(0 To Windows(WindowID).textscount)
        With Windows(WindowID).texts(Windows(WindowID).textscount)
            .background_grh = grh
            .text = text
            .top = top
            .left = left
            .width = width
            .height = height
            .visible = True
            .font_index = font
            .FontColor = FontColor
            .FontColor2 = FontColor2
            .Background_Color = BackgroundColor
            .password = password
        End With
        
        Text_Create = Windows(WindowID).textscount
    Else
        Text_Create = -1
    End If
End Function

Public Function Window_Create(ByVal grh As Integer, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, Optional ByVal Moveable As Boolean = False, _
                            Optional Resizeable As Boolean = False, Optional ByVal Closeable As Boolean = False, Optional ByVal GUMP As Boolean = False, Optional ByVal animate As Boolean = False, Optional ByVal loops As Integer = -1) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    WindowsCount = WindowsCount + 1
    ReDim Preserve Windows(0 To WindowsCount) As GUIWindow
    With Windows(WindowsCount)
        .top = top
        .left = left
        .width = width
        .height = height
        .buttoncount = -1
        .labelcount = -1
        .textscount = -1
        .checkboxcount = -1
        .optionbuttoncount = -1
        .imagecount = -1
        .visible = False
        .Moveable = Moveable
        .Resizeable = Resizeable
        .Closeable = .Closeable
        .animate = animate
        .Frame = 1
        .loops = loops
        .GUMP = GUMP
        If GUMP Then
            .grhs(1) = grh
            .grhs(2) = grh + 1
            .grhs(3) = grh + 2
            .grhs(4) = grh + 3
            .grhs(5) = grh + 4
            .grhs(6) = grh + 5
            .grhs(7) = grh + 6
            .grhs(8) = grh + 7
            .grhs(9) = grh + 8
            
            'Calculate the size of each grh
            Dim h1 As Integer
            Dim h2 As Integer
            Dim h3 As Integer
            Dim h4 As Integer
            Dim h5 As Integer
            Dim h6 As Integer
            Dim h7 As Integer
            Dim h8 As Integer
            Dim h9 As Integer
            Dim w1 As Integer
            Dim w2 As Integer
            Dim w3 As Integer
            Dim w4 As Integer
            Dim w5 As Integer
            Dim w6 As Integer
            Dim w7 As Integer
            Dim w8 As Integer
            Dim w9 As Integer
            Dim temp_str As String
            Dim temp As Integer
            Dim x_factor As Single
            Dim y_factor As Single
            
            'Get height and width of all grhs
            Engine.Grh_Info_Get grh, temp_str, temp, temp, w1, h1, temp
            Engine.Grh_Info_Get grh + 1, temp_str, temp, temp, w2, h2, temp
            Engine.Grh_Info_Get grh + 2, temp_str, temp, temp, w3, h3, temp
            Engine.Grh_Info_Get grh + 3, temp_str, temp, temp, w4, h4, temp
            Engine.Grh_Info_Get grh + 4, temp_str, temp, temp, w5, h5, temp
            Engine.Grh_Info_Get grh + 5, temp_str, temp, temp, w6, h6, temp
            Engine.Grh_Info_Get grh + 6, temp_str, temp, temp, w7, h7, temp
            Engine.Grh_Info_Get grh + 7, temp_str, temp, temp, w8, h8, temp
            Engine.Grh_Info_Get grh + 8, temp_str, temp, temp, w9, h9, temp
            
            'Calculate stretching factor
            x_factor = width / (w1 + w2 + w3)
            y_factor = height / (h1 + h2 + h3)
            
            'Set x and width for all grhs
            .grh_x(1) = .grh_x(4) = .grh_x(7) = left
            
            'First row
            .grh_width(1) = .grh_x(2) = w1 * x_factor
            .grh_width(2) = .grh_x(3) = w2 * x_factor
            .grh_width(3) = w3 * x_factor
            
            'Second row
            .grh_width(4) = .grh_x(5) = w4 * x_factor
            .grh_width(5) = .grh_x(6) = w5 * x_factor
            .grh_width(6) = w6 * x_factor
            
            'Third row
            .grh_width(7) = .grh_x(8) = w7 * x_factor
            .grh_width(8) = .grh_x(9) = w8 * x_factor
            .grh_width(9) = w9 * x_factor
            
            'Set y and height for all grhs
            .grh_y(1) = .grh_y(2) = .grh_y(3) = top
            'First column
            .grh_height(1) = .grh_y(4) = h1 * y_factor
            .grh_height(4) = .grh_y(7) = h4 * y_factor
            .grh_height(7) = h7 * y_factor
            
            'Second column
            .grh_height(2) = .grh_y(5) = h2 * y_factor
            .grh_height(5) = .grh_y(8) = h5 * y_factor
            .grh_height(8) = h8 * y_factor
            
            'Third column
            .grh_height(3) = .grh_y(3) = h3 * y_factor
            .grh_height(6) = .grh_y(6) = h6 * y_factor
            .grh_height(9) = h9 * y_factor
        Else
            .grhs(1) = grh
        End If
        
        'Update handles and set it to the window
        Last_Handle = Last_Handle + 1
        .handle = Last_Handle
    End With
    Window_Create = Last_Handle
    
    'It´s set on top
    Dim loopc As Long
    
    ReDim Preserve ZOrder(WindowsCount)
    For loopc = WindowsCount - 1 To 0 Step -1
        ZOrder(loopc + 1) = ZOrder(loopc)
    Next loopc
    
    ZOrder(0) = Last_Handle
End Function

Public Function Window_Hit_Test(ByVal X As Integer, ByVal Y As Integer, Optional ByVal right_click As Boolean = False) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    If WindowsCount = -1 Then
        Window_Hit_Test = -1
        Exit Function
    End If
    
    'We check in Z-Order
    For loopc = 0 To UBound(ZOrder)
        If Windows(ZOrder(loopc)).visible = True Then
            If Windows(ZOrder(loopc)).top < Y And Windows(ZOrder(loopc)).height + Windows(ZOrder(loopc)).top > Y Then
                If Windows(ZOrder(loopc)).left < X And Windows(ZOrder(loopc)).left + Windows(ZOrder(loopc)).width > X Then
                    'If we right clicked, and it's closeable, then we close it
                    If right_click Then
                        If Windows(ZOrder(loopc)).Closeable Then
                            'Hide the window
                            Window_Show ZOrder(loopc), False
                            Window_Active = -1
                            Window_Hit_Test = -1
                        End If
                        Exit Function
                    End If
                    Window_Active = ZOrder(loopc)
                    Window_Hit_Test = ZOrder(loopc)
                    Window_Set_On_Top ZOrder(loopc)
                    Exit Function
                End If
            End If
        End If
    Next loopc
    
    Window_Active = -1
    Window_Hit_Test = -1
End Function

Public Function Window_Show(ByVal id As Integer, ByVal show As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).visible = show
        Window_Set_On_Top id
        Window_Show = True
    End If
End Function

Public Function Window_Set_Movable(ByVal id As Integer, ByVal Moveable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Moveable = Moveable
        Window_Set_Movable = True
    End If
End Function

Public Function Window_Set_Closeable(ByVal id As Integer, ByVal Closeable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Closeable = Closeable
        Window_Set_Closeable = True
    End If
End Function

Public Function Window_Set_Resizeable(ByVal id As Integer, ByVal Resizeable As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Windows(id).Resizeable = Resizeable
        Window_Set_Resizeable = True
    End If
End Function

Public Function Render_Visible() As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
On Error GoTo ErrHandler
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim text As String
    Dim window_index As Integer
    
    If WindowsCount = -1 Then Exit Function
    
    For i = WindowsCount To 0 Step -1
        window_index = Window_Index_Get(ZOrder(i))
        
        If Window_Get_Visible(window_index) Then
            With Windows(window_index)
                'Render the Window´s Grh
                If .GUMP Then
                    For j = 1 To 9
                        Engine.GUI_Grh_Render .grhs(j), .grh_x(j), .grh_y(j), , , , , .grh_width(j), .grh_height(j), , .grh_alphablending(j)
                    Next j
                ElseIf .animate Then
                    Engine.GUI_Grh_Animated_Render .grhs(1), .left, .top, .Frame, .width, .height
                    'Check if animation has ended
                    If .Frame = 1 Then
                        If .loops > 0 Then
                            .loops = .loops - 1
                            If .loops = 0 Then
                                'Set frame to the last grh in the animation
                                Dim TempInt As Integer
                                Dim frames As Integer
                                Engine.Grh_Info_Get .grhs(1), text, TempInt, TempInt, TempInt, TempInt, frames
                                Engine.Grh_Info_Get_Frame .grhs(1), frames, TempInt
                                .grhs(1) = TempInt
                                .animate = False
                            End If
                        End If
                    End If
                    'Don't show any controls it may have while animation is on
                    GoTo Continue
                Else
                    Engine.GUI_Grh_Render .grhs(1), .left, .top, , , , , .width, .height
                End If
                
                'Render Images
                If .imagecount > -1 Then
                    For j = 0 To .imagecount
                        Engine.GUI_Grh_Render .images(j).grh, .left + .images(j).left, .top + .images(j).top, , , , , .images(j).width, .images(j).height
                    Next j
                End If
                
                'Render Labels
                If .labelcount > -1 Then
                    For j = 0 To .labelcount
                        If .labels(j).Background_Color <> -1 Then
                            Engine.GUI_Box_Filled_Render .left + .labels(j).left, .top + .labels(j).top, .labels(j).width, .labels(j).height, .labels(j).Background_Color
                        End If
                        Engine.GUI_Text_Render .labels(j).text, .labels(j).font, .left + .labels(j).left, .top + .labels(j).top, .labels(j).width, .labels(j).height, .labels(j).font_color, .labels(j).Alignment, True
                    Next j
                End If
                
                'Render Buttons
                If .buttoncount > -1 Then
                    For j = 0 To .buttoncount
                        If .buttons(j).visible Then
                            If .buttons(j).Background_Color <> -1 Then
                                Engine.GUI_Box_Filled_Render .left + .buttons(j).left, .top + .buttons(j).top, .buttons(j).width, .buttons(j).height, .buttons(j).Background_Color
                            End If
                            If (ButtonClicked = j And ClickedControlwindow = window_index) Or Not .buttons(j).enabled Then
                                Engine.GUI_Grh_Render .buttons(j).Grh_Pressed, .left + .buttons(j).left, .top + .buttons(j).top, , , , , .buttons(j).width, .buttons(j).height
                            Else
                                Engine.GUI_Grh_Render .buttons(j).Grh_Unpressed, .left + .buttons(j).left, .top + .buttons(j).top, , , , , .buttons(j).width, .buttons(j).height
                            End If
                            If .buttons(j).text <> "" Then
                                Engine.GUI_Text_Render .buttons(j).text, 1, .left + .buttons(j).left, .top + .buttons(j).top, .buttons(j).width, .buttons(j).height, .buttons(j).font_color, fa_center
                            End If
                        End If
                    Next j
                End If
                
                'Render TextBoxes
                If .textscount > -1 Then
                    For j = 0 To .textscount
                        If .texts(j).visible Then
                            If .texts(j).Background_Color <> -1 Then
                                Engine.GUI_Box_Filled_Render .left + .texts(j).left, .top + .texts(j).top, .texts(j).width, .texts(j).height, .texts(j).Background_Color
                            End If
                            Engine.GUI_Grh_Render .texts(j).background_grh, .left + .texts(j).left, .top + .texts(j).top, , , , , .texts(j).width, .texts(j).height
                            If .texts(j).text <> "" Then
                                text = ""
                                If .texts(j).password Then
                                    For m = 1 To Len(.texts(j).text)
                                        text = text + "*"
                                    Next m
                                Else
                                    text = .texts(j).text
                                End If
                                If .texts(j).FontColor2 <> -1 Then
                                    Engine.GUI_Text_Render text, .texts(j).font_index, .left + .texts(j).left + 1, .top + .texts(j).top + 1, .texts(j).width, .texts(j).height, .texts(j).FontColor2, DT_TOP Or DT_LEFT, False
                                End If
                                Engine.GUI_Text_Render text, .texts(j).font_index, .left + .texts(j).left, .top + .texts(j).top, .texts(j).width, .texts(j).height, .texts(j).FontColor, DT_TOP Or DT_LEFT, False
                            End If
                        End If
                    Next j
                End If
                
                'Render CheckBoxes
                If .checkboxcount > -1 Then
                    For j = 0 To .checkboxcount
                        If .checkboxes(j).visible Then
                            If .checkboxes(j).Background_Color <> -1 Then
                                Engine.GUI_Box_Filled_Render .left + .checkboxes(j).left, .top + .checkboxes(j).top, .checkboxes(j).width, .checkboxes(j).height, .checkboxes(j).Background_Color, , , , False
                            End If
                            Engine.GUI_Text_Render .checkboxes(j).text, .checkboxes(j).font_index, .left + .checkboxes(j).left + CHECKBOX_WIDTH + 5, .top + .checkboxes(j).top, .checkboxes(j).width - CHECKBOX_WIDTH - 5, .checkboxes(j).height, .checkboxes(j).font_color, fa_left
                            If Not .checkboxes(j).value Then
                                If .checkboxes(j).grh_unchecked Then
                                    Engine.GUI_Grh_Render .checkboxes(j).grh_checked, .left + .checkboxes(j).left, .top + .checkboxes(j).top
                                Else
                                    Engine.GUI_Box_Filled_Render .left + .checkboxes(j).left, .top + .checkboxes(j).top, CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFFFFFFFF, , , , False
                                End If
                            Else
                                If .checkboxes(j).grh_checked Then
                                    Engine.GUI_Grh_Render .checkboxes(j).grh_checked, .left + .checkboxes(j).left, .top + .checkboxes(j).top
                                Else
                                    Engine.GUI_Box_Filled_Render .left + .checkboxes(j).left, .top + .checkboxes(j).top, CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFFFFFFFF, , , , False
                                    Engine.GUI_Text_Render "x", .checkboxes(j).font_index, .left + .checkboxes(j).left, .top + .checkboxes(j).top - 3, CHECKBOX_WIDTH, CHECKBOX_HEIGHT, &HFF000000, fa_center
                                End If
                            End If
                        End If
                    Next j
                End If
                
                'Render OptionButtons
                If .optionbuttoncount > -1 Then
                    For j = 0 To .optionbuttoncount
                        If .optionbuttons(j).visible Then
                            If .optionbuttons(j).Background_Color <> -1 Then
                                Engine.GUI_Box_Filled_Render .left + .optionbuttons(j).left, .top + .optionbuttons(j).top, OPTIONBUTTON_WIDTH + 5 + .optionbuttons(j).width, .optionbuttons(j).height, .optionbuttons(j).Background_Color
                            End If
                            Engine.GUI_Text_Render .optionbuttons(j).text, .optionbuttons(j).font_index, .left + .optionbuttons(j).left + OPTIONBUTTON_WIDTH + 5, .top + .optionbuttons(j).top, .optionbuttons(j).width - OPTIONBUTTON_WIDTH - 5, .optionbuttons(j).height, .optionbuttons(j).font_color, fa_left
                            If Not .optionbuttons(j).enabled Then
                                If .optionbuttons(j).grh_disabled Then
                                    Engine.GUI_Grh_Render .optionbuttons(j).grh_disabled, .left + .optionbuttons(j).left, .top + .optionbuttons(j).top
                                Else
                                    Engine.GUI_Box_Filled_Render .left + .optionbuttons(j).left, .top + .optionbuttons(j).top, OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFC0C0C0, , , , False, General_Convert_Degrees_To_Radians(45)
                                End If
                            ElseIf Not .optionbuttons(j).value Then
                                If .optionbuttons(j).grh_unselected Then
                                    Engine.GUI_Grh_Render .optionbuttons(j).grh_unselected, .left + .optionbuttons(j).left, .top + .optionbuttons(j).top
                                Else
                                    Engine.GUI_Box_Filled_Render .left + .optionbuttons(j).left, .top + .optionbuttons(j).top, OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFFFFFFF, , , , False, General_Convert_Degrees_To_Radians(45)
                                End If
                            Else
                                If .optionbuttons(j).grh_selected Then
                                    Engine.GUI_Grh_Render .optionbuttons(j).grh_selected, .left + .optionbuttons(j).left, .top + .optionbuttons(j).top
                                Else
                                    Engine.GUI_Box_Filled_Render .left + .optionbuttons(j).left, .top + .optionbuttons(j).top, OPTIONBUTTON_WIDTH, OPTIONBUTTON_HEIGHT, &HFFFFFFFF, , , , False, General_Convert_Degrees_To_Radians(45)
                                    Engine.GUI_Box_Filled_Render .left + .optionbuttons(j).left + 2, .top + .optionbuttons(j).top + 2, OPTIONBUTTON_WIDTH - 4, OPTIONBUTTON_HEIGHT - 4, &HFF000000, , , , False, General_Convert_Degrees_To_Radians(45)
                                End If
                            End If
                        End If
                    Next j
                End If
            End With
        End If
Continue:
    Next i
    
    Render_Visible = True
Exit Function

ErrHandler:
End Function

Public Function Window_Get_Button_Count(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Button_Count = Windows(id).buttoncount + 1
    End If
End Function

Public Function Window_Get_Text_Count(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Text_Count = Windows(id).textscount + 1
    End If
End Function

Public Function Window_Get_Label_Count(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Label_Count = Windows(id).labelcount + 1
    End If
End Function

Public Function Window_Change_Alphablending(ByVal id As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        If Windows(id).GUMP Then
            Dim index As Integer
            Dim loopc As Long
            Dim LoopD As Long
            
            index = X + (Y - 1) * 3
            
            'Make usre it's within bounds
            If index > 0 And index < 10 Then
                For LoopD = 0 To height - 1
                    For loopc = 0 To width - 1
                        If X + loopc < 4 Then
                            Windows(id).grh_alphablending(index + loopc) = Not Windows(id).grh_alphablending(index + loopc)
                        End If
                    Next loopc
                    index = index + 3
                    If index > 9 Then Exit For
                Next LoopD
                Window_Change_Alphablending = True
            End If
        End If
    End If
End Function

Public Function Window_Get_Image_Count(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Image_Count = Windows(id).imagecount + 1
    End If
End Function

Public Function Window_Get_Top(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Top = Windows(id).top
    End If
End Function

Public Function Window_Get_Grh(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Grh = Windows(id).grhs(1)
    End If
End Function

Public Function Window_Get_Movable(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Movable = Windows(id).Moveable
    End If
End Function

Public Function Window_Get_Visible(ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Visible = Windows(id).visible
    End If
End Function

Public Function Window_Get_Left(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Left = Windows(id).left
    End If
End Function

Public Function Window_Get_Width(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Width = Windows(id).width
    End If
End Function

Public Function Window_Get_Height(ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Window_Get_Height = Windows(id).height
    End If
End Function

Public Function Window_Set_Grh(ByVal id As Integer, ByVal grh As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        'Allow it only if it's not a GUMP generated window
        If Not Windows(id).GUMP Then
            Windows(id).grhs(1) = grh
            Window_Set_Grh = True
        End If
    End If
End Function

Public Function Window_Set_Top(ByVal id As Integer, ByVal top As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).top = top
        Window_Set_Top = True
    End If
End Function

Public Function Window_Set_Left(ByVal id As Integer, ByVal left As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).left = left
        Window_Set_Left = True
    End If
End Function

Public Function Window_Mouse_Set_Down(ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        WindowPressed = id
        
        Window_Set_On_Top id
        
        Window_Mouse_Set_Down = True
    End If
End Function

Public Function Window_Mouse_Get_Pressed() As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    Window_Mouse_Get_Pressed = WindowPressed
End Function

Public Sub Window_Mouse_Set_Up()
'*********************************************************
'Coded by Fredrik Alexandersson
'Modified by Juan Martín Sotuyo Dodero
'*********************************************************
    WindowPressed = -1
    ButtonClicked = -1
End Sub

Public Sub Window_Mouse_Reset()
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    WindowPressed = -1
End Sub

Public Function Window_Get_Active() As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If WindowsCount > -1 Then
        Window_Get_Active = Window_Active
    End If
End Function

Public Function Window_Set_Active(ByVal WindowID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Or WindowID = -1 Then
        Window_Active = WindowID
        Window_Set_Active = True
    End If
End Function

Public Function Window_Set_Animate(ByVal WindowID As Integer, ByVal animate As Boolean, ByVal loops As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).animate = animate
        Windows(WindowID).loops = loops
        Window_Set_Animate = True
    End If
End Function

Public Function Window_Set_Width(ByVal id As Integer, ByVal width As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).width = width
        Window_Set_Width = True
    End If
End Function

Public Function Window_Resize(ByVal id As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) Then
        If Windows(id).Resizeable Then
            With Windows(id)
                .top = Y
                .left = X
                .width = width
                .height = height
                If .GUMP Then
                    'Calculate the size of each grh
                    Dim h1 As Integer
                    Dim h2 As Integer
                    Dim h3 As Integer
                    Dim h4 As Integer
                    Dim h5 As Integer
                    Dim h6 As Integer
                    Dim h7 As Integer
                    Dim h8 As Integer
                    Dim h9 As Integer
                    Dim w1 As Integer
                    Dim w2 As Integer
                    Dim w3 As Integer
                    Dim w4 As Integer
                    Dim w5 As Integer
                    Dim w6 As Integer
                    Dim w7 As Integer
                    Dim w8 As Integer
                    Dim w9 As Integer
                    Dim temp_str As String
                    Dim temp As Integer
                    Dim x_factor As Single
                    Dim y_factor As Single
                    
                    'Get height and width of all grhs
                    Engine.Grh_Info_Get .grhs(1), temp_str, temp, temp, w1, h1, temp
                    Engine.Grh_Info_Get .grhs(2), temp_str, temp, temp, w2, h2, temp
                    Engine.Grh_Info_Get .grhs(3), temp_str, temp, temp, w3, h3, temp
                    Engine.Grh_Info_Get .grhs(4), temp_str, temp, temp, w4, h4, temp
                    Engine.Grh_Info_Get .grhs(5), temp_str, temp, temp, w5, h5, temp
                    Engine.Grh_Info_Get .grhs(6), temp_str, temp, temp, w6, h6, temp
                    Engine.Grh_Info_Get .grhs(7), temp_str, temp, temp, w7, h7, temp
                    Engine.Grh_Info_Get .grhs(8), temp_str, temp, temp, w8, h8, temp
                    Engine.Grh_Info_Get .grhs(9), temp_str, temp, temp, w9, h9, temp
                    
                    'Calculate stretching factor
                    x_factor = width / (w1 + w2 + w3)
                    y_factor = height / (h1 + h2 + h3)
                    
                    'Set x and width for all grhs
                    .grh_x(1) = .grh_x(4) = .grh_x(7) = X
                    
                    'First row
                    .grh_width(1) = .grh_x(2) = w1 * x_factor
                    .grh_width(2) = .grh_x(3) = w2 * x_factor
                    .grh_width(3) = w3 * x_factor
                    
                    'Second row
                    .grh_width(4) = .grh_x(5) = w4 * x_factor
                    .grh_width(5) = .grh_x(6) = w5 * x_factor
                    .grh_width(6) = w6 * x_factor
                    
                    'Third row
                    .grh_width(7) = .grh_x(8) = w7 * x_factor
                    .grh_width(8) = .grh_x(9) = w8 * x_factor
                    .grh_width(9) = w9 * x_factor
                    
                    'Set y and height for all grhs
                    .grh_y(1) = .grh_y(2) = .grh_y(3) = Y
                    'First column
                    .grh_height(1) = .grh_y(4) = h1 * y_factor
                    .grh_height(4) = .grh_y(7) = h4 * y_factor
                    .grh_height(7) = h7 * y_factor
                    
                    'Second column
                    .grh_height(2) = .grh_y(5) = h2 * y_factor
                    .grh_height(5) = .grh_y(8) = h5 * y_factor
                    .grh_height(8) = h8 * y_factor
                    
                    'Third column
                    .grh_height(3) = .grh_y(3) = h3 * y_factor
                    .grh_height(6) = .grh_y(6) = h6 * y_factor
                    .grh_height(9) = h9 * y_factor
                End If
            End With
            Window_Resize = True
        End If
    End If
End Function

Public Function Window_Set_Height(ByVal id As Integer, ByVal height As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Window_Check(id) Then
        Windows(id).height = height
        Window_Set_Height = True
    End If
End Function

Public Function Window_Move(ByVal id As Integer, ByVal left As Integer, ByVal top As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(id) And Window_Get_Movable(id) Then
        Windows(id).left = left
        Windows(id).top = top
        Window_Move = True
    End If
End Function

Public Function Button_Hit_Test(ByVal WindowID As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).buttoncount
            If Windows(WindowID).buttons(i).visible And Windows(WindowID).buttons(i).enabled Then
                If Windows(WindowID).top + Windows(WindowID).buttons(i).top < Y And Windows(WindowID).buttons(i).height + Windows(WindowID).buttons(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).buttons(i).left < X And Windows(WindowID).left + Windows(WindowID).buttons(i).left + Windows(WindowID).buttons(i).width > X Then
                        Button_Hit_Test = i
                        ButtonClicked = i
                        ClickedControlwindow = WindowID
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    Button_Hit_Test = -1
End Function

Public Function Button_Create(ByVal Grh_Unpressed As Integer, ByVal WindowID As Integer, ByVal text As String, ByVal FontColor As Long, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, Optional ByVal Grh_Pressed As Integer = 0, Optional ByVal Background_Color As Long = -1, Optional enabled As Boolean = True, Optional button_id As Integer = 0, Optional ByVal GUMP_type As Integer = 2) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).buttoncount = Windows(WindowID).buttoncount + 1
        ReDim Preserve Windows(WindowID).buttons(0 To Windows(WindowID).buttoncount)
        With Windows(WindowID).buttons(Windows(WindowID).buttoncount)
            .Grh_Unpressed = Grh_Unpressed
            .text = text
            .top = top
            .left = left
            .width = width
            .height = height
            .font_color = FontColor
            .Background_Color = Background_Color
            .visible = True
            .enabled = enabled
            .button_id = button_id
            .GUMP_type = GUMP_type
            If Grh_Pressed = 0 Then
                .Grh_Pressed = .Grh_Unpressed
            Else
                .Grh_Pressed = Grh_Pressed
            End If
        End With
        Button_Create = Windows(WindowID).buttoncount
    Else
        Button_Create = -1
    End If
End Function

Public Function Button_Get_Font_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Font_Color = Windows(WindowID).buttons(id).font_color
    End If
End Function

Public Function Button_Get_Background_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Background_Color = Windows(WindowID).buttons(id).Background_Color
    End If
End Function

Public Function Button_Get_GUMP_Type(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_GUMP_Type = Windows(WindowID).buttons(id).GUMP_type
    End If
End Function

Public Function Button_Get_Button_ID(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Button_ID = Windows(WindowID).buttons(id).button_id
    End If
End Function

Public Function Button_Get_Text(ByVal WindowID As Integer, ByVal id As Integer) As String
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Text = Windows(WindowID).buttons(id).text
    End If
End Function

Public Function Button_Get_Top(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Top = Windows(WindowID).buttons(id).top
    End If
End Function

Public Function Button_Get_Left(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Left = Windows(WindowID).buttons(id).left
    End If
End Function

Public Function Button_Get_Width(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Windows(WindowID).buttoncount > id Or Windows(WindowID).buttoncount = id Then
        Button_Get_Width = Windows(WindowID).buttons(id).width
    End If
End Function

Public Function Button_Get_Height(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Height = Windows(WindowID).buttons(id).height
    End If
End Function

Public Function Text_Get_Grh(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Grh = Windows(WindowID).texts(id).background_grh
    End If
End Function

Public Function Text_Get_Font_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font_Color = Windows(WindowID).texts(id).FontColor
    End If
End Function

Public Function Text_Get_Font_Color2(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font_Color2 = Windows(WindowID).texts(id).FontColor2
    End If
End Function

Public Function Text_Get_Background_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Background_Color = Windows(WindowID).texts(id).Background_Color
    End If
End Function

Public Function Text_Get_Password(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Password = Windows(WindowID).texts(id).password
    End If
End Function

Public Function Text_Get_Font(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Font = Windows(WindowID).texts(id).font_index
    End If
End Function

Public Function Text_Get_Text(ByVal WindowID As Integer, ByVal id As Integer) As String
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Text = Windows(WindowID).texts(id).text
    End If
End Function

Public Function Text_Get_Top(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Top = Windows(WindowID).texts(id).top
    End If
End Function

Public Function Text_Get_Left(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Left = Windows(WindowID).texts(id).left
    End If
End Function

Public Function Text_Get_Width(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Width = Windows(WindowID).texts(id).width
    End If
End Function

Public Function Text_Get_Height(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    If Text_Check(WindowID, id) Then
        Text_Get_Height = Windows(WindowID).texts(id).height
    End If
End Function

Public Function Button_Get_Grh_Unpressed(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Grh_Unpressed = Windows(WindowID).buttons(id).Grh_Unpressed
    End If
End Function

Public Function Button_Get_Grh_Pressed(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Grh_Pressed = Windows(WindowID).buttons(id).Grh_Pressed
    End If
End Function

Public Function Text_Add_Text(ByVal WindowID As Integer, ByVal id As Integer, ByVal text As String) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan MArtín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Windows(WindowID).texts(id).text = Windows(WindowID).texts(id).text & text
        Text_Add_Text = True
        Exit Function
    End If
End Function

Public Function Text_Remove_Char(ByVal WindowID As Integer, ByVal id As Integer, ByVal NumChars As Integer) As Boolean
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        If Len(Windows(WindowID).texts(id).text) >= NumChars Then
            Windows(WindowID).texts(id).text = left$(Windows(WindowID).texts(id).text, Len(Windows(WindowID).texts(id).text) - NumChars)
        Else
            Windows(WindowID).texts(id).text = ""
        End If
        Text_Remove_Char = True
        Exit Function
    End If
End Function

Public Function Text_Hit_Test(ByVal WindowID As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
'*********************************************************
'Coded by Fredrik Alexandersson
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).textscount
            If Windows(WindowID).texts(i).visible Then
                If Windows(WindowID).top + Windows(WindowID).texts(i).top < Y And Windows(WindowID).texts(i).height + Windows(WindowID).texts(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).texts(i).left < X And Windows(WindowID).left + Windows(WindowID).texts(i).left + Windows(WindowID).texts(i).width > X Then
                        Text_Hit_Test = i
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    TextActive = -1
    Text_Hit_Test = -1
End Function

Private Function Window_Check(ByVal WindowID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Window_Check = (WindowID <= Last_Handle And WindowID > -1 And Window_Index_Get(WindowID) > -1)
End Function

Private Function Window_Index_Get(ByVal WindowID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    For loopc = 0 To WindowsCount
        If Windows(loopc).handle = WindowID Then
            Window_Index_Get = loopc
            Exit Function
        End If
    Next loopc
    
    Window_Index_Get = -1
End Function

Private Function Button_Check(ByVal WindowID As Integer, ByVal ButtonID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Button_Check = (ButtonID <= Windows(WindowID).buttoncount And ButtonID > -1)
    End If
End Function

Private Function Text_Check(ByVal WindowID As Integer, ByVal TextID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Text_Check = (TextID <= Windows(WindowID).textscount And TextID > -1)
    End If
End Function

Public Function Label_Create(ByVal WindowID As Integer, ByVal FontIndex As Integer, ByVal FontColor As Long, ByVal text As String, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, ByVal Alignment As FontAlignment, Optional ByVal BackgroundColor As Long = -1) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).labelcount = Windows(WindowID).labelcount + 1
        ReDim Preserve Windows(WindowID).labels(0 To Windows(WindowID).labelcount)
        With Windows(WindowID).labels(Windows(WindowID).labelcount)
            .text = text
            .font = FontIndex
            .font_color = FontColor
            .Alignment = Alignment
            .Background_Color = BackgroundColor
            .top = top
            .left = left
            .height = height
            .width = width
        End With
        Label_Create = Windows(WindowID).labelcount
    Else
        Label_Create = -1
    End If
End Function

Public Function Label_Set_Text(ByVal WindowID As Integer, ByVal id As Integer, ByVal text As String) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Windows(WindowID).labels(id).text = text
        Label_Set_Text = True
    End If
End Function

Private Function Label_Check(ByVal WindowID As Integer, ByVal LabelID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Label_Check = (LabelID <= Windows(WindowID).labelcount And LabelID > -1)
    End If
End Function

Public Function Label_Get_Top(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Top = Windows(WindowID).labels(id).top
    End If
End Function

Public Function Label_Get_Left(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Left = Windows(WindowID).labels(id).left
    End If
End Function

Public Function Label_Get_Height(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Height = Windows(WindowID).labels(id).height
    End If
End Function

Public Function Label_Get_Width(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Width = Windows(WindowID).labels(id).width
    End If
End Function

Public Function Label_Get_Text(ByVal WindowID As Integer, ByVal id As Integer) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Text = Windows(WindowID).labels(id).text
    End If
End Function

Public Function Label_Get_Font(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Font = Windows(WindowID).labels(id).font
    End If
End Function

Public Function Label_Get_Font_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Font_Color = Windows(WindowID).labels(id).font_color
    End If
End Function

Public Function Label_Get_Background_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Background_Color = Windows(WindowID).labels(id).Background_Color
    End If
End Function

Public Function Label_Get_Alignment(ByVal WindowID As Integer, ByVal id As Integer) As FontAlignment
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Label_Check(WindowID, id) Then
        Label_Get_Alignment = Windows(WindowID).labels(id).Alignment
    End If
End Function

Public Function CheckBox_Create(ByVal WindowID As Integer, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, ByVal grh_unchecked As Integer, ByVal grh_checked As Integer, ByVal text As String, Optional ByVal BackgroundColor As Long = -1, Optional ByVal value As Boolean = False, Optional ByVal font As Integer = 1, Optional ByVal font_color As Long, Optional ByVal switch_id As Integer = 0) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).checkboxcount = Windows(WindowID).checkboxcount + 1
        ReDim Preserve Windows(WindowID).checkboxes(0 To Windows(WindowID).checkboxcount)
        With Windows(WindowID).checkboxes(Windows(WindowID).checkboxcount)
            .top = top
            .left = left
            .width = width
            .height = height
            .value = value
            .font_color = font_color
            .font_index = font
            .grh_unchecked = grh_unchecked
            .grh_checked = grh_checked
            .text = text
            .Background_Color = BackgroundColor
            .visible = True
            .enabled = True
            .switch_id = switch_id
        End With
        CheckBox_Create = Windows(WindowID).checkboxcount
    Else
        CheckBox_Create = -1
    End If
End Function

Private Function CheckBox_Check(ByVal WindowID As Integer, ByVal CheckBoxID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        CheckBox_Check = (CheckBoxID <= Windows(WindowID).checkboxcount And CheckBoxID > -1)
    End If
End Function

Public Function CheckBox_Get_Top(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Top = Windows(WindowID).checkboxes(id).top
    End If
End Function

Public Function CheckBox_Get_Switch_ID(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Switch_ID = Windows(WindowID).checkboxes(id).switch_id
    End If
End Function

Public Function CheckBox_Get_Left(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Left = Windows(WindowID).checkboxes(id).left
    End If
End Function

Public Function CheckBox_Get_Width(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Width = Windows(WindowID).checkboxes(id).width
    End If
End Function

Public Function CheckBox_Get_Height(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Height = Windows(WindowID).checkboxes(id).height
    End If
End Function

Public Function CheckBox_Get_Value(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Value = Windows(WindowID).checkboxes(id).value
    End If
End Function

Public Function CheckBox_Get_Enabled(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Enabled = Windows(WindowID).checkboxes(id).enabled
    End If
End Function

Public Function CheckBox_Hit_Test(ByVal WindowID As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).checkboxcount
            If Windows(WindowID).checkboxes(i).visible = True And Windows(WindowID).checkboxes(i).enabled = True Then
                If Windows(WindowID).top + Windows(WindowID).checkboxes(i).top < Y And CHECKBOX_HEIGHT + Windows(WindowID).checkboxes(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).checkboxes(i).left < X And Windows(WindowID).left + Windows(WindowID).checkboxes(i).left + CHECKBOX_WIDTH > X Then
                        CheckBox_Hit_Test = i
                        'Change value
                        Windows(WindowID).checkboxes(i).value = Not Windows(WindowID).checkboxes(i).value
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    CheckBox_Hit_Test = -1
End Function

Public Function CheckBox_Get_Font(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Font = Windows(WindowID).checkboxes(id).font_index
    End If
End Function

Public Function CheckBox_Get_Font_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Font_Color = Windows(WindowID).checkboxes(id).font_color
    End If
End Function

Public Function CheckBox_Get_Text(ByVal WindowID As Integer, ByVal id As Integer) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Text = Windows(WindowID).checkboxes(id).text
    End If
End Function

Public Function CheckBox_Get_Grh_Unchecked(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Grh_Unchecked = Windows(WindowID).checkboxes(id).grh_unchecked
    End If
End Function

Public Function CheckBox_Get_Grh_Checked(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Grh_Checked = Windows(WindowID).checkboxes(id).grh_checked
    End If
End Function

Public Function CheckBox_Get_Background_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Background_Color = Windows(WindowID).checkboxes(id).Background_Color
        Exit Function
    End If
    CheckBox_Get_Background_Color = -1
End Function

Public Function CheckBox_Get_Visible(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        CheckBox_Get_Visible = Windows(WindowID).checkboxes(id).visible
    End If
End Function

Public Function CheckBox_Set_Value(ByVal WindowID As Integer, ByVal id As Integer, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).value = value
        CheckBox_Set_Value = True
    End If
End Function

Public Function CheckBox_Set_Visible(ByVal WindowID As Integer, ByVal id As Integer, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).visible = value
        CheckBox_Set_Visible = True
    End If
End Function

Public Function CheckBox_Set_Enabled(ByVal WindowID As Integer, ByVal id As Integer, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If CheckBox_Check(WindowID, id) Then
        Windows(WindowID).checkboxes(id).enabled = value
        CheckBox_Set_Enabled = True
    End If
End Function

Public Function Button_Set_Unpressed_Grh(ByVal WindowID As Integer, ByVal id As Integer, ByVal grh_index As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).Grh_Unpressed = grh_index
        Button_Set_Unpressed_Grh = True
    End If
End Function

Public Function Button_Set_Pressed_Grh(ByVal WindowID As Integer, id As Integer, grh_index As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).Grh_Pressed = grh_index
        Button_Set_Pressed_Grh = True
    End If
End Function

Public Function Button_Set_Top(ByVal WindowID As Integer, ByVal id As Integer, ByVal top As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).top = top
        Button_Set_Top = True
    End If
End Function

Public Function Button_Set_Text(ByVal WindowID As Integer, ByVal id As Integer, ByVal text As String) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).text = text
        Button_Set_Text = True
    End If
End Function

Public Function Button_Set_Enabled(ByVal WindowID As Integer, ByVal id As Integer, ByVal enabled As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).enabled = enabled
        Button_Set_Enabled = True
    End If
End Function

Public Function Button_Get_Enabled(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Button_Get_Enabled = Windows(WindowID).buttons(id).enabled
    End If
End Function

Public Function Button_Set_Visible(ByVal WindowID As Integer, ByVal id As Integer, ByVal visible As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Button_Check(WindowID, id) Then
        Windows(WindowID).buttons(id).visible = visible
        Button_Set_Visible = True
    End If
End Function

Public Function Window_Get_CheckBox_Count(ByVal WindowID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Window_Get_CheckBox_Count = Windows(WindowID).buttoncount + 1
    End If
End Function

Private Sub Window_Set_On_Top(ByVal WindowID As Integer)
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    'Move it to the top of the ZOrder list
    loopc = UBound(ZOrder)
    
    'Find it´s previous index
    Do Until ZOrder(loopc) = WindowID
        loopc = loopc - 1
    Loop
    
    'Move back all indexes untill we hit top of the list
    Do Until loopc = 0
        ZOrder(loopc) = ZOrder(loopc - 1)
        loopc = loopc - 1
    Loop
    
    'Place new head of the list
    ZOrder(0) = WindowID
End Sub

Public Function OptionButton_Create(ByVal WindowID As Integer, ByVal top As Integer, ByVal left As Integer, ByVal width As Integer, ByVal height As Integer, ByVal grh_unselected As Integer, ByVal grh_selected As Integer, ByVal grh_disabled As Long, ByVal text As String, Optional ByVal group_index As Integer = 0, Optional ByVal BackgroundColor As Long = -1, Optional ByVal font As Integer = 1, Optional ByVal font_color As Long, Optional ByVal switch_id As Integer = 0) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Windows(WindowID).optionbuttoncount = Windows(WindowID).optionbuttoncount + 1
        ReDim Preserve Windows(WindowID).optionbuttons(0 To Windows(WindowID).optionbuttoncount)
        With Windows(WindowID).optionbuttons(Windows(WindowID).optionbuttoncount)
            .top = top
            .left = left
            .width = width
            .height = height
            .value = False
            .font_color = font_color
            .font_index = font
            .grh_unselected = grh_unselected
            .grh_selected = grh_selected
            .grh_disabled = grh_disabled
            .text = text
            .Background_Color = BackgroundColor
            .visible = True
            .enabled = True
            .group = group_index
            .switch_id = switch_id
        End With
        
        OptionButton_Create = Windows(WindowID).checkboxcount
        
        'If it´s the first OptionButton of a group, set it´s value to true
        Dim loopc As Long
        For loopc = 0 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(loopc).group = group_index And loopc <> Windows(WindowID).optionbuttoncount Then
                Exit Function
            End If
        Next loopc
        
        Windows(WindowID).optionbuttons(Windows(WindowID).optionbuttoncount).value = True
    Else
        OptionButton_Create = -1
    End If
End Function

Private Function OptionButton_Check(ByVal WindowID As Integer, ByVal OptionButtonID As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        OptionButton_Check = (OptionButtonID <= Windows(WindowID).optionbuttoncount And OptionButtonID > -1)
    End If
End Function

Public Function OptionButton_Get_Top(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Top = Windows(WindowID).optionbuttons(id).top
    End If
End Function

Public Function OptionButton_Get_Switch_ID(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Switch_ID = Windows(WindowID).optionbuttons(id).switch_id
    End If
End Function

Public Function OptionButton_Get_Left(WindowID As Integer, id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Left = Windows(WindowID).optionbuttons(id).left
    End If
End Function

Public Function OptionButton_Get_Width(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Width = Windows(WindowID).optionbuttons(id).width
    End If
End Function

Public Function OptionButton_Get_Height(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Height = Windows(WindowID).optionbuttons(id).height
    End If
End Function

Public Function OptionButton_Get_Value(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Value = Windows(WindowID).optionbuttons(id).value
    End If
End Function

Public Function OptionButton_Get_Visible(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Visible = Windows(WindowID).optionbuttons(id).visible
    End If
End Function

Public Function OptionButton_Hit_Test(ByVal WindowID As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    Dim j As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(i).visible = True And Windows(WindowID).optionbuttons(i).enabled = True Then
                If Windows(WindowID).top + Windows(WindowID).optionbuttons(i).top < Y And OPTIONBUTTON_HEIGHT + Windows(WindowID).optionbuttons(i).top + Windows(WindowID).top > Y Then
                    If Windows(WindowID).left + Windows(WindowID).optionbuttons(i).left < X And Windows(WindowID).left + Windows(WindowID).optionbuttons(i).left + OPTIONBUTTON_WIDTH > X Then
                        OptionButton_Hit_Test = i
                        'Change value
                        Windows(WindowID).optionbuttons(i).value = True
                        'Look for each option button in the group and set it´s value to false
                        For j = 0 To Windows(WindowID).optionbuttoncount
                            If Windows(WindowID).optionbuttons(j).group = Windows(WindowID).optionbuttons(i).group And i <> j Then
                                Windows(WindowID).optionbuttons(j).value = False
                            End If
                        Next j
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    
    OptionButton_Hit_Test = -1
End Function

Public Function OptionButton_Get_Font(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Font = Windows(WindowID).optionbuttons(id).font_index
    End If
End Function

Public Function OptionButton_Get_Font_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Font_Color = Windows(WindowID).optionbuttons(id).font_color
    End If
End Function

Public Function OptionButton_Get_Text(ByVal WindowID As Integer, ByVal id As Integer) As String
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Text = Windows(WindowID).optionbuttons(id).text
    End If
End Function

Public Function OptionButton_Get_Grh_Unselected(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Unselected = Windows(WindowID).optionbuttons(id).grh_unselected
    End If
End Function

Public Function OptionButton_Get_Grh_Selected(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Selected = Windows(WindowID).optionbuttons(id).grh_selected
    End If
End Function

Public Function OptionButton_Get_Background_Color(ByVal WindowID As Integer, ByVal id As Integer) As Long
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Background_Color = Windows(WindowID).optionbuttons(id).Background_Color
    End If
End Function

Public Function OptionButton_Get_Selected_From_Group(ByVal WindowID As Integer, ByVal group_index As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'Returns the index of the selected OptionButton within a group
'*********************************************************
    Dim loopc As Long
    
    If Window_Check(WindowID) Then
        If Windows(WindowID).optionbuttoncount > -1 Then
            For loopc = 0 To Windows(WindowID).optionbuttoncount
                If Windows(WindowID).optionbuttons(loopc).group = group_index And _
                  Windows(WindowID).optionbuttons(loopc).value = True Then
                    OptionButton_Get_Selected_From_Group = loopc
                    Exit Function
                End If
            Next loopc
        End If
    End If
    
    OptionButton_Get_Selected_From_Group = -1
End Function

Public Function OptionButton_Group_Disable(ByVal WindowID As Integer, ByVal group_index As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    If Windows(WindowID).optionbuttoncount = -1 Then Exit Function
    
    For loopc = 0 To Windows(WindowID).optionbuttoncount
        If Windows(WindowID).optionbuttons(loopc).group = group_index Then
            Windows(WindowID).optionbuttons(loopc).enabled = False
        End If
    Next loopc
    
    OptionButton_Group_Disable = True
End Function

Public Function OptionButton_Group_Enable(ByVal WindowID As Integer, ByVal group_index As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    If Windows(WindowID).optionbuttoncount = -1 Then Exit Function
    
    For loopc = 0 To Windows(WindowID).optionbuttoncount
        If Windows(WindowID).optionbuttons(loopc).group = group_index Then
            Windows(WindowID).optionbuttons(loopc).enabled = True
        End If
    Next loopc
    
    OptionButton_Group_Enable = True
End Function

Public Function OptionButton_Set_Visible(ByVal WindowID As Integer, ByVal id As Integer, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).visible = value
        OptionButton_Set_Visible = True
    End If
End Function

Public Function OptionButton_Set_Enabled(ByVal WindowID As Integer, ByVal id As Integer, ByVal value As Boolean) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).enabled = value
        OptionButton_Set_Enabled = True
    End If
End Function

Public Function OptionButton_Set_Selected(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim loopc As Long
    
    If OptionButton_Check(WindowID, id) Then
        Windows(WindowID).optionbuttons(id).value = True
        For loopc = 1 To Windows(WindowID).optionbuttoncount
            If Windows(WindowID).optionbuttons(loopc).group = Windows(WindowID).optionbuttons(id).group And loopc <> id Then
                Windows(WindowID).optionbuttons(loopc).value = False
            End If
        Next loopc
        OptionButton_Set_Selected = True
    End If
End Function

Public Function Window_Get_OptionButton_Count(ByVal WindowID As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(WindowID) Then
        Window_Get_OptionButton_Count = Windows(WindowID).optionbuttoncount + 1
    End If
End Function

Public Function OptionButton_Get_Grh_Disabled(ByVal WindowID As Integer, ByVal id As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Grh_Disabled = Windows(WindowID).optionbuttons(id).grh_disabled
    End If
End Function

Public Function OptionButton_Get_Enabled(ByVal WindowID As Integer, ByVal id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If OptionButton_Check(WindowID, id) Then
        OptionButton_Get_Enabled = Windows(WindowID).optionbuttons(id).enabled
    End If
End Function

Public Function Text_Move_Focus() As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(Text_Active_Get_Window, Text_Get_Active) Then
        If Windows(Text_Active_Get_Window).textscount > Text_Get_Active Then
            TextActive = TextActive + 1
        Else
            TextActive = 0
        End If
        
        Text_Move_Focus = True
    End If
End Function

Public Function Text_Set_Text(ByVal WindowID As Integer, ByVal id As Integer, ByVal text As String) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Text_Check(WindowID, id) Then
        Windows(WindowID).texts(id).text = text
        Text_Set_Text = True
    End If
End Function

Public Function Label_Hit_Test(ByVal WindowID As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Dim i As Long
    
    If Window_Check(WindowID) Then
        Window_Set_On_Top WindowID
        For i = 0 To Windows(WindowID).labelcount
            If Windows(WindowID).top + Windows(WindowID).labels(i).top < Y And Windows(WindowID).labels(i).height + Windows(WindowID).labels(i).top + Windows(WindowID).top > Y Then
                If Windows(WindowID).left + Windows(WindowID).labels(i).left < X And Windows(WindowID).left + Windows(WindowID).labels(i).left + Windows(WindowID).labels(i).width > X Then
                    Label_Hit_Test = i
                    Exit Function
                End If
            End If
        Next i
    End If
    
    Label_Hit_Test = -1
End Function

Public Function Window_Delete(ByVal window_id As Integer) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    If Window_Check(window_id) Then
        'Delete all controls
        Erase Windows(window_id).buttons
        Erase Windows(window_id).checkboxes
        Erase Windows(window_id).images
        Erase Windows(window_id).labels
        Erase Windows(window_id).optionbuttons
        Erase Windows(window_id).texts
        
        'Remove window index from ZOrder array
        Dim loopc As Long
        For loopc = 1 To WindowsCount
            If ZOrder(loopc) = window_id Then Exit For
        Next loopc
        For loopc = loopc To WindowsCount - 1
            ZOrder(loopc) = ZOrder(loopc + 1)
        Next loopc
        
        'Remove window index from the windows array
        For loopc = Window_Index_Get(window_id) To WindowsCount - 1
            Windows(loopc) = Windows(loopc + 1)
        Next loopc
        
        'Decrease Window's Count
        WindowsCount = WindowsCount - 1
        'Decrease last handle
        Last_Handle = Last_Handle - 1
        
        ReDim Preserve ZOrder(WindowsCount)
        ReDim Preserve Windows(WindowsCount)
        
        Window_Delete = True
    End If
End Function

Public Function GUI_Initialize(ByRef tile_engine As clsTileEngineX) As Boolean
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Set Engine = tile_engine
    
    GUI_Initialize = True
End Function

Private Sub Class_Initialize()
'*********************************************************
'Coded by Fredrik Alexandersson
'Edited by Juan Martín Sotuyo Dodero
'*********************************************************
    WindowsCount = -1
    ButtonClicked = -1
    ClickedControlwindow = -1
    
    Window_Active = -1
    
    Last_Handle = -1
    
    TextWindow = -1
    TextActive = -1
End Sub

Private Sub Class_Terminate()
'*********************************************************
'Coded by Juan Martín Sotuyo Dodero
'*********************************************************
    Set Engine = Nothing
End Sub
