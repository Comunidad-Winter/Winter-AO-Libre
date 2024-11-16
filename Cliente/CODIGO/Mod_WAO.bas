Attribute VB_Name = "Mod_WAO"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Public Const WM_SETTEXT = &HC
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const EM_SETREADONLY = &HCF
    
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer
'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
(ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As Long, lpdwProcessId As Long) As Long


Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Const PROCESS_TERMINATE = &H1
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103
Public MinEleccion As Integer, MaxEleccion As Integer
Public Actual As Integer
 'Old fashion BitBlt function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public PistaActual As Integer
Public VisorTxt As Integer
'<------------>TRANSPARENCIAS<------------>
'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
  
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long
  

  
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 '<------------>FIN TRANSPARENCIAS<------------>
'<------------MSN------------------>
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Public Const WM_COPYDATA = &H4A
'<--------------FIN MSN<----------------->



'Drag & Drop de Formularios:
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessagges Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const RGN_OR = 2

'Seguridad AntiEditados:

Public Const SegCode As String * 18 = "LOR817W0I@KPFKY*"

'Configurar Teclas:
Public CustomKeys As New clsCustomKeys

'Opciones
Public MiniMapX As Boolean
Public MapNameX As Boolean
Public EfectosDiaX As Boolean
Public EfectosAlphaX As Boolean
Public ConsolaX As Boolean
Public FpsX As Boolean

Public MiniMapY As Boolean
Public MapNameY As Boolean
Public EfectosDiaY As Boolean
Public EfectosAlphaY As Boolean
Public ConsolaY As Boolean
Public FpsY As Boolean

'Efectos de Alpha Blending:
Public AlphaX As Byte
Public Desbanecimiento1 As Boolean
Public Desbanecimiento2 As Boolean

'Seguridad AntiDobleCliente
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
 
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
 
Private Const ERROR_ALREADY_EXISTS = 183&
 
Private mutexHID As Long


'Efectos Climaticos:
Public Anochecer As Byte
Public Atardecer As Byte
Public Amanecer As Byte
Public Niebla As Byte

'******************************************************************************
'Drag & Drop de Formularios:
'******************************************************************************
Public Sub HookSurfaceHwnd(frm As Form)
    Call ReleaseCapture
    Call SendMessagges(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
'******************************************************************************
'/Drag & Drop de Formularios:
'******************************************************************************

'******************************************************************************
'Actualizador:
'******************************************************************************
Public Function CheckUpdates()
On Error Resume Next
    Dim iX As Integer, tX As Integer, DifX As Integer, strsX As String
        'Take iX an tX valors:
            iX = frmConnect.Updates.OpenURL("http://wao.webcindario.com/VEREXE.txt")
                tX = GetVar(App.Path & "\INIT\Update.ini", "INIT", "X")
                
        'Check Valor
        If Val(iX) <= 0 Then
            Exit Function
        End If
            DifX = iX - tX
    'Chek for Updates!
    If Not (DifX = 0) Then
        MsgBox "Hay Actualizaciones Disponibles, Se abrira el Autoupdate para Actualizar", vbInformation
            Call Shell(App.Path & "\AutoUpdate.EXE")
            Call ReleaseInstance
        End
    End If
End Function

'******************************************************************************
'/Actualizador
'******************************************************************************

'******************************************************************************
'Macros
'******************************************************************************
Public Function DoAccionTecla(ByVal Tecla As String)

Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", Tecla, "Accion")
    
    If Accion = 1 Then
        Dim Comando As String
        Comando = GetVar(IniPath & "Macros.bin", Tecla, "Comando")
            Call SendData("/" & Comando)
            
    ElseIf Accion = 2 Then
        Dim Usar As Byte
        Usar = GetVar(IniPath & "Macros.bin", Tecla, "UsarItem")
            Call SendData("USA" & Usar)
            
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
        Equipar = GetVar(IniPath & "Macros.bin", Tecla, "EquiparItem")
            Call SendData("EQUI" & Equipar)
            
    ElseIf Accion = 4 Then
        Dim Hechizo As Byte
        Hechizo = GetVar(IniPath & "Macros.bin", Tecla, "LanzarHechizo")
            Call SendData("LH" & Hechizo)
            Call SendData("UK" & Magia)
            
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Call frmMacros.Show(vbModeless, frmMain)
        
    ElseIf Accion = "" Then
        Exit Function
    End If
    
End Function

Public Function DibujarMacros(ByVal Tecla As Integer)

Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", "F" & Tecla, "Accion")

    If Accion = 1 Then
        frmMain.Macros(Tecla).Picture = LoadPicture(App.Path & "\Interfaces\Comandos.bmp")
        
    ElseIf Accion = 2 Then
        Dim Usar As Byte
            Usar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "UsarItem")
        Dim Grh As Integer
            Grh = Inventario.GrhIndex(Usar)
                Call DibujarMacrosUsarEquipar(Grh, Tecla)
         
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
            Equipar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "EquiparItem")
        Dim Grhs As Integer
            Grhs = Inventario.GrhIndex(Equipar)
                Call DibujarMacrosUsarEquipar(Grhs, Tecla)
                
    ElseIf Accion = 4 Then
        frmMain.Macros(Tecla).Picture = LoadPicture(App.Path & "\Interfaces\Hechizos.bmp")
        
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Exit Function
    End If
End Function
Public Function DibujarMacrosUsarEquipar(ByVal Grh As Integer, ByVal Tecla As Integer)
Dim SR As RECT, DR As RECT
SR.Left = 0
SR.Top = 0
SR.Right = 34
SR.Bottom = 34
DR.Left = 0
DR.Top = 0
DR.Right = 34
DR.Bottom = 34
Call DrawGrhtoHdc(frmMain.Macros(Tecla).hwnd, frmMain.Macros(Tecla).hdc, Grh, SR, DR)
End Function
Public Function CargarMacros()
    Dim i As Byte
        For i = 1 To 11
            Call DibujarMacros(i)
        Next i
End Function
'AntiDoble Cliente
Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
   
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
   
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function
 
Public Function FindPreviousInstance() As Boolean
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        FindPreviousInstance = False
    Else
        FindPreviousInstance = True
    End If
End Function
 
Public Sub ReleaseInstance()
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub



'Opciones
Public Function CargarOpciones()


'MiniMapa
MiniMapX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "Minimapa")
If Val(MiniMapX) = 1 Then
    MiniMapY = True
    frmMain.MiniMap.Visible = True
    frmOpciones.Minimapa.Caption = "Desactivar MiniMapa"
Else
    MiniMapY = False
    frmMain.MiniMap.Visible = False
    frmOpciones.Minimapa.Caption = "Activar MiniMapa"
End If

'Nombre del Mapa
MapNameX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "NombreMapa")
If Val(MapNameX) = 1 Then
    MapNameY = True
    frmOpciones.Check1.value = 1
    frmMain.MapName.Visible = True
Else
    MapNameY = False
    frmOpciones.Check1.value = 0
    frmMain.MapName.Visible = False
End If

'Efectos de Dia
EfectosDiaX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "DiaNoche")
If Val(EfectosDiaX) = 1 Then
    EfectosDiaY = True
    frmOpciones.ActivarNoche.value = 1
Else
    EfectosDiaY = False
    frmOpciones.ActivarNoche.value = 0
End If

'FPS
FpsX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "FPS")
If Val(FpsX) = 1 Then
    FpsY = True
    frmOpciones.Check2.value = 1
Else
    FpsY = False
    frmOpciones.Check2.value = 0
End If

'Nombres
Nombres = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "Nombres")
If Val(Nombres) = 1 Then
    frmOpciones.chkop.value = 1
    Nombres = True
Else
    frmOpciones.chkop.value = 0
    Nombres = False
End If

'Musica
Musica = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "Musica")
If Val(Musica) = 1 Then
    Musica = True
Else
    Musica = False
End If

End Function

Public Function GuardarOpciones()

If Musica = True Then
    Call WriteVar(IniPath & "CONFIG", "OPCIONES", "Musica", "1")
Else
    Call WriteVar(IniPath & "CONFIG", "OPCIONES", "Musica", "0")
End If

If MiniMapY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Minimapa", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Minimapa", "0")
End If

If MapNameY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "NombreMapa", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "NombreMapa", "0")
End If

If EfectosDiaY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "DiaNoche", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "DiaNoche", "0")
End If

If EfectosAlphaY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "EfectosAlpha", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "EfectosAlpha", "0")
End If

If ConsolaY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Consola", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Consola", "0")
End If

If Nombres Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Nombres", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "Nombres", "0")
End If

If FpsY = True Then
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "FPS", "1")
Else
    Call WriteVar(IniPath & "CONFIG.INI", "OPCIONES", "FPS", "0")
End If

End Function
'<---------->MSN<---------->
Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)

   Dim udtData As COPYDATASTRUCT
   Dim sBuffer As String
   Dim hMSGRUI As Long
   
   'Total length can Not be longer Then 256 characters!
   'Any longer will simply be ignored by Messenger.
   sBuffer = "\0Games\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
   
   udtData.dwData = &H547
   udtData.lpData = StrPtr(sBuffer)
   udtData.cbData = LenB(sBuffer)
   
   Do
       hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
       
       If (hMSGRUI > 0) Then
           Call SendMessage(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
       End If
       
   Loop Until (hMSGRUI = 0)

End Sub
'<------------>FIN MSN<------------->
'<------------>ANTI CHEAT ENGINE<-------->
Public Sub BuscarEngine()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject("Wscript.Shell")
Dim X As String
X = "1"
X = MiObjeto.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
If X = "0" Then
MsgBox "Debes desinstalar el CheatEngine para poder jugar."
End
End If
Set MiObjeto = Nothing
End Sub

'<------->FIN ANTI CHEAT ENGINE<----------->
'<------->TRANSPARENCIAS<-------------->

  
Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
  
Dim msg As Long
  
    msg = GetWindowLong(hwnd, GWL_EXSTYLE)
          
       If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If
  
    If Err Then
       Is_Transparent = False
    End If
  
End Function
  
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long
  
Dim msg As Long
  
On Error Resume Next
  
If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   msg = msg Or WS_EX_LAYERED
      
   SetWindowLong hwnd, GWL_EXSTYLE, msg
      
   'Establece la transparencia
   SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
  
   Aplicar_Transparencia = 0
  
End If
  
  
If Err Then
   Aplicar_Transparencia = 2
End If
  
End Function

'<--------->FIN TRANSPARENCIAS<----------->

Public Sub MP3Config()
PistaActual = 1
VisorTxt = 0
End Sub
Public Sub Reloguear()
frmMain.Socket1.HostName = IPServidor
frmMain.Socket1.RemotePort = PortServidor

If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
        If frmConnect.MousePointer = 11 Then Exit Sub

        If CheckUserData(False) = True Then
            EstadoLogin = LogCuenta
            frmMain.Socket1.Connect
        End If
End Sub

 Public Sub DibujarMiniMapa()
Dim map_x As Long, map_y As Long
 
    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
              If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color
            End If
            
            If bTecho Then

            ElseIf Not bTecho Then
                If MapData(map_x, map_y).Graphic(4).GrhIndex <> 0 Then
                    SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
                End If
            End If
        Next map_x
    Next map_y
   
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X + 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X - 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y - 1, RGB(255, 0, 0)
    SetPixel frmMain.MiniMap.hdc, UserPos.X, UserPos.Y + 1, RGB(255, 0, 0)
 
    Dim MinX As Byte
    Dim MinY As Byte
    Dim MaxX As Byte
    Dim MaxY As Byte
   
    map_x = 0
    map_y = 0
   
    MinX = UserPos.X - 5
    MaxX = UserPos.X + 5
   
    MinY = UserPos.Y - 5
    MaxY = UserPos.Y + 5
   
    For map_y = MinY To MaxY
        SetPixel frmMain.MiniMap.hdc, MinX, map_y, RGB(255, 255, 255)
    Next map_y
   
    For map_y = MinY To MaxY
        SetPixel frmMain.MiniMap.hdc, MaxX, map_y, RGB(255, 255, 255)
    Next map_y
   
    For map_x = MinX To MaxX
        SetPixel frmMain.MiniMap.hdc, map_x, MinY, RGB(255, 255, 255)
    Next map_x
 
    For map_x = MinX To MaxX
        SetPixel frmMain.MiniMap.hdc, map_x, MaxY, RGB(255, 255, 255)
    Next map_x
   
    frmMain.MiniMap.Refresh
End Sub

Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.listIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                Actual = 1
                MaxEleccion = 30
                MinEleccion = 1
            Case "Elfo"
                Actual = 101
                MaxEleccion = 113
                MinEleccion = 101
            Case "Elfo Oscuro"
                Actual = 202
                MaxEleccion = 209
                MinEleccion = 202
            Case "Enano"
                Actual = 301
                MaxEleccion = 305
                MinEleccion = 301
            Case "Gnomo"
                Actual = 401
                MaxEleccion = 406
                MinEleccion = 401
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                Actual = 70
                MaxEleccion = 76
                MinEleccion = 70
            Case "Elfo"
                Actual = 170
                MaxEleccion = 176
                MinEleccion = 170
            Case "Elfo Oscuro"
                Actual = 270
                MaxEleccion = 280
                MinEleccion = 270
            Case "Gnomo"
                Actual = 470
                MaxEleccion = 474
                MinEleccion = 470
            Case "Enano"
                Actual = 370
                MaxEleccion = 373
                MinEleccion = 370
            Case Else
                Actual = 70
                MaxEleccion = 70
                MinEleccion = 70
        End Select
End Select
 
frmCrearPersonaje.HeadView.Cls
Call DrawGrhtoHdc2(frmCrearPersonaje.HeadView.hdc, HeadData(Actual).Head(3).GrhIndex, 8, 5)
 
End Sub
Public Sub DrawGrhtoHdc2(desthDC As Long, ByVal grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
 
    On Error Resume Next
   
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
 
    If grh_index <= 0 Then Exit Sub
 
    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
        file_path = App.Path & "\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
       
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
       
       
           
        hdcsrc = CreateCompatibleDC(desthDC)
         
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
       
 
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
 
        DeleteDC hdcsrc
 
 
End Sub
 Public Sub txtReceived(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)
If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " está subastando: " & S2 & " (Cantidad: " & S3 & ") con un precio inicial de " & S4 & " monedas. Tipea /OFERTAR <cantidad> si deseas participar.", 100, 100, 120, 0, 1)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha mejorado la oferta a " & S2 & " monedas de oro. Escribe /OFERTAR cantidad para participar de la subasta.", 100, 100, 120, 0, 1)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "La subasta ha finalizado sin oferentes.", 100, 100, 120, 0, 1)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu oferta ha sido superada por otro usuario. Escribe /OFERTAR cantidad para ingresar nuevamente en la subasta.", 100, 100, 120, 0, 1)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! la subasta ha finalizado en " & S1 & " monedas de oro.", 100, 100, 120, 0, 1)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "La subasta de " & S1 & " (Cantidad: " & S2 & ") ha finalizado en " & S3 & " monedas de oro", 100, 100, 120, 0, 1)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "Los dioses le otorgan el Gran Poder a " & S1 & " en el mapa " & S2 & ".", 255, 255, 255, 1, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el gran poder.", 255, 255, 255, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estas muriendo de frio, abrigate o moriras!!.", 65, 190, 156, 0, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas envenenado, si no te curas moriras.", 0, 255, 0, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el gran poder.", 255, 255, 255, 1, 0)
If txtIndex = 12 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " es poseedor del Gran Poder en el mapa " & S2 & ".", 255, 255, 255, 1, 0)

If txtIndex = 13 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Norte pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 14 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Oeste pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 15 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Este pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 16 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Sur pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)

If txtIndex = 17 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes pertenecer a un Clan para poder atacar un Castillo.", 255, 0, 0, 1, 0)
If txtIndex = 18 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas obstruyendo la via publica, muévete o seras encarcelado!!!", 65, 190, 156, 0, 0)
If txtIndex = 19 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes atacar Castillos que le pertenecen a tu Clan.", 255, 0, 0, 1, 0)

If txtIndex = 20 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " está siendo atacado por el clan " & S1 & ".", 244, 190, 136, 1, 0)
End If

If txtIndex = 21 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & S1 & " está atacando el Castillo " & S1 & " perteneciente a tu clan!!!.", 245, 140, 135, 1, 0)
End If

If txtIndex = 22 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " está a punto de caer en manos del Clan " & S1 & "!!", 221, 34, 34, 1, 1)
End If

If txtIndex = 23 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " perteneciente a tu clan está a punto de caer en manos del Clan " & S1 & "!!!", 165, 36, 22, 1, 1)
End If

If txtIndex = 24 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & S1 & " ha conquistado el Castillo " & S2 & ".", 255, 255, 255, 1, 0)
    Call Audio.PlayWave("44.wav")
End If

If txtIndex = 25 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado al Rey del Castillo.", 255, 0, 0, 1, 0)

If txtIndex = 26 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! Los dioses han ofrendado a su clan por la mantener la conquista al Castillo " & S1 & ".", 255, 255, 255, 1, 0)
    Call Audio.PlayWave("56.wav")
End If

If txtIndex = 27 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " espera contrincante en la sala de duelos.", 91, 159, 196, 1, 0)
If txtIndex = 28 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha aceptado el duelo.", 91, 159, 196, 1, 0)
If txtIndex = 29 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha abandonado la sala de duelos.", 91, 159, 196, 1, 0)
If txtIndex = 30 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha ganado el duelo. Lleva " & S2 & " victoria(s) consecutiva.", 191, 238, 4, 1, 0)
If txtIndex = 31 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el duelo.", 191, 238, 4, 1, 0)
If txtIndex = 32 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes ingresar con mascotas a este mapa.", 255, 0, 0, 1, 0)
If txtIndex = 33 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes invocar criaturas en este mapa.", 65, 190, 156, 0, 0)
If txtIndex = 34 Then Call AddtoRichTextBox(frmMain.RecTxt, "En zona segura no puedes invocar criaturas.", 65, 190, 156, 0, 0)
If txtIndex = 35 Then Call AddtoRichTextBox(frmMain.RecTxt, "Necesitas al menos 50 skills points en domar animales para poder montar a caballo. ", 65, 190, 156, 0, 0)
If txtIndex = 36 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has terminado de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 37 Then Call AddtoRichTextBox(frmMain.RecTxt, "Recuperarás mana de a " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 38 Then Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 39 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes moverte porque estas paralizado.", 65, 190, 156, 0, 0)
If txtIndex = 40 Then Call AddtoRichTextBox(frmMain.RecTxt, "No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", 65, 190, 156, 0, 0)
If txtIndex = 41 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has salido del modo de combate. ", 65, 190, 156, 0, 0)
If txtIndex = 42 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has pasado al modo de combate. ", 65, 190, 156, 0, 0)
If txtIndex = 43 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tipea S para quitar el seguro", 255, 0, 0, 1, 0)
If txtIndex = 44 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas muy cansado para luchar.", 65, 190, 156, 0, 0)
If txtIndex = 45 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Primero selecciona el hechizo que quieres lanzar!", 65, 190, 156, 0, 0)
If txtIndex = 46 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota está muy debilitada para ser invocada. Dirígete al sacerdote mas cercano para que le de una curación.", 65, 190, 156, 0, 0)
If txtIndex = 47 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha muerto. Dirígete al sacerdote mas cercano para que reciba la curación.", 65, 190, 156, 0, 0)
If txtIndex = 48 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha sido curada.", 65, 190, 156, 0, 0)

If txtIndex = 49 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sacerdote alza sus manos, recita en voz alta unas palabras y recuperas la vida.", 65, 190, 156, 0, 0)
    Call Audio.PlayWave("100.wav")
End If

If txtIndex = 50 Then Call AddtoRichTextBox(frmMain.RecTxt, "Cerrando... Se cerrará el juego en " & S1 & " segundos...", 65, 190, 156, 0, 0)
If txtIndex = 51 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de experiencia.", 255, 0, 0, 1, 0)
If txtIndex = 52 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes ser un nivel superior a 7 para ofertar en una subasta.", 65, 190, 156, 0, 0)
End Sub

Public Sub txtReceivedB(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)

If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre el.", 65, 190, 156, 0, 0)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha ganado " & S1 & " puntos de experiencia.", 128, 255, 0, 1, 0)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu mascota ha subido al nivel " & S1 & "!", 65, 190, 156, 0, 0)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de experiencia.", 255, 0, 0, 1, 0)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado a la criatura!", 255, 0, 0, 1, 0)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "No has ganado experiencia al matar la criatura.", 255, 0, 0, 1, 0)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha envenenado!!", 255, 0, 0, 1, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado a " & S1 & "!", 255, 0, 0, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha matado!", 255, 0, 0, 1, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has subido de nivel!", 65, 190, 156, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " skillpoints.", 65, 190, 156, 0, 0)
If txtIndex = 12 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de vida.", 65, 190, 156, 0, 0)
If txtIndex = 13 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de vitalidad.", 65, 190, 156, 0, 0)
If txtIndex = 14 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de magia.", 65, 190, 156, 0, 0)
If txtIndex = 15 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu golpe maximo aumento en " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 16 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu golpe minimo aumento en " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 17 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has ganado 50 puntos de experiencia!", 255, 0, 0, 1, 0)
If txtIndex = 18 Then Call AddtoRichTextBox(frmMain.RecTxt, "Pierdes el control de tus mascotas.", 65, 190, 156, 0, 0)
If txtIndex = 19 Then Call AddtoRichTextBox(frmMain.RecTxt, "Record de usuarios conectados simultaniamente. Hay " & S1 & " usuarios.", 65, 190, 156, 0, 0)
If txtIndex = 20 Then Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> WorldSave ha concluído.", 0, 185, 0, 0, 0)
If txtIndex = 21 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha quitado " & S2 & " puntos de vida.", 255, 0, 0, 1, 0)
If txtIndex = 22 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tenes suficientes puntos de magia para lanzar este hechizo.", 65, 190, 156, 0, 0)
If txtIndex = 23 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tenes suficiente mana.", 65, 190, 156, 0, 0)
If txtIndex = 24 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mascota desinvocada.", 255, 0, 0, 1, 0)
If txtIndex = 25 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mascota invocada.", 255, 0, 0, 1, 0)
If txtIndex = 26 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!.", 65, 190, 156, 0, 0)
If txtIndex = 27 Then Call AddtoRichTextBox(frmMain.RecTxt, "Le has causado " & S1 & " puntos de daño a la criatura!", 255, 0, 0, 1, 0)
If txtIndex = 28 Then Call AddtoRichTextBox(frmMain.RecTxt, "Le has quitado " & S1 & " puntos de vida a " & S2, 255, 0, 0, 1, 0)
If txtIndex = 29 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha quitado " & S2 & " puntos de vida.", 255, 0, 0, 1, 0)
If txtIndex = 30 Then Call AddtoRichTextBox(frmMain.RecTxt, "Te estás concentrando. En " & S1 & " segundos comenzarás a meditar.", 65, 190, 156, 0, 0)
If txtIndex = 31 Then Call AddtoRichTextBox(frmMain.RecTxt, "Comenzas a meditar.", 65, 190, 156, 0, 0)
If txtIndex = 32 Then Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 33 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has sanado.", 65, 190, 156, 0, 0)
If txtIndex = 34 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! has logrado vencer a un animal sagrado, te has ganado el ingreso a la cueva de los sabios para aprender el hechizo de tu clase.", 4, 166, 179, 1, 0)
    Call Audio.PlayWave("HARP3.WAV")
End If

If txtIndex = 35 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mapa exclusivo para newbies.", 65, 190, 156, 0, 0)
If txtIndex = 36 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para ingresar a la cueva de los sabios tienes que derrotar a un animal sagrado previamente.", 65, 190, 156, 0, 0)
If txtIndex = 37 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes derrotar a un animal sagrado para poder aprender hechizos de clase.", 65, 190, 156, 0, 0)
If txtIndex = 38 Then Call AddtoRichTextBox(frmMain.RecTxt, "El hechizo no puede ser aprendido por tu clase.", 65, 190, 156, 0, 0)
If txtIndex = 39 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has vuelto a ser visible.", 65, 190, 156, 0, 0)
If txtIndex = 40 Then Call AddtoRichTextBox(frmMain.RecTxt, "No hay ninguna subasta en curso.", 65, 190, 156, 0, 0)
If txtIndex = 41 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tienes esa cantidad de monedas para ofertar.", 65, 190, 156, 0, 0)
If txtIndex = 42 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes ofertar como subastante.", 65, 190, 156, 0, 0)
If txtIndex = 43 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para subastar objetos debes ser nivel 20 o mayor.", 65, 190, 156, 0, 0)
If txtIndex = 44 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para subastar objetos debes tener al menos 20 skills points en Comerciar.", 65, 190, 156, 0, 0)
If txtIndex = 45 Then Call AddtoRichTextBox(frmMain.RecTxt, "El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM.", 65, 190, 156, 0, 0)
If txtIndex = 46 Then Call AddtoRichTextBox(frmMain.RecTxt, "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", 65, 190, 156, 0, 0)
If txtIndex = 47 Then Call AddtoRichTextBox(frmMain.RecTxt, "El pedido debe contener un mensaje que se adecue a tu problema y el mismo debe ser coherente, caso contrario los GMs no acudirán a tu pedido.", 65, 190, 156, 0, 0)
If txtIndex = 48 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Tu pregunta ha sido respondida por un GameMaster: " & S1, 19, 215, 209, 1, 0)
    Call Audio.PlayWave("59.WAV")
End If
If txtIndex = 49 Then Call AddtoRichTextBox(frmMain.RecTxt, "Pregunta respondida satisfactoriamente.", 65, 190, 156, 0, 0)
If txtIndex = 50 Then Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global se encuentra desactivado por el momento.", 65, 190, 156, 0, 0)
If txtIndex = 51 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para hablar por mensaje global debes ser nivel 10 como mínimo.", 65, 190, 156, 0, 0)
If txtIndex = 52 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global ha sido activado. Para hablar su mensaje deberá contener el prefijo "".""" & S1, 31, 36, 252, 1, 0)
    Call Audio.PlayWave("43.WAV")
End If
If txtIndex = 53 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global ha sido desactivado." & S1, 31, 36, 252, 1, 0)
    Call Audio.PlayWave("45.WAV")
End If
If txtIndex = 54 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has revelado tu posición, pierdes el efecto invisibilidad.", 65, 190, 156, 0, 0)
If txtIndex = 55 Then Call AddtoRichTextBox(frmMain.RecTxt, "Este item no puede ser subastado.", 32, 51, 233, 1, 1)
If txtIndex = 56 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has conseguido algo de leña!", 65, 190, 156, 0, 0)
If txtIndex = 57 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has pescado un lindo pez!", 65, 190, 156, 0, 0)
If txtIndex = 58 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has pescado algunos peces!", 65, 190, 156, 0, 0)
If txtIndex = 59 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has extraido algunos minerales!", 65, 190, 156, 0, 0)
If txtIndex = 60 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has obtenido un lingote!", 65, 190, 156, 0, 0)
If txtIndex = 61 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficientes minerales para hacer un lingote.", 65, 190, 156, 0, 0)
If txtIndex = 62 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has construido el objeto!.", 65, 190, 156, 0, 0)
If txtIndex = 63 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 255, 0, 0, 1, 0)
If txtIndex = 64 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has apuñalado la criatura por " & S1, 255, 0, 0, 1, 0)
If txtIndex = 65 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 255, 0, 0, 1, 0)
If txtIndex = 66 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has apuñalado a " & S1 & " por " & S2, 255, 0, 0, 1, 0)
If txtIndex = 67 Then Call AddtoRichTextBox(frmMain.RecTxt, "Te ha apuñalado " & S1 & " por " & S2, 255, 0, 0, 1, 0)
If txtIndex = 68 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Debes esperar unos momentos para tomar otra pocion!!", 65, 190, 156, 0, 0)
If txtIndex = 69 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedo cargar mas objetos.", 65, 190, 156, 0, 0)
If txtIndex = 70 Then Call AddtoRichTextBox(frmMain.RecTxt, "Solo los newbies pueden usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 71 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu clase no puede usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 72 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu clase,genero o raza no puede usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 73 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Debes aproximarte al agua para usar el barco!", 65, 190, 156, 0, 0)
If txtIndex = 74 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S", 255, 0, 0, 1, 0)
If txtIndex = 75 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando paralizado.", 65, 190, 156, 0, 0)
If txtIndex = 76 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando encarcelado.", 65, 190, 156, 0, 0)
If txtIndex = 77 Then Call AddtoRichTextBox(frmMain.RecTxt, "Ya te encuentras en el castillo.", 255, 0, 0, 1, 0)
If txtIndex = 78 Then Call AddtoRichTextBox(frmMain.RecTxt, "GM " & S1 & " acudió a " & S2 & ".", 66, 213, 157, 1, 0)

End Sub

Public Sub txtReceivedT(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)

If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas muy cansado para luchar.", 65, 190, 156, 0, 0)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡ Has formado una party !", 255, 180, 255, 0, 0)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu carisma y liderazgo no son suficientes para liderar una party.", 255, 180, 255, 0, 0)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu carisma y liderazgo no son suficientes para liderar una party.", 255, 180, 255, 0, 0)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu oferta debe superar las " & S1 & " monedas de oro.", 65, 190, 156, 0, 0)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "Hay " & S1 & " jugadores online.", 255, 255, 255, 1, 0)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "Server> " & S1 & " ha sido expulsado por posible utilización de aplicaciones ilegales.", 0, 185, 0, 0, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, "Petición de salida cancelada.", 12, 149, 250, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> " & S1 & " ha sido echado por el servidor por posible uso de SH.", 0, 185, 0, 0, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "La descripcion a cambiado.", 65, 190, 156, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has mejorado tu skill " & S1 & " en un punto!. Ahora tienes " & S2 & " pts.", 65, 190, 156, 0, 0)

End Sub
Public Sub CerrarProceso(TítuloVentana As String)
Dim hProceso As Long
Dim lEstado As Long
Dim idProc As Long
Dim winHwnd As Long

winHwnd = FindWindow(vbNullString, TítuloVentana)
If winHwnd = 0 Then
Debug.Print "El proceso no está abierto": Exit Sub
End If
Call GetWindowThreadProcessId(winHwnd, idProc)

' Obtenemos el handle al proceso
hProceso = OpenProcess(PROCESS_TERMINATE Or _
PROCESS_QUERY_INFORMATION, 0, idProc)
If hProceso <> 0 Then
' Comprobamos estado del proceso
GetExitCodeProcess hProceso, lEstado
If lEstado = STILL_ACTIVE Then
' Cerramos el proceso
If TerminateProcess(hProceso, 9) <> 0 Then
Debug.Print "Proceso cerrado"
Else
Debug.Print "No se pudo matar el proceso"
End If
End If
' Cerramos el handle asociado al proceso
CloseHandle hProceso
Else
Debug.Print "No se pudo tener acceso al proceso"
End If
End Sub
Function SoportaDisplay(DD As DirectDraw7, DDSDaTestear As DDSURFACEDESC2) As Boolean
Dim ddsd As DDSURFACEDESC2
Dim DDEM As DirectDrawEnumModes

Set DDEM = DD.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd)

Dim loopc As Integer
Dim flag As Boolean
loopc = 1
   
Do While loopc <> DDEM.GetCount And Not flag

    DDEM.GetItem loopc, ddsd
    flag = ddsd.lHeight = DDSDaTestear.lHeight _
    And ddsd.lWidth = DDSDaTestear.lWidth _
    And ddsd.ddpfPixelFormat.lRGBBitCount = _
    DDSDaTestear.ddpfPixelFormat.lRGBBitCount
    loopc = loopc + 1
Loop
SoportaDisplay = flag
End Function

Function ModosDeVideoIguales(dd1 As DDSURFACEDESC2, dd2 As DDSURFACEDESC2) As Boolean
ModosDeVideoIguales = _
    dd1.lHeight = dd2.lHeight _
    And dd1.lWidth = dd2.lWidth _
    And dd1.ddpfPixelFormat.lRGBBitCount = _
    dd2.ddpfPixelFormat.lRGBBitCount
End Function
Public Function hexMd52Asc(ByVal md5 As String) As String
    Dim i As Integer, l As String
    
    md5 = UCase$(md5)
    If Len(md5) Mod 2 = 1 Then md5 = "0" & md5
    
    For i = 1 To Len(md5) \ 2
        l = mid$(md5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(l))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    Dim i As Integer, l As String
    For i = 1 To Len(hex)
        l = mid$(hex, i, 1)
        Select Case l
            Case "A": l = 10
            Case "B": l = 11
            Case "C": l = 12
            Case "D": l = 13
            Case "E": l = 14
            Case "F": l = 15
        End Select
        
        hexHex2Dec = (l * 16 ^ ((Len(hex) - i))) + hexHex2Dec
    Next i
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Integer, l As String
    For i = 1 To Len(Text)
        l = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(l) + off) Mod 256)
    Next i
End Function


Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    Dim tempint As Integer
    
    MinLimiteX = (X \ 9 - 1) * 9
    MaxLimiteX = MinLimiteX + 26
    
    MinLimiteY = (Y \ 9 - 1) * 9
    MaxLimiteY = MinLimiteY + 26
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        tempint = MapData(loopX, loopY).CharIndex
                        Call EraseChar(MapData(loopX, loopY).CharIndex)
                        charlist(tempint).Nombre = loopX & "-" & loopY
                    End If
                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
        Next
    Next
    
    Call RefreshAllChars
End Sub

Public Sub ClearMap()
Dim loopX As Long, loopY As Long
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            'Erase NPCs
            If MapData(loopX, loopY).CharIndex > 0 Then
                If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                    Call EraseChar(MapData(loopX, loopY).CharIndex)
                End If
            End If
            
            'Erase OBJs
            MapData(loopX, loopY).ObjGrh.GrhIndex = 0
        Next
    Next
    
    Call RefreshAllChars
End Sub
Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    i = 0
    Call DarFormato(Leyenda, i, k, anti)
    i = 0
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
       i = i + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To i)
Else
    Exit Sub
End If
End Sub


Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function


Sub DibujarCartel()
If Not Cartel Then Exit Sub
Dim X As Integer, Y As Integer
X = XPosCartel + 20
Y = YPosCartel + 60
Call DDrawTransGrhIndextoSurface(BackBufferSurface, textura, XPosCartel, YPosCartel, 0, 0)
Dim j As Integer, desp As Integer

For j = 0 To UBound(LeyendaFormateada)
Dialogos.DrawText X, Y + desp, LeyendaFormateada(j), vbWhite
  desp = desp + (frmMain.font.Size) + 5
Next
End Sub


