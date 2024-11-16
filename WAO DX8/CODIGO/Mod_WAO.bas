Attribute VB_Name = "Mod_WAO"
Option Explicit

'Drag & Drop de Formularios:
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessagges Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const RGN_OR = 2

'Seguridad AntiEditados:

Public Const SegCode As String * 18 = "PII421A5X%ZAXI-"

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

'Cabezas Seleccionables
Public MiCuerpo As Integer, MiCabeza As Integer

'Seguridad AntiDobleCliente
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
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
    Call SendMessagges(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
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
                tX = GetVar(App.path & "\INIT\Update.ini", "INIT", "X")
                
        'Check Valor
        If Val(iX) <= 0 Then
            Exit Function
        End If
            DifX = iX - tX
    'Chek for Updates!
    If Not (DifX = 0) Then
        MsgBox "Hay Actualizaciones Disponibles, Se abrira el Autoupdate para Actualizar", vbInformation
            Call Shell(App.path & "\AutoUpdate.EXE")
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


Sub DameOpciones()

Dim i As Integer

If frmCrearPersonaje.lstGenero.ListIndex < 0 Or frmCrearPersonaje.lstRaza.ListIndex < 0 Then
    frmCrearPersonaje.cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.ListIndex <> -1 And frmCrearPersonaje.lstRaza.ListIndex <> -1 Then
    frmCrearPersonaje.cabeza.Enabled = True
End If

frmCrearPersonaje.cabeza.Clear
    
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                For i = 1 To 30
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 101 To 113
                    If i = 113 Then i = 201
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 202 To 209
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Enano"
                For i = 301 To 305
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Gnomo"
                For i = 401 To 406
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case Else
                UserHead = 1
                MiCuerpo = 1
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                For i = 70 To 76
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 170 To 176
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 270 To 280
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Gnomo"
                For i = 470 To 474
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                MiCuerpo = 52
            Case Else
                frmCrearPersonaje.cabeza.AddItem "70"
                MiCuerpo = 1
        End Select
End Select

frmCrearPersonaje.PlayerView.Cls

End Sub


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

'Efectos de Alpha
EfectosAlphaX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "EfectosAlpha")
If Val(EfectosAlphaX) = 1 Then
    EfectosAlphaY = True
    frmOpciones.Check3.value = 0
    frmMain.EfectosAlpha.Enabled = True
Else
    EfectosAlphaY = False
    frmOpciones.Check3.value = 1
    frmMain.EfectosAlpha.Enabled = False
End If

'Consola
ConsolaX = GetVar(IniPath & "CONFIG.INI", "OPCIONES", "Consola")
If Val(ConsolaX) = 0 Then
    Call SendData("/SEF")
    frmOpciones.Check4.value = 0
Else
    frmOpciones.Check4.value = 1
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
