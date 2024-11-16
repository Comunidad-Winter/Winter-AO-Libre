VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   Caption         =   "Winter-AO Server"
   ClientHeight    =   5355
   ClientLeft      =   1965
   ClientTop       =   1830
   ClientWidth     =   10020
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin InetCtlsObjects.Inet IP 
      Left            =   9360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrSubasta 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer TimeFunctionAutomatic 
      Interval        =   60000
      Left            =   3960
      Top             =   600
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Funciones Automaticas"
      Height          =   2415
      Left            =   5160
      TabIndex        =   22
      Top             =   2640
      Width           =   3975
      Begin VB.TextBox TimeLimpiezaMundo 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Text            =   "15"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TimeBackUp 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Text            =   "500"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TimeGrabarPJ 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Text            =   "50"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TimeActNPC 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Text            =   "30"
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TimeRec 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Text            =   "20"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label TimeLM 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   39
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Restante para Limpieza de Mundo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "Tiempo Restante para BackUp:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label TimeBU 
         BackColor       =   &H00808080&
         Caption         =   "00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Tiempo Restante para Grabar Personajes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label TimeGP 
         BackColor       =   &H00808080&
         Caption         =   "00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808080&
         Caption         =   "Tiempo Restante para actualizar NPC's:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label TimeNP 
         BackColor       =   &H00808080&
         Caption         =   "00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3000
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   "BackU,  Grabar PJ,  Act. NPC's, Limpieza de Mundo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "Tiempo Restante para Recordatorios:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         Caption         =   "Tiempo de Recordatorios (Minutos) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label TimeRecordatorio 
         BackColor       =   &H00808080&
         Caption         =   "00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cambiar de propaganda  (Esto es automatico)"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   2280
      Width           =   4815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar cuadro de propagandas en clientes"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Control de Propagandas"
      Height          =   1935
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   4815
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label NumPropaganda 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de propaganda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label NomPropaganda 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de propaganda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label URLPropaganda 
         BackStyle       =   0  'Transparent
         Caption         =   "URL de propaganda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label RedicPropaganda 
         BackStyle       =   0  'Transparent
         Caption         =   "Propaganda rediccionable:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label URLARediccionar 
         BackStyle       =   0  'Transparent
         Caption         =   "URL a rediccionar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   4935
      End
   End
   Begin VB.Timer Propaganda 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer tGuerra 
      Interval        =   60000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer AutoClean 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3480
      Top             =   600
   End
   Begin VB.Timer Climatologia 
      Interval        =   60000
      Left            =   2040
      Top             =   120
   End
   Begin VB.Timer tGranPoder 
      Interval        =   60000
      Left            =   3480
      Top             =   120
   End
   Begin VB.CommandButton Command85 
      Caption         =   "Limpiar"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   4935
   End
   Begin VB.ListBox ListadoM 
      Height          =   2160
      ItemData        =   "frmMain.frx":08CA
      Left            =   120
      List            =   "frmMain.frx":09D6
      TabIndex        =   10
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3000
      Top             =   120
   End
   Begin VB.Timer CmdExec 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   600
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   2040
      Top             =   600
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3000
      Top             =   600
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1560
      Top             =   600
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label State 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado del Mundo:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label zoom 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Escuch 
      BackColor       =   &H00808080&
      Caption         =   "Label2"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5160
      TabIndex        =   0
      Top             =   5160
      Width           =   4005
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumeroPropaganda As String
Dim TiempoPropagandas As String
Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function



Private Sub Auditoria_Timer()
On Error GoTo errhand

Call PasarSegundo 'sistema de desconexion de 10 segs

Call ActualizaEstadisticasWeb
Call ActualizaStatsES



Exit Sub

errhand:
Call LogError("Error en Timer Auditoria. Err: " & Err.Description & " - " & Err.Number)
End Sub

Private Sub AutoClean_Timer()
Dim Tiempo As Integer
    Dim Y As Integer
        Dim X As Integer
                Dim MapaActual As Integer
    Tiempo = val(Tiempo) + 1
    
    If val(Tiempo) = 12 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Limpieza del Mundo en 3 Minutos." & FONTTYPE_SERVER)
    ElseIf val(Tiempo) = 13 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Limpieza del Mundo en 2 Minutos." & FONTTYPE_SERVER)
    ElseIf val(Tiempo) = 14 Then
        Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Limpieza del Mundo en 1 Minuto." & FONTTYPE_SERVER)
    ElseIf val(Tiempo) = 15 Then
        Call LimpiarMundo
    End If
End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsSocketReset As Long
Static MinsPjesSave As Long
Static MinutosNumUsersCheck As Long

Dim i As Integer
Dim num As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        
        Horas = 0
        
    End If
    MinsRunning = 0
End If

    
Minutos = Minutos + 1

'�?�?�?�?�?�?�?�?�?�?�
Call ModAreas.AreasOptimizacion
'�?�?�?�?�?�?�?�?�?�?�

'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

#If UsarQueSocket = 1 Then
' ok la cosa es asi, este cacho de codigo es para
' evitar los problemas de socket. a menos que estes
' seguro de lo que estas haciendo, te recomiendo
' que lo dejes tal cual est�.
' alejo.
MinsSocketReset = MinsSocketReset + 1
' cada 1 minutos hacer el checkeo
If MinsSocketReset >= 5 Then
    MinsSocketReset = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then
            If UserList(i).Counters.IdleCount > ((IntervaloCerrarConexion * 2) / 3) Then
                Call CloseSocket(i)
            End If
        End If
    Next i
    'Call ReloadSokcet
    
    Call LogCriticEvent("NumUsers: " & NumUsers & " WSAPISock2Usr: " & WSAPISock2Usr.Count)
End If
#End If

MinutosNumUsersCheck = MinutosNumUsersCheck + 1

If MinutosNumUsersCheck >= 2 Then
    MinutosNumUsersCheck = 0
    num = 0
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).flags.UserLogged Then
            num = num + 1
        End If
    Next i
    If num <> NumUsers Then
        NumUsers = num
        'Call SendData(SendTarget.ToAdmins, 0, 0, "Servidor> Error en NumUsers. Contactar a algun Programador." & FONTTYPE_SERVER)
        Call LogCriticEvent("Num <> NumUsers")
    End If
End If

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, 0, "||Worldsave en 1 minuto ..." & FONTTYPE_VENENO)
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.Description)

End Sub






Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub CmdExec_Timer()
Dim i As Integer
Static N As Long

On Error Resume Next ':(((

N = N + 1

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
        If N >= 10 Then
            If UserList(i).ColaSalida.Count > 0 Then ' And UserList(i).SockPuedoEnviar Then
    #If UsarQueSocket = 1 Then
                Call IntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 0 Then
    '            Call WrchIntentarEnviarDatosEncolados(i)
    '#ElseIf UsarQueSocket = 2 Then
    '            Call ServIntentarEnviarDatosEncolados(i)
    #ElseIf UsarQueSocket = 3 Then
        'NADA, el control deberia ocuparse de esto!!!
        'si la cola se llena, dispara un on close
    #End If
            End If
        End If
    End If
Next i

If N >= 10 Then
    N = 0
End If

Exit Sub
hayerror:

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> " & BroadMsg.Text & FONTTYPE_SERVER)
End Sub

Private Sub Command7_Click()
    TimeRecordatorio.Caption = TimeRec.Text
    TimeBU.Caption = TimeBackUp.Text
    TimeGP.Caption = TimeGrabarPJ.Text
    TimeNP.Caption = TimeActNPC.Text
    TimeLM.Caption = TimeLimpiezaMundo.Text
    Call SendData(ToAll, 0, 0, "�" & "OK")
    If val(TimeBU.Caption) <= "5" Then
    Call SendData(ToAll, 0, 0, "�" & "Tiempo restante para el BackUp: " & TimeBU.Caption & " Min.")
    End If
    MsgBox "Actualizado con �xito!"
End Sub

Private Sub Command85_Click()
 ListadoM.Clear
zoom.Caption = ""
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi(frmMain.hWnd)
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Call DescargaNpcsDat


Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub

Private Sub FX_Timer()
On Error GoTo hayerror

Call SonidosMapas.ReproducirSonidosDeMapas

Exit Sub
hayerror:

End Sub

Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer

Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long

'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If

If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If

If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]

On Error GoTo hayerror

 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '�User valido?

      If UserList(iUserIndex).ConnIDValida And UserList(iUserIndex).flags.UserLogged Then
         
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         
         UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0

         
         Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.X, UserList(iUserIndex).Pos.Y)
         
                
         If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
         If UserList(iUserIndex).flags.Ceguera = 1 Or _
            UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
         
          
         If UserList(iUserIndex).flags.Muerto = 0 Then
               
               '[Consejeros]
               If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoFrio(iUserIndex)
               If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
               If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = PlayerType.User Then Call EfectoVeneno(iUserIndex, bEnviarStats)
               If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
               If UserList(iUserIndex).flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                
               Call DuracionPociones(iUserIndex)
                
               Call HambreYSed(iUserIndex, bEnviarAyS)
                
               If Lloviendo Then
                    If Not Intemperie(iUserIndex) Then
                        If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                        'No esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                            
                        ElseIf UserList(iUserIndex).flags.Descansar Then
                        'esta descansando
                            
                            Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                            Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                            If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                                 'termina de descansar automaticamente
                            If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                    Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                    Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                    UserList(iUserIndex).flags.Descansar = False
                            End If
                            
                        End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    'No esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                        
                    ElseIf UserList(iUserIndex).flags.Descansar Then
                    'esta descansando
                        
                        Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASH" & UserList(iUserIndex).Stats.MinHP): bEnviarStats = False
                        Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                        If bEnviarStats Then Call SendData(SendTarget.ToIndex, iUserIndex, 0, "ASS" & UserList(iUserIndex).Stats.MinSta): bEnviarStats = False
                             'termina de descansar automaticamente
                        If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                            UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                Call SendData(SendTarget.ToIndex, iUserIndex, 0, "DOK")
                                Call SendData(SendTarget.ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                UserList(iUserIndex).flags.Descansar = False
                        End If
                        
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If
               
               If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)

               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
     Else 'no esta logeado?
     'UserList(iUserIndex).Counters.IdleCount = 0
     '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
     'de dados :)
        
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call CloseSocket(iUserIndex)
        End If
        
     End If 'UserLogged

   End If

   Next iUserIndex

'[Alejo]
If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If

If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If

If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If

Exit Sub
hayerror:
LogError ("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex)
'[/Alejo]
  'DoEvents
End Sub

Private Sub ListadoM_Click()
zoom.Caption = ListadoM.Text
End Sub

Private Sub mnuCerrar_Click()
If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. �Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
      Call QuitarIconoSystray '[/About] Quitamos el �cono del systray
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If Not FileExist(App.Path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"
End If

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "Winter-AO Server"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub




Private Sub Premio_Timer()

End Sub

Private Sub tGranPoder_Timer()
Static Minutoss As Integer
Minutoss = Minutoss + 1
If Minutoss >= 5 Then
    Minutoss = 0
    If GranPoder = 0 Then
        OtorgarGranPoder (0)
    Else
        Call SendData(SendTarget.ToAll, GranPoder, 0, "||" & UserList(GranPoder).name & " con Gran Poder en el mapa " & UserList(GranPoder).Pos.Map & "." & FONTTYPE_GUILD)
        Call SendData(SendTarget.ToPCArea, GranPoder, UserList(GranPoder).Pos.Map, "CFX" & UserList(GranPoder).Char.CharIndex & "," & FXWARP & "," & 0)
    End If
Else
    If GranPoder > 0 Then Call SendData(SendTarget.ToPCArea, GranPoder, UserList(GranPoder).Pos.Map, "CFX" & UserList(GranPoder).Char.CharIndex & "," & FXWARP & "," & 0)
End If
End Sub

Private Sub tGuerra_Timer()
Call TimeGuerra
End Sub
 


Private Sub TimeFunctionAutomatic_Timer()
On Error Resume Next
TimeBU.Caption = val(TimeBU.Caption) - 1
TimeGP.Caption = val(TimeGP.Caption) - 1
TimeNP.Caption = val(TimeNP.Caption) - 1
If val(TimeBU.Caption) <= "5" Then
Call SendData(ToAll, 0, 0, "�" & "Tiempo restante para el BackUp: " & TimeBU.Caption & " Min.")
End If
If TimeBU.Caption = "0" Then
    TimeBU.ForeColor = vbRed
    TimeBU.Caption = "Wait..."
    Call SendData(ToAll, 0, 0, "�" & "BackUp en proceso...")
    Call DoBackUp
    TimeBU.ForeColor = vbGreen
    TimeBU.Caption = "OK!"
    TimeBU.Caption = TimeBackUp.Text
    TimeBU.ForeColor = vbRed
    Call SendData(ToAll, 0, 0, "�" & "OK")
End If
If TimeGP.Caption = "0" Then
    Call SendData(ToAll, 0, 0, "�" & "Grabando Personajes...")
    TimeGP.ForeColor = vbRed
    TimeGP.Caption = "Wait..."
    Call mdParty.ActualizaExperiencias
    Call GuardarUsuarios
    TimeGP.ForeColor = vbGreen
    TimeGP.Caption = "OK!"
    TimeGP.Caption = TimeGrabarPJ.Text
    TimeGP.ForeColor = vbRed
    Call SendData(ToAll, 0, 0, "�" & "OK")
End If
If TimeNP.Caption = "0" Then
    Call SendData(ToAll, 0, 0, "�" & "Actualizando NPC's...")
    TimeNP.ForeColor = vbRed
    TimeNP.Caption = "Wait..."
    Call ReSpawnOrigPosNpcs
    TimeNP.ForeColor = vbGreen
    TimeNP.Caption = "OK!"
    TimeNP.Caption = TimeActNPC.Text
    TimeNP.ForeColor = vbRed
    Call SendData(ToAll, 0, 0, "�" & "OK")
    End If
    If TimeLM.Caption = "0" Then
    Call SendData(ToAll, 0, 0, "�" & "Limpiando Mundo...")
    TimeLM.ForeColor = vbRed
    TimeLM.Caption = "Wait..."
    Call LimpiarMundo
    TimeLM.ForeColor = vbGreen
    TimeLM.Caption = "OK!"
    TimeLM.Caption = TimeLimpiezaMundo.Text
    TimeLM.ForeColor = vbRed
    Call SendData(ToAll, 0, 0, "�" & "OK")
End If
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            e_p = esPretoriano(NpcIndex)
            If e_p > 0 Then
                If Npclist(NpcIndex).flags.Paralizado = 1 Then Call EfectoParalisisNpc(NpcIndex)
                Select Case e_p
                    Case 1  ''clerigo
                        Call PRCLER_AI(NpcIndex)
                    Case 2  ''mago
                        Call PRMAGO_AI(NpcIndex)
                    Case 3  ''cazador
                        Call PRCAZA_AI(NpcIndex)
                    Case 4  ''rey
                        Call PRREY_AI(NpcIndex)
                    Case 5  ''guerre
                        Call PRGUER_AI(NpcIndex)
                End Select
            Else
                ''ia comun
                If Npclist(NpcIndex).flags.Paralizado = 1 Then
                      Call EfectoParalisisNpc(NpcIndex)
                Else
                     'Usamos AI si hay algun user en el mapa
                     If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                     End If
                     mapa = Npclist(NpcIndex).Pos.Map
                     If mapa > 0 Then
                          If MapInfo(mapa).NumUsers > 0 Then
                                  If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                  End If
                          End If
                     End If
                     
                End If
            End If
        End If
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i

End Sub




Private Sub Timer2_Timer()

End Sub

Private Sub tmrSubasta_Timer()
On Error GoTo errhandler

Static Minutos As Integer

If Subastando = True Then

    Minutos = Minutos + 1

    If Minutos = 2 Then
        Call ResolverSubasta
        Minutos = 0
    End If

End If

Exit Sub
errhandler:
    Call LogError("Error en TimerSubasta " & Err.Number & ": " & Err.Description)
End Sub

Private Sub tPiqueteC_Timer()
On Error GoTo errhandler
Static Segundos As Integer
Dim NuevaA As Boolean
Dim NuevoL As Boolean
Dim GI As Integer

Segundos = Segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
            
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.X, UserList(i).Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    Call SendData(SendTarget.ToIndex, i, 0, "||Estas obstruyendo la via publica, mu�vete o seras encarcelado!!!" & FONTTYPE_INFO)
                    If UserList(i).Counters.PiqueteC > 23 Then
                            UserList(i).Counters.PiqueteC = 0
                            Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
            Else
                    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If

            'ustedes se preguntaran que hace esto aca?
            'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
            'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
            'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable

            GI = UserList(i).GuildIndex
            If GI > 0 Then
                NuevaA = False
                NuevoL = False
                If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
                    Call SendData(SendTarget.ToIndex, i, 0, "||Has sido expulsado del clan. �El clan ha sumado un punto de antifacci�n!" & FONTTYPE_GUILD)
                End If
                If NuevaA Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||�El clan ha pasado a tener alineaci�n neutral!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan cambio de alineacion!")
                End If
                If NuevoL Then
                    Call SendData(SendTarget.ToGuildMembers, GI, 0, "||�El clan tiene un nuevo l�der!" & FONTTYPE_GUILD)
                    Call LogClanes("El clan tiene nuevo lider!")
                End If
            End If

            If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
            
    End If
Next i

If Segundos >= 18 Then Segundos = 0
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.Description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).IP = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.Description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
Dim t() As String
Dim LoopC As Long
Dim RD As String
On Error GoTo errorh
If UserList(MiDato).ConnID <> UserList(MiDato).ConnID Then
    Call LogError("Recibi un read de un usuario con ConnId alterada")
    Exit Sub
End If

RD = StrConv(Datos, vbUnicode)

'call logindex(MiDato, "Read. ConnId: " & ID & " Midato: " & MiDato & " Dato: " & RD)

UserList(MiDato).RDBuffer = UserList(MiDato).RDBuffer & RD

t = Split(UserList(MiDato).RDBuffer, ENDC)
If UBound(t) > 0 Then
    UserList(MiDato).RDBuffer = t(UBound(t))
    
    For LoopC = 0 To UBound(t) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If t(LoopC) <> "" Then
                If Not UserList(MiDato).CommandsBuffer.Push(t(LoopC)) Then
                    Call LogError("Cerramos por no encolar. Userindex:" & MiDato)
                    Call CloseSocket(MiDato)
                End If
            End If
        Else ' no encolamos los comandos (MUY VIEJO)
              If UserList(MiDato).ConnID <> -1 Then
                Call HandleData(MiDato, t(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If
Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.Description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Standelfito 2008
Private Sub Climatologia_Timer()
'Sumamos el Tiempo:
 
If Hour(Now) = 6 Then
            Call Ma�ana
        End If
        If Hour(Now) = 12 Then
            Call Dia
        End If
        If Hour(Now) = 18 Then
          Call Tarde
        End If
       If Hour(Now) = 20 Then
           Call Noche
        End If
        
'Actualizamos el Label:
frmMain.State.Caption = "Estado del Mundo: " & Clima
End Sub
'/Standelfito 2008
    Private Sub Propaganda_Timer()

    NumeroPropaganda = NumeroPropaganda + 1
    TiempoPropagandas = val(GetVar(App.Path & "\Dat\Propagandas.ini", "INIT", "TiempoPropagandas"))
    If NumeroPropaganda = TiempoPropagandas Then
    Call Propagandas
    NumeroPropaganda = 0
    End If
    End Sub
   
   Private Sub Command3_Click()

 Call Propagandas
   End Sub
 Private Sub Command4_Click()

 Call SendData(ToAll, 0, 0, "=" & "Apagar")
 End Sub
 Private Sub Form_Load()
     TimeRecordatorio.Caption = TimeRec.Text
    TimeBU.Caption = TimeBackUp.Text
    TimeGP.Caption = TimeGrabarPJ.Text
    TimeNP.Caption = TimeActNPC.Text
    TimeLM.Caption = TimeLimpiezaMundo.Text
    Call SendData(ToAll, 0, 0, "�" & "OK")
    If val(TimeBU.Caption) <= "5" Then
    Call SendData(ToAll, 0, 0, "�" & "Tiempo restante para el BackUp: " & TimeBU.Caption & " Min.")
    End If
Call Propagandas
 End Sub

