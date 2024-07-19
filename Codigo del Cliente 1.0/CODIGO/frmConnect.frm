VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Updates 
      Left            =   11040
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   7980
      TabIndex        =   6
      Top             =   3450
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   11160
      Top             =   360
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      IMEMode         =   3  'DISABLE
      Left            =   3975
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3360
      Width           =   2520
   End
   Begin VB.TextBox NameTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   1020
      TabIndex        =   4
      Top             =   3360
      Width           =   2475
   End
   Begin VB.ListBox lst_servers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   4350
      ItemData        =   "frmConnect.frx":000C
      Left            =   -3000
      List            =   "frmConnect.frx":0013
      TabIndex        =   3
      Top             =   9000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   -1560
      TabIndex        =   0
      Text            =   "7666"
      Top             =   8910
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   -600
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   8910
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   1200
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   2280
   End
   Begin VB.Image imgServEspana 
      Height          =   435
      Left            =   -2280
      MousePointer    =   99  'Custom
      Top             =   9000
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   -2400
      MousePointer    =   99  'Custom
      Top             =   9000
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Image imgGetPass 
      Height          =   495
      Left            =   -4320
      MousePointer    =   99  'Custom
      Top             =   9000
      Width           =   4575
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   705
      Index           =   0
      Left            =   4200
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   2
      Left            =   -960
      MousePointer    =   99  'Custom
      Top             =   9000
      Width           =   3120
   End
   Begin VB.Image FONDO 
      Height          =   9000
      Left            =   0
      Picture         =   "frmConnect.frx":0024
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PasswordTexT As String
Option Explicit


Private Sub FONDO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        frmCargando.Show
        frmCargando.Refresh
        AddtoRichTextBox frmCargando.status, "Cerrando Winter-AO.", 0, 0, 0, 1, 0, 1
        
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        
        AddtoRichTextBox frmCargando.status, "Liberando recursos..."
        frmCargando.Refresh
        LiberarObjetosDX
        AddtoRichTextBox frmCargando.status, "Hecho", 0, 0, 0, 1, 0, 1
        AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Winter-AO!!", 0, 0, 0, 1, 0, 1
        frmCargando.Refresh
        Call UnloadAllForms
        Call ReleaseInstance
End If
End Sub


Private Sub Form_Load()
    EngineRun = False
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 
 frmConnect.Picture = LoadPicture(App.Path & "\Interfaces\Conectar.jpg")
 Call WebBrowser1.Navigate("http://wao.webcindario.com/noti.html")
 
 Call CheckUpdates
End Sub
Private Sub Image1_Click(Index As Integer)
'opera
'ahora en un rato seguimops con el curso de programacion a distancia xd cierro el vnc tenog que d

'ok1 segundito
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0

EstadoLogin = CrearCuenta

    #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = IpServidor
        frmMain.Socket1.RemotePort = "7500"
        frmMain.Socket1.Connect
    #Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect IpServidor, "7500"
    #End If
        Me.MousePointer = 11
        
        DoEvents 'importantisimo xd

        CreandoCuenta.Show
        
    Case 1
    
    #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
        If frmConnect.MousePointer = 11 Then
            Exit Sub
        End If

        
        
        'update user info
        'coco cuenta

        Cuenta = NameTxt.Text
        UserPassword = PasswordTxt
        
        Debug.Print "cuenta:" & Cuenta & vbCrLf & "pass:" & UserPassword
        ';),fran rlz.
        If CheckUserData(False) = True Then
            EstadoLogin = LogCuenta
            Me.MousePointer = 11
#If UsarWrench = 1 Then
            frmMain.Socket1.HostName = IpServidor
            frmMain.Socket1.RemotePort = "7500"
            frmMain.Socket1.Connect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
            frmMain.Winsock1.Connect IpServidor, "7500"
#End If
        End If
'lee
End Select
Exit Sub
End Sub


Private Sub Timer1_Timer()
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
ElseIf FindWindow(vbNullString, UCase$("!xspeed.net v2.0 *")) Then
ElseIf FindWindow(vbNullString, UCase$("Ao Fast Type v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Ao Life Pro Calculator v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Accelerated Flech Creator v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Amenakhte by Proko v0.01.0008")) Then
ElseIf FindWindow(vbNullString, UCase$("AutoRecorder v3.0 *")) Then
ElseIf FindWindow(vbNullString, UCase$("AO-BOT 2 v1.0 by culd")) Then
ElseIf FindWindow(vbNullString, UCase$("AO-Ice v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AO-Ice v1.1")) Then
ElseIf FindWindow(vbNullString, UCase$("AO-ZimX Cheat")) Then
ElseIf FindWindow(vbNullString, UCase$("v0.09.0010")) Then
ElseIf FindWindow(vbNullString, UCase$("AoMacro v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AoMacro2102 v1.00.0002")) Then
ElseIf FindWindow(vbNullString, UCase$("ArgenTrap v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Argentum (Dinamico) v1.02.7117")) Then
ElseIf FindWindow(vbNullString, UCase$("ArgentumH v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("ArgentumSC v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("Argentum Pesca 0.2b")) Then
ElseIf FindWindow(vbNullString, UCase$("Manchess")) Then
ElseIf FindWindow(vbNullString, UCase$("Alkon Aoh v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("ANuByS Radar v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AOItems v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AOItems Alkon v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AOItems v2.01")) Then
ElseIf FindWindow(vbNullString, UCase$("AOFlechas v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AoH2004 v0.2")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT BK-AO v1.05")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT v1.1")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT v1.2")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.3")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.4")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.5")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.6")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.7")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.8")) Then
ElseIf FindWindow(vbNullString, UCase$("AoT2006 v1.9")) Then
ElseIf FindWindow(vbNullString, UCase$("Arg")) Then
ElseIf FindWindow(vbNullString, UCase$("v0.01.0008")) Then
ElseIf FindWindow(vbNullString, UCase$("Calculos de Lucha v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Cheat by Fran v0.11.0002")) Then
ElseIf FindWindow(vbNullString, UCase$("ChiteroMegamix")) Then
ElseIf FindWindow(vbNullString, UCase$("v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("Cliente v0.9.5")) Then
ElseIf FindWindow(vbNullString, UCase$("(PokClient) v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Clicks v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("ClienteClyba v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("Dados 9.5 v0.9.5")) Then
ElseIf FindWindow(vbNullString, UCase$("Dados v0.9.5")) Then
ElseIf FindWindow(vbNullString, UCase$("DemonDark Cliente v0.01.0008")) Then
ElseIf FindWindow(vbNullString, UCase$("DemonDark Items v2.01")) Then
ElseIf FindWindow(vbNullString, UCase$("DemonDark SH v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Easy AO Makro v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Enano AO v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("EzMacros v5.0a *")) Then
ElseIf FindWindow(vbNullString, UCase$("FFF v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("FFF v1.1")) Then
ElseIf FindWindow(vbNullString, UCase$("Garchentum v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("HotKey Changer v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("LysoCliente v0.01.0008")) Then
ElseIf FindWindow(vbNullString, UCase$("macro1 v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro2005 v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro2005 v1.0.4")) Then
ElseIf FindWindow(vbNullString, UCase$("MacroCid v2.0")) Then
ElseIf FindWindow(vbNullString, UCase$("MacroCid v3.0")) Then
ElseIf FindWindow(vbNullString, UCase$("MacroCrack (macro2) v1.00.0001")) Then
ElseIf FindWindow(vbNullString, UCase$("MacroEditor v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("MacroMaker *")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro (project1) v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro Resucitar v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro Mage v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro Ocultarse v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro Flechas v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro Magic v4.1")) Then
ElseIf FindWindow(vbNullString, UCase$("Macro TiraDados v1.0 (AZ)")) Then
ElseIf FindWindow(vbNullString, UCase$("Makro v1.0 by Cavallero")) Then
ElseIf FindWindow(vbNullString, UCase$("MakroK33 (macro2) v1.00.0001")) Then
ElseIf FindWindow(vbNullString, UCase$("Makro KorveN (macro2)")) Then
ElseIf FindWindow(vbNullString, UCase$("v1.00.0001")) Then
ElseIf FindWindow(vbNullString, UCase$("MAXKro v1.2 (VF)")) Then
ElseIf FindWindow(vbNullString, UCase$("msgplus v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("MultiMacro")) Then
ElseIf FindWindow(vbNullString, UCase$("v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Multiplicador v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("MiniDoS v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Nenin v2.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Piringulete2003 v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("PikeCheat v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("PikeCheat v1.2c")) Then
ElseIf FindWindow(vbNullString, UCase$("PikeCheat v1.2.X")) Then
ElseIf FindWindow(vbNullString, UCase$("Pike-PJB v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Proxy v2.00.0005")) Then
ElseIf FindWindow(vbNullString, UCase$("PegaRapido v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("Radar dddr (vosoloco) v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Radar dddr (2005) v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Radar Silver v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("ServerEdit v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("sh v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Speeder v1.0 *")) Then
ElseIf FindWindow(vbNullString, UCase$("Speeder XP v1.01 *")) Then
ElseIf FindWindow(vbNullString, UCase$("Speeder XP v1.60 *")) Then
ElseIf FindWindow(vbNullString, UCase$("Tira Oro v9.09")) Then
ElseIf FindWindow(vbNullString, UCase$("Tira Dados v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Tuky Rlz")) Then
ElseIf FindWindow(vbNullString, UCase$("v88.88.88.88")) Then
ElseIf FindWindow(vbNullString, UCase$("Turbinas DoS Alkon v1.0")) Then
ElseIf FindWindow(vbNullString, UCase$("Volks ")) Then
ElseIf FindWindow(vbNullString, UCase$("UltraCheaut v2.0.6c")) Then
ElseIf FindWindow(vbNullString, UCase$("UltraCheat v9.09 (v1.0)")) Then
ElseIf FindWindow(vbNullString, UCase$("Cheats Taiku")) Then
ElseIf FindWindow(vbNullString, UCase$("!Speednet")) Then
ElseIf FindWindow(vbNullString, UCase$("!xSpeednet")) Then
ElseIf FindWindow(vbNullString, UCase$("MakroK33")) Then
MsgBox ("Programas externos detectados.")
End
End If

End Sub
