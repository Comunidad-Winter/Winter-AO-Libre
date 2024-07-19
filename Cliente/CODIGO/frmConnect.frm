VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Winter AO 2.0"
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
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6720
      TabIndex        =   10
      Top             =   3480
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   4920
      Width           =   180
   End
   Begin InetCtlsObjects.Inet Updates 
      Left            =   10440
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2025
      Left            =   1155
      TabIndex        =   6
      Top             =   5820
      Width           =   9540
      ExtentX         =   16828
      ExtentY         =   3572
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
      Left            =   10560
      Top             =   360
   End
   Begin VB.TextBox PasswordTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   3075
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3570
      Width           =   2340
   End
   Begin VB.TextBox NameTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2130
      TabIndex        =   4
      Top             =   3105
      Width           =   2340
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
      Height          =   135
      Left            =   -600
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   8970
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Primario (Argentino)"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label rupass 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar Usuario y Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblCuenta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dejar de recordar la Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Image BotonCerrar 
      Height          =   255
      Left            =   11760
      Top             =   0
      Width           =   255
   End
   Begin VB.Image BotonMinimizar 
      Height          =   255
      Left            =   11400
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1320
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1920
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   585
      Index           =   0
      Left            =   3480
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1965
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
Private Sub BotonCerrar_Click()
If MsgBox("¿Esta seguro/a que desea salir?", vbYesNo, "Winter AO 2.0") = vbYes Then
        MP3P.stopMP3
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
        Else
    End If
End Sub

Private Sub BotonMinimizar_Click()
Me.WindowState = vbMinimized
End Sub


Private Sub FONDO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
If MsgBox("¿Esta seguro/a que desea salir?", vbYesNo, "Winter AO 2.0") = vbYes Then
MP3P.stopMP3
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
If GetVar(App.Path & "\INIT\Cuentas.wao", "check", "A") = 1 Then
Check1.value = 1
NameTxt.Text = GetVar(App.Path & "\INIT\Cuentas.wao", "Nick", "Name")
PasswordTxt.Text = GetVar(App.Path & "\INIT\Cuentas.wao", "Passwd", "Pass")
lblCuenta.Visible = True
Else
Check1.value = 0
End If
End Sub
Private Sub Image1_Click(index As Integer)
'opera
'ahora en un rato seguimops con el curso de programacion a distancia xd cierro el vnc tenog que d

'ok1 segundito
Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0

EstadoLogin = CrearCuenta

    #If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        End If
        frmMain.Socket1.HostName = IPServidor
        frmMain.Socket1.RemotePort = PortServidor
        frmMain.Socket1.Connect
    #Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
        End If
        frmMain.Winsock1.Connect IPServidor, PortServidor
    #End If
        Me.MousePointer = 11
        
        DoEvents 'importantisimo xd

        CreandoCuenta.Show
        
    Case 1
     
If Check1.value = 1 Then
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Nick", "Name", NameTxt.Text)
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Passwd", "Pass", PasswordTxt.Text)
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Check", "A", "1")
End If
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
            frmMain.Socket1.HostName = IPServidor
            frmMain.Socket1.RemotePort = PortServidor
            frmMain.Socket1.Connect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
            frmMain.Winsock1.Connect IPServidor, PortServidor
#End If

        End If
'lee
End Select
Exit Sub
End Sub



Private Sub Timer1_Timer()
CerrarProceso ("Cheat Engine 5.1")
CerrarProceso ("Cheat Engine 5.2")
CerrarProceso ("Cheat Engine 5.3")
CerrarProceso ("CHEAT ENGINE 5.1.1")
CerrarProceso ("CHEAT ENGINE 5.0")
CerrarProceso ("Auto Pots")
CerrarProceso ("CHEAT ENGINE 5.2")
CerrarProceso ("SOLOCOVO?")
CerrarProceso ("-=[ANUBYS RADAR]=-")
CerrarProceso ("CRAZY SPEEDER 1.05")
CerrarProceso ("SET !XSPEED.NET")
CerrarProceso ("SPEEDERXP V1.80 - UNREGISTERED")
CerrarProceso ("CHEAT ENGINE 5.3")
CerrarProceso ("CHEAT ENGINE 5.1")
CerrarProceso ("A SPEEDER")
CerrarProceso ("MEMO :P")
CerrarProceso ("ORK4M VERSION 1.5")
CerrarProceso ("By Fedex")
CerrarProceso ("!Xspeeder")
CerrarProceso ("Cambia titulos")
CerrarProceso ("Cambia titulos")
CerrarProceso ("Serbio Engine")
CerrarProceso ("ReyMix Engine")
CerrarProceso ("ReyMix Engine")
CerrarProceso ("AutoClick")
CerrarProceso ("Tonner")
CerrarProceso ("Buffy The vamp Slayer")
CerrarProceso ("Blorb Slayer 1.12.552 (BETA)")
CerrarProceso ("PumaEngine3.0")
CerrarProceso ("Vicious Engine 5.0")
CerrarProceso ("AkumaEngine33")
CerrarProceso ("Spuc3ngine")
CerrarProceso ("Ultra Engine")
CerrarProceso ("Engine")
CerrarProceso ("Cheat Engine V5.4")
CerrarProceso ("Cheat Engine V4.4")
CerrarProceso ("Cheat Engine V4.4 German Add-On")
CerrarProceso ("Cheat Engine V4.3")
CerrarProceso ("Cheat Engine V4.2")
CerrarProceso ("Cheat Engine V4.1.1")
CerrarProceso ("Cheat Engine V3.3")
CerrarProceso ("Cheat Engine V3.2")
CerrarProceso ("Cheat Engine V3.1")
CerrarProceso ("Cheat Engine")
CerrarProceso ("danza engine 5.2.150")
CerrarProceso ("zenx engine")
CerrarProceso ("Macro Maker")
CerrarProceso ("Macro Maker")
CerrarProceso ("Macro Fedex")
CerrarProceso ("Macro Mage")
CerrarProceso ("Macro Fisher")
CerrarProceso ("Macro K33")
CerrarProceso ("Macro K33")
CerrarProceso ("El Chit del Geri")
CerrarProceso ("Piringulete")
CerrarProceso ("Piringulete 2003")
CerrarProceso ("Makro Tuky")
CerrarProceso ("ORK4M VERSION 1.5")
CerrarProceso ("Pts")
CerrarProceso ("Auto Aim")
CerrarProceso ("Super Saiyan")
CerrarProceso ("!xSpeed.Net -4")
CerrarProceso ("!xSpeed.Net +4")
CerrarProceso ("!xSpeed.Net 1")
CerrarProceso ("-=[ANUBYS RADAR]=-")
CerrarProceso ("SPEEDER - REGISTERED")
CerrarProceso ("RADAR SILVERAO")
CerrarProceso ("SPEEDERXP X1.60 - REGISTERED")
CerrarProceso ("SPEEDERXP X1.60 - UNREGISTERED")
CerrarProceso ("A SPEEDER V2.1")
CerrarProceso ("VICIOUS ENGINE 5.0")
CerrarProceso ("Blorb Slayer 1.12.552 (BETA)")
CerrarProceso ("Buffy The vamp Slayer")
CerrarProceso ("makro-piringulete")
CerrarProceso ("makro K33")
CerrarProceso ("makro-Piringulete 2003")
CerrarProceso ("macrocrack <gonza_vi@hotmail.com>")
CerrarProceso ("windows speeder")
CerrarProceso ("Speeder - Unregistered")
CerrarProceso ("A Speeder")
CerrarProceso ("?????")
CerrarProceso ("speeder")
CerrarProceso ("argentum-pesca 0.2b por manchess")
CerrarProceso ("speeder XP - softwrap version")
CerrarProceso ("cambia titulos de cheats by fedex")
CerrarProceso ("NEWENG OCULTO")
CerrarProceso ("Macro 2005")
CerrarProceso ("Rey Engine 5.2")
CerrarProceso ("Serbio Engine")
CerrarProceso ("Cheat Engine V5.1.1")
CerrarProceso ("Cheat Engine 5.1.1")
CerrarProceso ("Ultra Engine")
CerrarProceso ("Engine")
CerrarProceso ("Cheat Engine V5.4")
CerrarProceso ("Cheat Engine V5.3")
CerrarProceso ("Cheat Engine V5.2")
CerrarProceso ("Cheat Engine V5.1")
CerrarProceso ("Cheat Engine V5.0")
CerrarProceso ("Cheat Engine V4.4")
CerrarProceso ("Cheat Engine V4.4 German Add-On")
CerrarProceso ("Cheat Engine V4.3")
CerrarProceso ("Cheat Engine V4.2")
CerrarProceso ("Cheat Engine V4.1.1")
CerrarProceso ("Cheat Engine V3.3")
CerrarProceso ("Cheat Engine")
CerrarProceso ("Samples Macros - EZ Macros")
CerrarProceso ("Cheat Engine 5.0")
CerrarProceso ("vosoloco?")
CerrarProceso ("solocovo?")
CerrarProceso ("Summer Ao - Proxy!")
CerrarProceso ("macrocrack")
CerrarProceso ("A Speeder")
CerrarProceso ("speeder XP - softwrap version")
CerrarProceso ("aoflechas")
CerrarProceso ("Macro")
CerrarProceso ("Macro 2005")
CerrarProceso ("!xspeed.net v2.0 *")
CerrarProceso ("Ao Fast Type v1.0")
CerrarProceso ("Ao Life Pro Calculator v1.0")
CerrarProceso ("Accelerated Flech Creator v1.0")
CerrarProceso ("Amenakhte by Proko v0.01.0008")
CerrarProceso ("AutoRecorder v3.0 *")
CerrarProceso ("AO-BOT 2 v1.0 by culd")
CerrarProceso ("AO-Ice v1.0")
CerrarProceso ("AO-Ice v1.1")
CerrarProceso ("AO-ZimX Cheat")
CerrarProceso ("v0.09.0010")
CerrarProceso ("AoMacro v1.0")
CerrarProceso ("AoMacro2102 v1.00.0002")
CerrarProceso ("ArgenTrap v1.0")
CerrarProceso ("Argentum (Dinamico) v1.02.7117")
CerrarProceso ("ArgentumH v9.09")
CerrarProceso ("ArgentumSC v9.09")
CerrarProceso ("Argentum Pesca 0.2b")
CerrarProceso ("Manchess")
CerrarProceso ("Alkon Aoh v9.09")
CerrarProceso ("ANuByS Radar v1.0")
CerrarProceso ("AOItems v1.0")
CerrarProceso ("AOItems Alkon v1.0")
CerrarProceso ("AOItems v2.01")
CerrarProceso ("AOFlechas v1.0")
CerrarProceso ("AoH2004 v0.2")
CerrarProceso ("AoT BK-AO v1.05")
CerrarProceso ("AoT v1.0")
CerrarProceso ("AoT v1.1")
CerrarProceso ("AoT v1.2")
CerrarProceso ("AoT2006 v1.3")
CerrarProceso ("AoT2006 v1.4")
CerrarProceso ("AoT2006 v1.5")
CerrarProceso ("AoT2006 v1.6")
CerrarProceso ("AoT2006 v1.7")
CerrarProceso ("AoT2006 v1.8")
CerrarProceso ("AoT2006 v1.9")
CerrarProceso ("Arg")
CerrarProceso ("v0.01.0008")
CerrarProceso ("Calculos de Lucha v1.0")
CerrarProceso ("Cheat by Fran v0.11.0002")
CerrarProceso ("ChiteroMegamix")
CerrarProceso ("v9.09")
CerrarProceso ("Cliente v0.9.5")
CerrarProceso ("(PokClient) v1.0")
CerrarProceso ("Clicks v1.0")
CerrarProceso ("ClienteClyba v9.09")
CerrarProceso ("Dados 9.5 v0.9.5")
CerrarProceso ("Dados v0.9.5")
CerrarProceso ("DemonDark Cliente v0.01.0008")
CerrarProceso ("DemonDark Items v2.01")
CerrarProceso ("DemonDark SH v1.0")
CerrarProceso ("Easy AO Makro v1.0")
CerrarProceso ("Enano AO v9.09")
CerrarProceso ("EzMacros v5.0a *")
CerrarProceso ("FFF v1.0")
CerrarProceso ("FFF v1.1")
CerrarProceso ("Garchentum v1.0")
CerrarProceso ("HotKey Changer v1.0")
CerrarProceso ("LysoCliente v0.01.0008")
CerrarProceso ("macro1 v1.0")
CerrarProceso ("Macro2005 v1.0")
CerrarProceso ("Macro2005 v1.0.4")
CerrarProceso ("MacroCid v2.0")
CerrarProceso ("MacroCid v3.0")
CerrarProceso ("MacroCrack (macro2) v1.00.0001")
CerrarProceso ("MacroEditor v1.0")
CerrarProceso ("MacroMaker *")
CerrarProceso ("Macro (project1) v1.0")
CerrarProceso ("Macro Resucitar v1.0")
CerrarProceso ("Macro Mage v1.0")
CerrarProceso ("Macro Ocultarse v1.0")
CerrarProceso ("Macro Flechas v1.0")
CerrarProceso ("Macro Magic v4.1")
CerrarProceso ("Macro TiraDados v1.0 (AZ)")
CerrarProceso ("Makro v1.0 by Cavallero")
CerrarProceso ("MakroK33 (macro2) v1.00.0001")
CerrarProceso ("Makro KorveN (macro2)")
CerrarProceso ("v1.00.0001")
CerrarProceso ("MAXKro v1.2 (VF)")
CerrarProceso ("msgplus v1.0")
CerrarProceso ("MultiMacro")
CerrarProceso ("v1.0")
CerrarProceso ("Multiplicador v1.0")
CerrarProceso ("MiniDoS v1.0")
CerrarProceso ("Nenin v2.0")
CerrarProceso ("Piringulete2003 v1.0")
CerrarProceso ("PikeCheat v1.0")
CerrarProceso ("PikeCheat v1.2c")
CerrarProceso ("PikeCheat v1.2.X")
CerrarProceso ("Pike-PJB v1.0")
CerrarProceso ("Proxy v2.00.0005")
CerrarProceso ("PegaRapido v9.09")
CerrarProceso ("Radar dddr (vosoloco) v1.0")
CerrarProceso ("Radar dddr (2005) v1.0")
CerrarProceso ("Radar Silver v1.0")
CerrarProceso ("ServerEdit v1.0")
CerrarProceso ("sh v1.0")
CerrarProceso ("Tira Oro v9.09")
CerrarProceso ("Tira Dados v1.0")
CerrarProceso ("Tuky Rlz")
CerrarProceso ("v88.88.88.88")
CerrarProceso ("Turbinas DoS Alkon v1.0")
CerrarProceso ("Volks ")
CerrarProceso ("UltraCheat v2.0.6c")
CerrarProceso ("UltraCheat v9.09 (v1.0)")
CerrarProceso ("Cheats Taiku")
CerrarProceso ("VolkS TurbinaS")
CerrarProceso ("Cheat Engine")
CerrarProceso ("CheatEngine5.4 by guillo0894")
CerrarProceso ("Makrok33")
CerrarProceso ("Macro Recorder")
CerrarProceso ("MoonlightEngine")
End Sub

Private Sub lblCuenta_Click()
'Rockon
If MsgBox("¿Está seguro que desea dejar de recordar su cuenta?", vbYesNo + vbQuestion, "Twisteros Ao") = vbYes Then
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Nick", "Name", "")
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Passwd", "Pass", "")
Call WriteVar(App.Path & "\INIT\Cuentas.wao", "Check", "A", "0")
MsgBox "Su Cuenta no esta guardada"
lblCuenta.Visible = False
End If
'Rockon
End Sub
