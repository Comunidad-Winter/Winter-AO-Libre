VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   1260
   ClientTop       =   1725
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6960
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer AntiCheatEngine 
      Interval        =   9000
      Left            =   5160
      Top             =   2520
   End
   Begin VB.Timer WorkMacro 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   480
      Top             =   7800
   End
   Begin VB.Timer EfectosAlpha 
      Interval        =   5
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Timer IntervaloMacro 
      Interval        =   1500
      Left            =   2640
      Top             =   2520
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   11
      Left            =   7575
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   10
      Left            =   6900
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   6240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   5265
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   4635
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   3990
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   3345
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2415
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1770
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1125
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   8445
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   8445
      Width           =   495
   End
   Begin VB.Timer AntiMacro 
      Interval        =   20000
      Left            =   3120
      Top             =   2520
   End
   Begin VB.Timer tAntiSH 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3600
      Top             =   2520
   End
   Begin VB.PictureBox MiniMap 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   10320
      ScaleHeight     =   1305
      ScaleWidth      =   1425
      TabIndex        =   26
      Top             =   7440
      Width           =   1425
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
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
      Height          =   285
      Left            =   285
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1845
      Visible         =   0   'False
      Width           =   7905
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4560
      Top             =   2520
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
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
      Height          =   285
      Left            =   285
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1845
      Visible         =   0   'False
      Width           =   7920
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1545
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   300
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":13435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PanelDer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   8280
      Picture         =   "frmMain.frx":134BA
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   2
      Top             =   0
      Width           =   3705
      Begin VB.CommandButton DespInv 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   600
         MouseIcon       =   "frmMain.frx":1DE52
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.CommandButton DespInv 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmMain.frx":1DFA4
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   5040
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2025
         Left            =   660
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   160
         TabIndex        =   8
         Top             =   2880
         Width           =   2400
      End
      Begin VB.ListBox hlst 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   600
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   720
         Top             =   8400
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Minimapa Desactivado"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   2040
         TabIndex        =   28
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Image PicSeg 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   360
         Picture         =   "frmMain.frx":1E0F6
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   495
      End
      Begin VB.Image PicMH 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         Picture         =   "frmMain.frx":20DDA
         Stretch         =   -1  'True
         Top             =   8160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label MapName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   25
         Top             =   7170
         Width           =   1935
      End
      Begin VB.Shape ExperienciaShp 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   105
         Left            =   1440
         Top             =   1290
         Width           =   75
      End
      Begin VB.Label ItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   5460
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   2880
         Top             =   0
         Width           =   255
      End
      Begin VB.Label hambar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   6930
         Width           =   1455
      End
      Begin VB.Label agubar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   7380
         Width           =   1275
      End
      Begin VB.Label HpBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   390
         TabIndex        =   22
         Top             =   6030
         Width           =   1455
      End
      Begin VB.Label StaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   7845
         Width           =   1215
      End
      Begin VB.Label ManaBar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   435
         TabIndex        =   20
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Agilidad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   6840
         Width           =   255
      End
      Begin VB.Label Fuerza 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   6480
         Width           =   255
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   3120
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblPorcLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "33.33%"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2025
         TabIndex        =   17
         Top             =   1620
         Width           =   660
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   180
         Index           =   0
         Left            =   3000
         MouseIcon       =   "frmMain.frx":21BEC
         MousePointer    =   99  'Custom
         Top             =   3240
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image cmdMoverHechi 
         Height          =   135
         Index           =   1
         Left            =   3000
         MouseIcon       =   "frmMain.frx":21D3E
         MousePointer    =   99  'Custom
         Top             =   3480
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image cmdInfo 
         Height          =   405
         Left            =   1920
         MouseIcon       =   "frmMain.frx":21E90
         MousePointer    =   99  'Custom
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image CmdLanzar 
         Height          =   405
         Left            =   480
         MouseIcon       =   "frmMain.frx":21FE2
         MousePointer    =   99  'Custom
         Top             =   4800
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1035
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label exp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2340
         LinkTimeout     =   49
         TabIndex        =   13
         Top             =   930
         Width           =   75
      End
      Begin VB.Image Image3 
         Height          =   75
         Index           =   2
         Left            =   3600
         Top             =   9000
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   195
         Index           =   1
         Left            =   3480
         Top             =   9000
         Width           =   360
      End
      Begin VB.Image Image3 
         Height          =   315
         Index           =   0
         Left            =   1920
         Top             =   5760
         Width           =   240
      End
      Begin VB.Label GldLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   5835
         Width           =   105
      End
      Begin VB.Shape AGUAsp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         Top             =   7380
         Width           =   1410
      End
      Begin VB.Shape COMIDAsp 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   165
         Left            =   360
         Top             =   6960
         Width           =   1410
      End
      Begin VB.Shape MANShp 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   180
         Left            =   360
         Top             =   6510
         Width           =   1410
      End
      Begin VB.Shape STAShp 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         Height          =   165
         Left            =   360
         Top             =   7875
         Width           =   1410
      End
      Begin VB.Shape Hpshp 
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   375
         Top             =   6060
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Erwin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   570
         TabIndex        =   11
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         MouseIcon       =   "frmMain.frx":22134
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         MouseIcon       =   "frmMain.frx":22286
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Image InvEqu 
         Height          =   3240
         Left            =   420
         Picture         =   "frmMain.frx":223D8
         Top             =   2160
         Width           =   2880
      End
      Begin VB.Label lbCRIATURA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   120
         Left            =   555
         TabIndex        =   5
         Top             =   1965
         Width           =   30
      End
      Begin VB.Label LvlLbl 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Roman"
            Size            =   12.75
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   765
         TabIndex        =   4
         Top             =   1230
         Width           =   165
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Image PicAU 
      Appearance      =   0  'Flat
      Height          =   15
      Left            =   7440
      Picture         =   "frmMain.frx":27EB9
      Stretch         =   -1  'True
      Top             =   9000
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   6165
      Left            =   120
      Top             =   2160
      Width           =   8205
   End
   Begin VB.Label fpps 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   27
      Top             =   15
      Width           =   1335
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Standelf As Boolean
Private Const VK_SNAPSHOT = &H2C

Private Declare Sub keybd_event _
Lib "user32" ( _
ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)

Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long
Dim PuedeMacrear As Boolean

Implements DirectXEvent


Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.listIndex + 1)

Select Case Index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub
Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub

Public Sub DibujarSatelite()
PicAU.Visible = True
End Sub

Public Sub DesDibujarSatelite()
PicAU.Visible = False
End Sub

Private Sub EfectosAlpha_Timer()

If Desbanecimiento1 = True Then
    If Val(AlphaX) = 0 Then
        Desbanecimiento1 = False
        Desbanecimiento2 = True
    Else
        AlphaX = Val(AlphaX) - 5
    End If
End If

If Desbanecimiento2 = True Then
    If Val(AlphaX) = 255 Then
        Desbanecimiento1 = True
        Desbanecimiento2 = False
    Else
        AlphaX = Val(AlphaX) + 5
    End If
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub


Private Sub Image1_Click()
Call frmMenu.Show(vbModeless, frmMain) 'Stand
End Sub

Private Sub Image2_Click()
Call Audio.PlayWave(SND_CLICK)
Me.WindowState = vbMinimized

End Sub

Private Sub Image4_Click()
Call Audio.PlayWave(SND_CLICK)

        If MsgBox("¿Esta Seguro que desea Salir del juego?", vbYesNo + vbQuestion, "Winter AO") = vbYes Then
            Call SendData("/SALIR")
        Else
            Exit Sub
        End If

End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Image6_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/PARTY")
End Sub

Private Sub Image7_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/GM")
End Sub


Private Sub IntervaloMacro_Timer()
Standelf = True
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicAU_Click()
    AddtoRichTextBox frmMain.RecTxt, "Hay actualizaciones pendientes. Cierra el juego y ejecuta el autoupdate.", 255, 255, 255, False, False, False
End Sub

Private Sub PicMH_Click()
    AddtoRichTextBox frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
End Sub

Private Sub PicSeg_Click()
Call SendData("/SEG")
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub

Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

Private Sub tAntiSH_Timer()
    Static counter As Byte
    Static firstTime As Boolean
    Static TiempoAnterior As Long
    
    Dim TActual As Long
    
    If firstTime = True Then
        TActual = GetTickCount
        If TActual - TiempoAnterior > 4000 Then
            If counter = 3 Then 'La condicion tiene que darse 3 veces seguidas para que no te saque porq por ahi se trabe, etc (revisalo nico)
                MsgBox ("El sistema anticheat le ha cerrado el juego, reloguee.")
                End
            Else
                counter = counter + 1
            End If
        Else
            counter = 0
        End If
    Else
        firstTime = True
    End If
    
    TiempoAnterior = GetTickCount 'Es estatica
    
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     TIMERS                         '
''''''''''''''''''''''''''''''''''''''

Private Sub Trabajo_Timer()
    'NoPuedeUsar = False
End Sub



''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.listIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("LH" & hlst.listIndex + 1)
        Call SendData("UK" & Magia)
        UsaMacro = True
        'UserCanAttack = 0
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub


Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.listIndex + 1)
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()

    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If SendTxt.Visible Or SendCMSTXT.Visible Then Exit Sub
    If frmCustomKeys.Visible = True Then Exit Sub

        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
        
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call SendData("TAB")
                    IScombate = Not IScombate
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData("UK" & Domar)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData("UK" & Robar)
                            
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call SendData("UK" & Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
            
            Case CustomKeys.BindedKey(eKeyType.mKeyMapView)
               FrmMap.Show , frmMain
               
            Case CustomKeys.BindedKey(eKeyType.mKeyTrabajo)
                If frmMain.WorkMacro.Enabled = True Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                Else
                    frmMain.WorkMacro.Enabled = True
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Activado.", 255, 255, 255, False, False, False)
                End If
               
                    
            End Select
        Else

        End If
        
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
                
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Dim i As Integer
                    For i = 1 To 1000
                If Not FileExist(App.Path & "\Fotos\Foto" & i & ".bmp", vbNormal) Then Exit For
                    Next
                    Call Capturar_Guardar(App.Path & "/Fotos/Foto" & i & ".bmp")
                Call AddtoRichTextBox(frmMain.RecTxt, "Foto" & i & ".bmp Guardada en la Carpeta Fotos", 255, 255, 255, False, False, False)
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
            End If
            
            Case vbKeyF1:
                If Standelf Then
                    Call DoAccionTecla("F1")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF2:
                If Standelf Then
                    Call DoAccionTecla("F2")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF3:
                If Standelf Then
                    Call DoAccionTecla("F3")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF4:
                If Standelf Then
                    Call DoAccionTecla("F4")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF5:
                If Standelf Then
                    Call DoAccionTecla("F5")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF6:
                If Standelf Then
                    Call DoAccionTecla("F6")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF7:
                If Standelf Then
                    Call DoAccionTecla("F7")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF8:
                If Standelf Then
                    Call DoAccionTecla("F8")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF9:
                If Standelf Then
                    Call DoAccionTecla("F9")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF10:
                If Standelf Then
                    Call DoAccionTecla("F10")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If
            Case vbKeyF11:
                If Standelf Then
                    Call DoAccionTecla("F11")
                    Standelf = False
                ElseIf Not Standelf Then
                    Exit Sub
                End If

        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                SendTxt.SetFocus
                End If
                
        End Select
End Sub

Private Sub Form_load()
    
        With frmMain
        .Width = 12000
        .Height = 9000
    End With
    
    frmMain.Picture = LoadPicture(App.Path & _
    "\Interfaces\todo.jpg")
     If GetVar(IniPath & "Macros.ini", "OPCIONES", "Minimapa") = "No" Then
 MiniMap.Visible = False
 Label2.Visible = True
 Else
 MiniMap.Visible = True
 Label2.Visible = False
 End If 'minimapa desactivable
 
    Dim result As Long
result = SetWindowLong(RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    PanelDer.Picture = LoadPicture(App.Path & _
    "\Interfaces\Principalnuevo_sin_energia.jpg")
    
    InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\Centronuevoinventario.jpg")
    
   Me.Left = 0
   Me.Top = 0
   
Unload Form2

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub


Private Sub Image3_Click(Index As Integer)
AddtoRichTextBox frmMain.RecTxt, "No esta permitido Tirar oro !!! Puedes eso usa el comando /DARORO.", 255, 255, 255, 1, 0
End Sub

Private Sub Label1_Click()
Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & j

            ElseIf UCase$(stxtbuffer) = "/HACERTORNEO" Then
                frmConsolaTorneo.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
            
            ElseIf Left$(stxtbuffer, 1) = ";" Then
            Call SendData(":" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.adentrocuenta Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.LogCuenta Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.CrearCuenta Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.Txtcorreo
        frmMain.Socket1.Write cmd, Len(cmd)
    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.tAntiSH.Enabled = False
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    LastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub


Private Sub Timer1_Timer()
Call BuscarEngine
End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
    
    Debug.Print "Winsock Connect"
    
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    
    Second.Enabled = True
    
    
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.CrearCuenta Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("gIvEmEvAlcOde")
    'Else
    ElseIf EstadoLogin = E_MODO.RecuperarPass Then
        Dim cmd As String
        cmd = "PASSRECO" & frmRecuperar.txtNombre.Text & "~" & frmRecuperar.Txtcorreo
        'frmMain.Socket1.Write cmd$, Len(cmd$)
        'Call SendData(cmd$)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If

Private Sub AntiMacro_Timer()
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Or FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Or FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Or FindWindow(vbNullString, UCase$("SOLOCOVO?")) Or FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Or FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Or FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Or FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Or FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Or FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Or FindWindow(vbNullString, UCase$("A SPEEDER")) Or FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Or FindWindow(vbNullString, UCase$("SERBIO ENGINE 1.0")) Then
MsgBox ("Programa Externo Detectado," & "Winter AO" & " Se Cerrara.")
Winsock1.Close
End
End If
End Sub
Private Sub Capturar_Guardar(Path As String)
Clipboard.Clear
keybd_event VK_SNAPSHOT, 1, 0, 0
DoEvents
    If Clipboard.GetFormat(vbCFBitmap) Then
            SavePicture Clipboard.GetData(vbCFBitmap), Path
            'MsgBox " Captura generada en: " & Path, vbInformation
    'Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    Else
    MsgBox " Error ", vbCritical
    End If
End Sub

Private Sub Macros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    frmMacros.Show vbModeless, frmMain
Else
    If Standelf Then
        Call DoAccionTecla("F" & Index)
        Standelf = False
    ElseIf Not Standelf Then
        Exit Sub
    End If
End If
End Sub

Private Sub WorkMacro_Timer()

If Me.ItemName.Caption = "Hacha de Leñador" Or Me.ItemName.Caption = "Piquete de Minero" Or Me.ItemName.Caption = "Caña de Pescar" Or Me.ItemName.Caption = "Minerales de Hierro" Or Me.ItemName.Caption = "Minerales de Plata" Or Me.ItemName.Caption = "Minerales de Oro" Or Me.ItemName.Caption = "Red de Pesca " Then
    SendData "USA" & Inventario.SelectedItem
    SendData "WLC" & tX & "," & tY & "," & UsingSkill
Else
    AddtoRichTextBox frmMain.RecTxt, "No Puedes Usar el Macro Con Este item!", 255, 255, 255, False, False, False
    frmMain.WorkMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
    Exit Sub
End If

End Sub
