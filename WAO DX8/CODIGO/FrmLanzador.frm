VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmLanzador 
   BorderStyle     =   0  'None
   Caption         =   "Winter AO"
   ClientHeight    =   5265
   ClientLeft      =   5805
   ClientTop       =   3075
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   Picture         =   "FrmLanzador.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1320
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   3840
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   1200
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   240
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   840
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1080
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "FrmLanzador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_load()
    FrmLanzador.Picture = LoadPicture(App.Path & _
    "\Interfaces\Lanzador.jpg")
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "200.43.193.121", "7500" ' <= direccion y puerto del server
End Sub

Private Sub Image4_Click()
Dim X
   X = ShellExecute(Me.Hwnd, "Open", "www.winter-ao.com.ar", &O0, &O0, SW_NORMAL)
End Sub

Private Sub Image1_Click()
Call Main
End Sub

Private Sub Image2_Click()
Dim X
   X = ShellExecute(Me.Hwnd, "Open", "www.winter-ao.com.ar/wiki/", &O0, &O0, SW_NORMAL)
End Sub

Private Sub Image3_Click()
Call Shell(App.Path & "\AutoUpdate.EXE")
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = vbBlue
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = vbYellow
End Sub

Private Sub Image5_Click()
End
End Sub

Private Sub Winsock1_Connect()
Label1.ForeColor = vbGreen
Label1.Caption = "Online"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label1.ForeColor = vbRed
Label1.Caption = "Offline"
End Sub
