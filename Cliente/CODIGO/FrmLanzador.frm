VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmLanzador 
   BorderStyle     =   0  'None
   Caption         =   "Winter AO"
   ClientHeight    =   5970
   ClientLeft      =   5805
   ClientTop       =   3075
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "FrmLanzador.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Resolucion 
      BackColor       =   &H00FFFFFF&
      Height          =   200
      Left            =   360
      Picture         =   "FrmLanzador.frx":A425
      TabIndex        =   2
      Top             =   5520
      Width           =   200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustar Resolucion (800x600)"
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
      Left            =   600
      TabIndex        =   3
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
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
      Left            =   4200
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   1560
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   720
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   1200
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1440
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLanzador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_Load()
If GetVar(App.Path & "\Init\config.ini", "INIT", "Res") = 1 Then
Resolucion.value = 1
End If
Windows_Temp_Dir = General_Get_Temp_Dir
 Set MP3P = New clsMP3Player
    Call Extract_File2(MP3, App.Path & "\ARCHIVOS\", "1.mp3", Windows_Temp_Dir, False)
    MP3P.mp3file = Windows_Temp_Dir & "1.mp3"
    MP3P.stopMP3
    MP3P.playMP3
    MP3P.Volume = 1000
    FrmLanzador.Picture = LoadPicture(App.Path & _
    "\Interfaces\Lanzador.jpg")
If Winsock1.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "200.43.192.195", PortServidor ' <= direccion y puerto del server
End Sub

Private Sub Image4_Click()
Dim X
   X = ShellExecute(Me.hwnd, "Open", "www.winter-ao.com.ar", &O0, &O0, SW_NORMAL)
End Sub

Private Sub Image1_Click()
If Resolucion.value = 1 Then
Call WriteVar(App.Path & "\Init\config.ini", "INIT", "Res", "1")
Else
Call WriteVar(App.Path & "\Init\config.ini", "INIT", "Res", "0")
End If
MP3P.stopMP3
Call Main
End Sub

Private Sub Image2_Click()
Dim X
   X = ShellExecute(Me.hwnd, "Open", "www.winter-ao.com.ar/wiki/", &O0, &O0, SW_NORMAL)
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
MP3P.stopMP3
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
