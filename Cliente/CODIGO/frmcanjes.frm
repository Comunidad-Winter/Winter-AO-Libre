VERSION 5.00
Begin VB.Form frmcanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Canjes de Puntos"
   ClientHeight    =   9675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      Height          =   495
      Left            =   2760
      Picture         =   "frmcanjes.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   30
      Top             =   5040
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Height          =   495
      Left            =   2760
      Picture         =   "frmcanjes.frx":0C42
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   28
      Top             =   3600
      Width           =   495
   End
   Begin VB.PictureBox Picture12 
      Height          =   495
      Left            =   2760
      Picture         =   "frmcanjes.frx":1904
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Picture11 
      Height          =   495
      Left            =   2760
      Picture         =   "frmcanjes.frx":2546
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Picture10 
      Height          =   495
      Left            =   2760
      Picture         =   "frmcanjes.frx":3188
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":3DCA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   8280
      Width           =   495
   End
   Begin VB.PictureBox Picture9 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":41AE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   7080
      Width           =   495
   End
   Begin VB.PictureBox Picture8 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":4DF0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   6120
      Width           =   495
   End
   Begin VB.PictureBox Picture7 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":5266
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   5040
      Width           =   495
   End
   Begin VB.PictureBox Picture5 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":5AA8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   4200
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":62EA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":6F2C
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3120
      TabIndex        =   2
      Top             =   6480
      Width           =   1935
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Para canjear tus puntos deves de hacer click en el Item deseado, una vez echo esta operacion no podras volver atras !! "
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":7B6E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "frmcanjes.frx":83B0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   5160
      MouseIcon       =   "frmcanjes.frx":8ACC
      MousePointer    =   99  'Custom
      Picture         =   "frmcanjes.frx":9796
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "Tunica Templaria (Bajos) Mago  (Sagrado): MaxDef=45/MinDef=40                                  55 Puntos"
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
      Height          =   1215
      Left            =   3360
      TabIndex        =   31
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "Armadura de la Oscuridad (Bajos) (Sagrado): MaxDef=70/MinDef=65                                  55 Puntos"
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
      Height          =   1215
      Left            =   3360
      TabIndex        =   29
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Algunos items sagrados que no se muestran en la lista se reservan para las ofertas de Donaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3120
      TabIndex        =   27
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Armadura Demonio (Sagrado): MaxDef=45/MinDef=40                                 45 Puntos"
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
      Height          =   975
      Left            =   3360
      TabIndex        =   26
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Montura Grifo (Sagrado): MaxDef=65/MinDef=65 Equitacion: 100                                        50 Puntos"
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
      Height          =   1215
      Left            =   3360
      TabIndex        =   24
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Arco Argentum (Sagrado): MinHit=15/MaxHit=25                                   50 Puntos"
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
      Height          =   975
      Left            =   3360
      TabIndex        =   22
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Para ver los puntos disponibles escribe el comando /est"
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
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Escudo Desintegrador(Sagrado): MaxDef=30/MinDef=30 35 Puntos"
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
      Height          =   855
      Left            =   600
      TabIndex        =   19
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Montura de Preclitus (Sagrado): MaxDef=50/MinDef=50 Equitacion: 100         40 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   17
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Armadura del Logouth (Sagrado): MaxDef=70/MinDef=65 55 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   15
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Vara Infernal (Sagrado): MinHit=5/MaxHit=15 40 Puntos"
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
      Height          =   855
      Left            =   600
      TabIndex        =   13
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Escudo de Torre + 1: MaxDef=24/MinDef=24 5 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   11
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Casco Bikingo: MaxDef=50/MinDef=45 10 Puntos"
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
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Gorro de Defensa Magica (+20)                     MaxDef=25/MinDef=20 15 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pendiente del Sacrificio: Con este Pendiente al morir solo se te caera el Pendiente.                     25 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Espada Argentum (Sagrado): MinHit=25/MaxHit=29                                   50 Puntos"
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
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmcanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmcanjes.Visible = False
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
Call SendData("KOTO1")
frmcanjes.Visible = False
End Sub

Private Sub Picture10_Click()
Call SendData("KITI2")
frmcanjes.Visible = False
End Sub

Private Sub Picture11_Click()
Call SendData("KITI3")
frmcanjes.Visible = False
End Sub

Private Sub Picture12_Click()
Call SendData("KITI4")
frmcanjes.Visible = False
End Sub

Private Sub Picture13_Click()
Call SendData("KITI5")
frmcanjes.Visible = False
End Sub

Private Sub Picture14_Click()
Call SendData("KITI6")
frmcanjes.Visible = False
End Sub

Private Sub Picture2_Click()
Call SendData("KOTO2")
frmcanjes.Visible = False
End Sub

Private Sub Picture3_Click()
Call SendData("KOTO3")
frmcanjes.Visible = False
End Sub

Private Sub Picture4_Click()
Call SendData("KOTO4")
frmcanjes.Visible = False
End Sub

Private Sub Picture5_Click()
Call SendData("KOTO5")
frmcanjes.Visible = False
End Sub

Private Sub Picture6_Click()
Call SendData("KITI1")
frmcanjes.Visible = False
End Sub

Private Sub Picture7_Click()
Call SendData("KOTO7")
frmcanjes.Visible = False
End Sub

Private Sub Picture8_Click()
Call SendData("KOTO8")
frmcanjes.Visible = False
End Sub

Private Sub Picture9_Click()
Call SendData("KOTO9")
frmcanjes.Visible = False
End Sub

