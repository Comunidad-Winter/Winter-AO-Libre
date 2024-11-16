VERSION 5.00
Begin VB.Form frmCuent 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmCuent.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   9000
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   615
      Left            =   3840
      TabIndex        =   35
      Top             =   7320
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Crear PJ"
      Height          =   615
      Left            =   840
      TabIndex        =   33
      Top             =   7320
      Width           =   2775
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   2880
      MouseIcon       =   "frmCuent.frx":0CCA
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":1994
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   32
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   5040
      MouseIcon       =   "frmCuent.frx":1C2F
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":28F9
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   7200
      MouseIcon       =   "frmCuent.frx":2B94
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":385E
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   840
      MouseIcon       =   "frmCuent.frx":3AF9
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":47C3
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   7200
      MouseIcon       =   "frmCuent.frx":4A5E
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":5728
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   5040
      MouseIcon       =   "frmCuent.frx":59C3
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":668D
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   2880
      MouseIcon       =   "frmCuent.frx":6928
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":75F2
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox PJ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   840
      MouseIcon       =   "frmCuent.frx":788D
      MousePointer    =   99  'Custom
      Picture         =   "frmCuent.frx":8557
      ScaleHeight     =   1215
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta:"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   7
      Left            =   6600
      TabIndex        =   29
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   6
      Left            =   4440
      TabIndex        =   28
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   27
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   26
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   25
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   23
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   22
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   21
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PJClick"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   19
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   12
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If PJClickeado = "Nada" Then
MsgBox "Seleccione un pj"
End If
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub
Private Sub Form_Load()
Dim i As Integer
Label3.Caption = UserName
End Sub
Private Sub Image1_Click()
frmMain.Socket1.Disconnect
Unload Me
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)

If nombre(7).Caption <> "Nada" Then
    MsgBox "Tu cuenta ha llegado al máximo de personajes."
    Exit Sub
End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show vbModal
    Me.MousePointer = 11
    
End Sub

Private Sub Label7_Click()
frmBorrar.Show , frmCuent
End Sub

Private Sub Label8_Click()
End Sub

Private Sub nombre_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub
Private Sub nombre_Click(index As Integer)
PJClickeado = frmCuent.nombre(index).Caption
End Sub
Private Sub PJ_Click(index As Integer)
PJClickeado = frmCuent.nombre(index).Caption
End Sub
Private Sub PJ_dblClick(index As Integer)
If PJClickeado = "Nada" Then Exit Sub
Call Audio.PlayWave(SND_CLICK)
UserName = PJClickeado
SendData ("OOLOGI" & UserName)
Unload Me
End Sub



