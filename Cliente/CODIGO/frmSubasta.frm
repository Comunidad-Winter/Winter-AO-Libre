VERSION 5.00
Begin VB.Form frmSubasta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSubasta.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmSubasta.frx":0CCA
   ScaleHeight     =   4050
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextBox2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3720
      TabIndex        =   9
      Text            =   "1"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox TextBox1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3720
      TabIndex        =   8
      Text            =   "1"
      Top             =   2110
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2175
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   1140
      Width           =   1845
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   2680
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   1180
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   5160
      MouseIcon       =   "frmSubasta.frx":64900
      MousePointer    =   99  'Custom
      Picture         =   "frmSubasta.frx":655CA
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1485
      TabIndex        =   6
      Top             =   420
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3435
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3435
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   1920
      MouseIcon       =   "frmSubasta.frx":65AF8
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3360
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   570
      Index           =   1
      Left            =   2640
      MouseIcon       =   "frmSubasta.frx":65C4A
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   2760
      Width           =   2460
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1950
      TabIndex        =   2
      Top             =   6420
      Width           =   645
   End
End
Attribute VB_Name = "frmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[CODE]:MatuX
'
'    Le puse el iconito de la manito a los botones ^_^ y
'   le puse borde a la ventana.
'
'[END]'

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub Form_Deactivate()
'frmMain.SetFocus
End Sub


Private Sub Form_Load()
'Cargamos la interfase
'Image1(0).Picture = LoadPicture(App.Path & "\Graficos\BotónComprar.jpg")
'Image1(1).Picture = LoadPicture(App.Path & "\Graficos\Botónvender.jpg")

End Sub

Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(index).List(List1(index).listIndex) = "Nada" Or _
   List1(index).listIndex < 0 Then Exit Sub

Select Case index
   Case 1
        LastIndex2 = List1(1).listIndex
        If Not Inventario.Equipped(List1(1).listIndex + 1) Then
            SendData ("SUBA" & "," & List1(1).listIndex + 1 & "," & TextBox1.Text & "," & TextBox2.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select

List1(1).Clear

frmMain.SetFocus
Unload Me

NPCInvDim = 0
End Sub

Private Sub Image2_Click()
SendData ("FINSUB")
End Sub

Private Sub list1_Click(index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case index
    Case 1
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.Hdc, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select

Picture1.Refresh

End Sub

Private Sub TextBox1_Change()
If Val(TextBox1.Text) < 1 Then
        TextBox1.Text = 1
    End If
    
    If Val(TextBox1.Text) > MAX_INVENTORY_OBJS Then
        TextBox1.Text = 1
    End If
End Sub


