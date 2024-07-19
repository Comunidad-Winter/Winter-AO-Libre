VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00004040&
   BorderStyle     =   0  'None
   Caption         =   "Cuentas"
   ClientHeight    =   9000
   ClientLeft      =   3615
   ClientTop       =   3150
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   8955
      Picture         =   "Form2.frx":1714C
      ScaleHeight     =   2070
      ScaleWidth      =   285
      TabIndex        =   16
      Top             =   3225
      Width           =   285
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7290
      Picture         =   "Form2.frx":176B6
      ScaleHeight     =   300
      ScaleWidth      =   1845
      TabIndex        =   15
      Top             =   5160
      Width           =   1845
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   7200
      Picture         =   "Form2.frx":17C37
      ScaleHeight     =   2115
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   3240
      Width           =   315
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7260
      Picture         =   "Form2.frx":1819D
      ScaleHeight     =   285
      ScaleWidth      =   1965
      TabIndex        =   13
      Top             =   3240
      Width           =   1965
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   325
      Index           =   8
      Left            =   960
      Picture         =   "Form2.frx":1874E
      TabIndex        =   8
      Top             =   6720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   7
      Left            =   960
      Picture         =   "Form2.frx":2A02E
      TabIndex        =   7
      Top             =   6240
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   960
      Picture         =   "Form2.frx":3B90E
      TabIndex        =   6
      Top             =   5760
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   960
      Picture         =   "Form2.frx":4D1EE
      TabIndex        =   5
      Top             =   5160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   960
      Picture         =   "Form2.frx":5EACE
      TabIndex        =   4
      Top             =   4680
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   960
      Picture         =   "Form2.frx":703AE
      TabIndex        =   3
      Top             =   4080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   960
      Picture         =   "Form2.frx":81C8E
      TabIndex        =   2
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   7200
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   1
      Top             =   3240
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   960
      Picture         =   "Form2.frx":9356E
      TabIndex        =   0
      Top             =   2880
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Mapa 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      TabIndex        =   11
      Top             =   9000
      Width           =   15
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1635
      Left            =   8280
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   1620
      Left            =   5520
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   600
      Top             =   8160
      Width           =   1305
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_load()
'Unload frmConnect
Label1 = "Cargando... " & vbNewLine & " Clickea en un PJ"

DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = ""
End Sub
Private Sub Image1_Click()
    #If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
            End If
    #Else
            If frmMain.Winsock1.State <> sckClosed Then
                frmMain.Winsock1.Close
            End If
    #End If
    frmConnect.Show
    Unload Me
End Sub

Private Sub Image2_Click()
    If NumPjs >= 8 Then
        MsgBox "No podes crear mas personajes con esta cuenta, si queres algun otro, borra uno y crea"
        Exit Sub
    End If
    
    frmCrearPersonaje.Show
    Form2.Hide
End Sub

Private Sub Image3_Click()
    Dim Cosa As String
    If NumPjs <= 0 Then Exit Sub
    
    For i = 1 To NumPjs
        If Option1(i).Visible = True Then
            If Option1(i).value = True Then
                UserName = Option1(i).Caption
                Exit For
            End If
        End If
    Next i
    
    EstadoLogin = adentrocuenta
    Call Login(ValidarLoginMSG(CInt(bRK)))
    Form2.Visible = False

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Labe2_Click()

End Sub

Private Sub Option1_Click(Index As Integer)
    Dim SR As RECT
    Dim grhindex As Long
    
    SR.Left = 0
    SR.Top = 0
    SR.Right = 152
    SR.bottom = 152
    
    Label2.Caption = PjCuenta.Nombre(Index)
    Mapa.Caption = PjCuenta.Pos(Index)
    Oro.Caption = PjCuenta.Oro(Index)
    
    SR.Left = 0
    SR.Top = 0
    SR.Right = 152
    SR.bottom = 152
    
    Picture1.Refresh
    
End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Mori As String
If PjCuenta.muertO(Index) = 1 Then Mori = "Sí"
If PjCuenta.muertO(Index) = 0 Then Mori = "No"


End Sub

