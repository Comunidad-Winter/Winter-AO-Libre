VERSION 5.00
Begin VB.Form frmPasswdSinPadrinos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5160
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   315
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1920
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   300
      Left            =   3840
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Seguridad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   4905
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear el PJ:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmPasswdSinPadrinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AntiKey As String


Function CheckDatos() As Boolean

'Standelf
If Text1.Text <> AntiKey Then
    MsgBox "El Codigo Ingresado no es el correcto."
    Call GenerateKey
    Label7.Caption = AntiKey
    Exit Function
End If

CheckDatos = True

End Function

Private Sub Form_Load()

Call GenerateKey
Label7.Caption = AntiKey
Label1 = "¿Desea Crear el PJ: " & vbNewLine & UserName & " ?"

End Sub

Private Function GenerateKey() As String

AntiKey = RandomNumber(1, 9) & Chr(97 + Rnd() * 862150000 Mod 26) & RandomNumber(1, 9) & Chr(97 + Rnd() * 862150000 Mod 26) & Chr(97 + Rnd() * 862150000 Mod 26) & RandomNumber(1, 9)

End Function




Private Sub Command1_Click()
MP3P.stopMP3
If CheckDatos() Then

#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = IPServidor
    frmMain.Socket1.RemotePort = PortServidor
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 11

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call Login(ValidarLoginMSG(CInt(bRK)))
    End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

