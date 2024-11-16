VERSION 5.00
Begin VB.Form frmBorrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar personaje"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   150
      MouseIcon       =   "frmBorrar.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2325
      Width           =   1005
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      MouseIcon       =   "frmBorrar.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtPasswd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.Label lblWhat 
      Caption         =   $"frmBorrar.frx":02A4
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   195
      TabIndex        =   8
      Top             =   585
      Width           =   4290
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   1635
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre del personaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Atención"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mediante esta acción borrarás el personaje almacenado en el servidor y no habrá manera de recuperarlo!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Visible         =   0   'False
      Width           =   4440
   End
End
Attribute VB_Name = "frmBorrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdBorrar_Click()
#If usarrwench = 1 Then

frmMain.Socket1.HostName = CurServerIp
frmMain.Socket1.RemotePort = CurServerPort
Me.MousePointer = 11
frmMain.Socket1.Connect

#Else

Me.MousePointer = 11
If frmMain.Winsock1.State <> sckClosed Then _
    frmMain.Winsock1.Close
frmMain.Winsock1.Connect CurServerIp, CurServerPort

#End If
End Sub

Private Sub Command1_Click()
Me.Visible = False
End Sub

