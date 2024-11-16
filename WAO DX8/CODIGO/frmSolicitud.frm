VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   Picture         =   "frmSolicitud.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmSolicitud.frx":1A9AB
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   3360
      MouseIcon       =   "frmSolicitud.frx":1AAFD
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   240
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSolicitud.frx":1AC4F
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CName As String

Private Sub Command1_Click()
Dim f$

f$ = "SOLICITUD" & CName
f$ = f$ & "," & Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º")

Call SendData(f$)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)

CName = GuildName

End Sub

