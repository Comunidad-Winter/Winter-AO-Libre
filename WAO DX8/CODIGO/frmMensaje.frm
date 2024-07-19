VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmMensaje.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1200
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label msg 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_load()
    FrmLanzador.Picture = LoadPicture(App.Path & _
    "\Interfaces\Mensaje.jpg")
Unload Me
End Sub

