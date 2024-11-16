VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione la criatura"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
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
   ScaleHeight     =   3705
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2460
      MouseIcon       =   "frmEntrenador.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3000
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Luchar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   780
      MouseIcon       =   "frmEntrenador.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1710
   End
   Begin VB.ListBox lstCriaturas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   780
      TabIndex        =   1
      Top             =   540
      Width           =   2490
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "¿Con qué criatura deseas combatir?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   0
      Top             =   105
      Width           =   3525
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'    Le puse el iconito de la manito a los botones ^_^,
'   le puse borde a la ventana y le cambié la letra a
'   una más linda :)
Option Explicit

Private Sub Command1_Click()
Call SendData("ENTR" & lstCriaturas.listIndex + 1)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

