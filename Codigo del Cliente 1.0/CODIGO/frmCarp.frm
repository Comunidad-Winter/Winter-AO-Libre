VERSION 5.00
Begin VB.Form frmCarp 
   Caption         =   "Carpintero"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "1"
      Top             =   2760
      Width           =   495
   End
   Begin VB.ListBox lstArmas 
      Height          =   2205
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Construir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      MouseIcon       =   "frmCarp.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2670
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
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
      Left            =   255
      MouseIcon       =   "frmCarp.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2670
      Width           =   1710
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command3_Click()
On Error Resume Next
If Int(Val(Cantidad)) < 1 Or Int(Val(Cantidad)) > 1000 Then
    MsgBox "La cantidad es invalida.", vbCritical
    Exit Sub
End If
Call SendData("CNC" & ObjCarpintero(lstArmas.listIndex) & "," & Cantidad.Text)

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

