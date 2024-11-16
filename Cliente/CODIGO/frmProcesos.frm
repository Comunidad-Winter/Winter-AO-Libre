VERSION 5.00
Begin VB.Form frmProcesos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesos"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
