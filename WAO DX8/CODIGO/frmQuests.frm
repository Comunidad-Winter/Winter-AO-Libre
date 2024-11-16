VERSION 5.00
Begin VB.Form frmQuests 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misiones"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
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
   ScaleHeight     =   3675
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbandonar 
      Caption         =   "&Abandonar misión"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   3300
      Width           =   2535
   End
   Begin VB.ListBox lstQuests 
      Height          =   3180
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblCriaturas 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   7
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "Criaturas matadas:"
      Height          =   255
      Index           =   2
      Left            =   2700
      TabIndex        =   6
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblDescripcion 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2700
      TabIndex        =   5
      Top             =   1500
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "Descripción de la misión:"
      Height          =   255
      Index           =   1
      Left            =   2700
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "Nombre de la misión:"
      Height          =   255
      Index           =   0
      Left            =   2700
      TabIndex        =   2
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Amra
Option Explicit

Private Sub cmdAbandonar_Click()
    If lstQuests.listIndex < 1 Or lstQuests.List(lstQuests.listIndex) = "NADA" Then Exit Sub
    
    If MsgBox("¿Estás seguro que deseas abandonar la misión " & Chr(34) & lstQuests.List(lstQuests.listIndex) & Chr(34) & "?", vbCritical + vbYesNo, "Argentum Online") = vbYes Then
        Call SendData("QA" & lstQuests.listIndex + 1)
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If lstQuests.List(0) <> "NADA" Then
        Call SendData("QIR1")
    End If
End Sub

Private Sub lstQuests_Click()
    If lstQuests.listIndex < 1 Then Exit Sub
    
    If lstQuests.List(lstQuests.listIndex) = "NADA" Then
        lblCriaturas.Caption = "-"
        lblDescripcion.Caption = "-"
        lblNombre.Caption = "-"
    Else
        Call SendData("QIR" & lstQuests.listIndex + 1)
    End If
End Sub
'/Amra
