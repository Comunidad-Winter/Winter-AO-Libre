VERSION 5.00
Begin VB.Form frmSendOro 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Enviar Oro"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   1215
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmSendOro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If MsgBox("Esta seguro que desea enviar " & Text1(1).Text & " monedas de oro al personaje " & Text1(0).Text & " ?", vbYesNo) = vbYes Then
                Call SendData("/DARORO " & Text1(0).Text & "@" & Text1(1).Text)
            Else
                Unload Me
            End If
    End Select
End Sub
