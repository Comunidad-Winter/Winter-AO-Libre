VERSION 5.00
Begin VB.Form frmReproductor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Winter AO - Reproductor"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
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
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tema:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmReproductor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Text1.Text = Val(Text1) + 1
        Case 1
            If Val(Text1) = 0 Then Exit Sub
            Text1.Text = Val(Text1) - 1
    End Select

End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
         If Not FileExist(App.Path & "\Mi MP3\" & Val(Text1) & ".mp3", vbNormal) Then
                Exit Sub
            End If
            MP3P.stopMP3
            MP3P.mp3file = App.Path + "\Mi MP3\" & Val(Text1) & ".mp3"
            MP3P.playMP3
        Case 1
           MP3P.stopMP3
            
End Select
End Sub

