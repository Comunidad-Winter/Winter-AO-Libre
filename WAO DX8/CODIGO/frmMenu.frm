VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Menú"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   495
      Left            =   480
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   720
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   840
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   345
      Index           =   1
      Left            =   480
      MouseIcon       =   "frmMenu.frx":9B27
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   720
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
            Unload Me
            Call frmOpciones.Show(vbModeless, frmMain) 'Stand
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            Call frmEstadisticas.Show(vbModeless, frmMain) 'Stand
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub


Private Sub Image2_Click()
Call frmcanjes.Show(vbModeless, frmMain) 'Stand
End Sub

Private Sub Image3_Click()
Unload Me
End Sub
