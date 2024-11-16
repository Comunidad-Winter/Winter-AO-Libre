VERSION 5.00
Begin VB.Form FrmHogar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   5925
   ClientTop       =   4290
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   Picture         =   "FrmHogar.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image2 
      Height          =   495
      Left            =   720
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   840
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "FrmHogar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FrmLanzador.Picture = LoadPicture(App.Path & _
    "\Interfaces\Muerte.jpg")
End Sub

Private Sub Image1_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Image2_Click()
Call SendData("/HOGAR")
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub
