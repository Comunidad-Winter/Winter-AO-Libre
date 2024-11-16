VERSION 5.00
Begin VB.Form FrmError 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "FrmError.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmError.frx":13DCE
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
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1200
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "FrmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
    'Valores máximos y mínimos para el ScrollBar
    HScroll1.max = 255
    HScroll1.min = 60
  
    ' Le establecemos un valor por defecto _
    a la barra apenas carga el form
  
    HScroll1.value = 200
End Sub
  
Private Sub HScroll1_Change()
  
    'Llamamos a la función pasándole el handle del form _
    y el valor de la transparencia, que es el de la barra
  
    Call Aplicar_Transparencia(Me.hWnd, CByte(HScroll1.value))
  
End Sub

Private Sub Image1_Click()
Me.Visible = False
End Sub

