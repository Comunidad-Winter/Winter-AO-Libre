VERSION 5.00
Begin VB.Form FrmMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mapa del Mundo de Winter-AO"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   6360
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   11175
      Left            =   -240
      Picture         =   "FrmMap.frx":0000
      ScaleHeight     =   11115
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   -360
      Width           =   8415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transparencia del Mapa"
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
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   6120
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transparencia del Mapa"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "FrmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
    'Valores máximos y mínimos para el ScrollBar
    HScroll1.max = 255
    HScroll1.min = 50
  
    ' Le establecemos un valor por defecto _
    a la barra apenas carga el form
  
    HScroll1.value = 200
End Sub
  
Private Sub HScroll1_Change()
  
    'Llamamos a la función pasándole el handle del form _
    y el valor de la transparencia, que es el de la barra
  
    Call Aplicar_Transparencia(Me.hWnd, CByte(HScroll1.value))
  
End Sub
