VERSION 5.00
Begin VB.Form FrmMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mapa del Mundo de Winter-AO"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   7560
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   0
      Picture         =   "FrmMap.frx":0000
      ScaleHeight     =   7995
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transparencia del Mapa"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   7320
         Width           =   2175
      End
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
