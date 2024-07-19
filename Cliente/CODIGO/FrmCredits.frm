VERSION 5.00
Begin VB.Form FrmCredits 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Creditos"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      Picture         =   "FrmCredits.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   4560
      Width           =   6855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Santo y todos aquellos que nos han ayudado al desarrollo de Winter-AO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   8
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Agradecimientos Especiales a:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Erwin"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Graficacion y Indexacion:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Erwin Wanzer"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mapeo:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Erwin"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Idea original y Lider del Proyecto:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Creditos"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "FrmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
