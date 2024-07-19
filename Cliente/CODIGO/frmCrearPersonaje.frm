VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox HeadView 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1095
      ScaleHeight     =   345
      ScaleWidth      =   495
      TabIndex        =   44
      Top             =   5700
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10800
      Top             =   600
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":2D59B
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":2D5D2
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2700
      Width           =   2820
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":2D66C
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":2D676
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2250
      Width           =   2820
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":2D689
      Left            =   5925
      List            =   "frmCrearPersonaje.frx":2D69C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1800
      Width           =   2820
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":2D6C9
      Left            =   9120
      List            =   "frmCrearPersonaje.frx":2D6D3
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3480
      Width           =   2565
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   3570
      TabIndex        =   0
      Top             =   1275
      Width           =   4815
   End
   Begin VB.Image MasHead 
      Height          =   495
      Left            =   1800
      Top             =   5640
      Width           =   375
   End
   Begin VB.Image MenosHead 
      Height          =   495
      Left            =   600
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Se va a crear:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9240
      TabIndex        =   43
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label lblinfo2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   42
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  'Transparent
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
      Height          =   855
      Left            =   3360
      TabIndex        =   41
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5205
      TabIndex        =   40
      Top             =   4425
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   4920
      MouseIcon       =   "frmCrearPersonaje.frx":2D6E9
      MousePointer    =   99  'Custom
      Top             =   2505
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   4920
      MouseIcon       =   "frmCrearPersonaje.frx":2E3B3
      MousePointer    =   99  'Custom
      Top             =   2850
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   4935
      MouseIcon       =   "frmCrearPersonaje.frx":2F07D
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   3
      Left            =   4920
      MouseIcon       =   "frmCrearPersonaje.frx":2FD47
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   4
      Left            =   4920
      MouseIcon       =   "frmCrearPersonaje.frx":30A11
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   270
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   0
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":316DB
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   1
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":323A5
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   2
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":3306F
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   3
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":33D39
      MousePointer    =   99  'Custom
      Top             =   3690
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   180
      Index           =   4
      Left            =   5640
      MouseIcon       =   "frmCrearPersonaje.frx":34A03
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label modAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   39
      Top             =   2850
      Width           =   315
   End
   Begin VB.Label modInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   3165
      Width           =   330
   End
   Begin VB.Label modCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5460
      TabIndex        =   37
      Top             =   3570
      Width           =   435
   End
   Begin VB.Label modConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5460
      TabIndex        =   36
      Top             =   3900
      Width           =   435
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   35
      Top             =   2580
      Width           =   330
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   7995
      TabIndex        =   34
      Top             =   7710
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   43
      Left            =   7800
      MouseIcon       =   "frmCrearPersonaje.frx":356CD
      MousePointer    =   99  'Custom
      Top             =   7725
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":3581F
      MousePointer    =   99  'Custom
      Top             =   7755
      Width           =   165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   4020
      TabIndex        =   33
      Top             =   4260
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image19 
      Height          =   3120
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   2475
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   32
      Top             =   8535
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":35971
      MousePointer    =   99  'Custom
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":35AC3
      MousePointer    =   99  'Custom
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   7785
      MouseIcon       =   "frmCrearPersonaje.frx":35C15
      MousePointer    =   99  'Custom
      Top             =   3885
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":35D67
      MousePointer    =   99  'Custom
      Top             =   4110
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":35EB9
      MousePointer    =   99  'Custom
      Top             =   4335
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":3600B
      MousePointer    =   99  'Custom
      Top             =   4545
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":3615D
      MousePointer    =   99  'Custom
      Top             =   4785
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":362AF
      MousePointer    =   99  'Custom
      Top             =   4965
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":36401
      MousePointer    =   99  'Custom
      Top             =   5175
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":36553
      MousePointer    =   99  'Custom
      Top             =   5385
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":366A5
      MousePointer    =   99  'Custom
      Top             =   5610
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":367F7
      MousePointer    =   99  'Custom
      Top             =   5820
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":36949
      MousePointer    =   99  'Custom
      Top             =   6015
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":36A9B
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":36BED
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":36D3F
      MousePointer    =   99  'Custom
      Top             =   3495
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":36E91
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":36FE3
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37135
      MousePointer    =   99  'Custom
      Top             =   4155
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37287
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   8355
      MouseIcon       =   "frmCrearPersonaje.frx":373D9
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3752B
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3767D
      MousePointer    =   99  'Custom
      Top             =   4995
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":377CF
      MousePointer    =   99  'Custom
      Top             =   5220
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37921
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37A73
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37BC5
      MousePointer    =   99  'Custom
      Top             =   5850
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37D17
      MousePointer    =   99  'Custom
      Top             =   6075
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":37E69
      MousePointer    =   99  'Custom
      Top             =   6285
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":37FBB
      MousePointer    =   99  'Custom
      Top             =   6270
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":3810D
      MousePointer    =   99  'Custom
      Top             =   6495
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":3825F
      MousePointer    =   99  'Custom
      Top             =   6465
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":383B1
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":38503
      MousePointer    =   99  'Custom
      Top             =   6690
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":38655
      MousePointer    =   99  'Custom
      Top             =   6945
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":387A7
      MousePointer    =   99  'Custom
      Top             =   6915
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":388F9
      MousePointer    =   99  'Custom
      Top             =   7170
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   7755
      MouseIcon       =   "frmCrearPersonaje.frx":38A4B
      MousePointer    =   99  'Custom
      Top             =   7125
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":38B9D
      MousePointer    =   99  'Custom
      Top             =   7395
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   7770
      MouseIcon       =   "frmCrearPersonaje.frx":38CEF
      MousePointer    =   99  'Custom
      Top             =   7335
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   8370
      MouseIcon       =   "frmCrearPersonaje.frx":38E41
      MousePointer    =   99  'Custom
      Top             =   7590
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7815
      MouseIcon       =   "frmCrearPersonaje.frx":38F93
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   495
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmCrearPersonaje.frx":390E5
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1125
   End
   Begin VB.Image boton 
      Height          =   570
      Index           =   0
      Left            =   9840
      MouseIcon       =   "frmCrearPersonaje.frx":39237
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1560
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7995
      TabIndex        =   28
      Top             =   7515
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7995
      TabIndex        =   27
      Top             =   7306
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7995
      TabIndex        =   26
      Top             =   7092
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7995
      TabIndex        =   25
      Top             =   6878
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7995
      TabIndex        =   24
      Top             =   6664
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7995
      TabIndex        =   23
      Top             =   6450
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7995
      TabIndex        =   22
      Top             =   6236
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7995
      TabIndex        =   21
      Top             =   6022
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7995
      TabIndex        =   20
      Top             =   5808
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7995
      TabIndex        =   19
      Top             =   5594
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7995
      TabIndex        =   18
      Top             =   5380
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7995
      TabIndex        =   17
      Top             =   5166
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7995
      TabIndex        =   16
      Top             =   4952
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7995
      TabIndex        =   15
      Top             =   4738
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7995
      TabIndex        =   14
      Top             =   4524
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7995
      TabIndex        =   13
      Top             =   4310
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7995
      TabIndex        =   12
      Top             =   4096
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7995
      TabIndex        =   11
      Top             =   3882
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7995
      TabIndex        =   10
      Top             =   3668
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7995
      TabIndex        =   9
      Top             =   3240
      Width           =   270
   End
   Begin VB.Label Skill 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7995
      TabIndex        =   8
      Top             =   3450
      Width           =   270
   End
   Begin VB.Image imgHogar 
      Height          =   2445
      Left            =   9120
      Picture         =   "frmCrearPersonaje.frx":39389
      Top             =   3960
      Width           =   1845
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   3570
      Width           =   225
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   3
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   2
      Top             =   2910
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5280
      TabIndex        =   1
      Top             =   2580
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PuedeDados As Boolean

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If


If Label1.Caption <> 0 Then
    MsgBox "Los atributos del personaje son invalidos."
    Exit Function
End If

If Len(txtNombre.Text) < 3 Then
    MsgBox "El nombre debe de tener mas de 3 caracteres!!"
    Exit Function
End If

If Len(txtNombre.Text) >= 12 Then
    MsgBox "El nombre debe de tener menos de 12 caracteres!!"
    Exit Function
End If

If Right$(UserName, 1) = " " Then
    UserName = RTrim$(UserName)
        MsgBox "Nombre invalido."
    Exit Function
End If

CheckData = True


End Function
Private Sub boton_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
      Call SendData("TIRDAD" & lbFuerza.Caption & "," & lbAgilidad.Caption & "," & lbInteligencia.Caption & "," & lbCarisma.Caption & "," & lbConstitucion.Caption)
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        

        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido."
                Exit Sub
        End If
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmPasswdSinPadrinos.Show vbModal, Me
        End If
        
    Case 1
    If MsgBox("¿Esta seguro/a que desea salir de la creacion de Personaje?", vbYesNo, "Winter AO 2.0") = vbYes Then
    
    Windows_Temp_Dir = General_Get_Temp_Dir
 Set MP3P = New clsMP3Player
    Call Extract_File2(MP3, App.Path & "\ARCHIVOS\", "2.mp3", Windows_Temp_Dir, False)
    MP3P.mp3file = Windows_Temp_Dir & "2.mp3"
    MP3P.stopMP3
    MP3P.playMP3
    MP3P.Volume = 1000
    End If
        
        Form2.Show
        Me.Visible = False
        End Select
End Sub



Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()

lbFuerza.Caption = "10"
lbInteligencia.Caption = "10"
lbAgilidad.Caption = "10"
lbCarisma.Caption = "10"
lbConstitucion.Caption = "10"

End Sub
Private Sub Command1_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
'Form2.Visible = False
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\Interfaces\CP-Interface.jpg")
imgHogar.Picture = LoadPicture(App.Path & "\Interfaces\CP-Ullathorpe.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

Image19.Picture = LoadPicture(App.Path & "\Interfaces\" & lstProfesion.Text & ".jpg")
Call TirarDados
End Sub



Private Sub Image1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Label1.Caption > 0 Then

    Select Case index
    Case 0
    
        If lbFuerza.Caption < 18 Then
            lbFuerza.Caption = lbFuerza.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 1
    
        If lbAgilidad.Caption < 18 Then
            lbAgilidad.Caption = lbAgilidad.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 2
    
        If lbInteligencia.Caption < 18 Then
            lbInteligencia.Caption = lbInteligencia.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 3
        
        If lbCarisma.Caption < 18 Then
            lbCarisma.Caption = lbCarisma.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    Case 4
        
        If lbConstitucion.Caption < 18 Then
            lbConstitucion.Caption = lbConstitucion.Caption + 1
            Label1.Caption = Label1.Caption - 1
        End If
        
    End Select
    
End If

End Sub

Private Sub Image3_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

If Label1.Caption >= 0 Then

    Select Case index
    Case 0
    
        If lbFuerza.Caption > 10 Then
            lbFuerza.Caption = lbFuerza.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 1
    
        If lbAgilidad.Caption > 10 Then
            lbAgilidad.Caption = lbAgilidad.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 2
    
        If lbInteligencia.Caption > 10 Then
            lbInteligencia.Caption = lbInteligencia.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 3
        
        If lbCarisma.Caption > 10 Then
            lbCarisma.Caption = lbCarisma.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    Case 4
        
        If lbConstitucion.Caption > 10 Then
            lbConstitucion.Caption = lbConstitucion.Caption - 1
            Label1.Caption = Label1.Caption + 1
        End If
        
    End Select
    
End If

End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
lblinfo.Caption = "Clase elejida " & lstProfesion.Text
Image19.Picture = LoadPicture(App.Path & "\Interfaces\" & lstProfesion.Text & ".jpg")
End Sub
Private Sub Timer1_Timer()
lblinfo2.Caption = lstProfesion.Text & "," & lstRaza.Text & "," & lstGenero.Text
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub
Private Sub lstGenero_Click()
Call DameOpciones
End Sub
 
Private Sub lstRaza_Click()
Call DameOpciones
End Sub
Private Sub MenosHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual - 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
Call DrawGrhtoHdc2(HeadView.hdc, HeadData(Actual).Head(3).GrhIndex, 8, 5)
End Sub
Private Sub MasHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual + 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
Call DrawGrhtoHdc2(HeadView.hdc, HeadData(Actual).Head(3).GrhIndex, 8, 5)
End Sub
Private Sub lstProfesion_GotFocus()
lblinfo.Caption = "Elija Su Clase"
End Sub
Private Sub lstHogar_GotFocus()
lblinfo.Caption = "Elija Su Hogar"
End Sub
Private Sub lstRaza_GotFocus()
lblinfo.Caption = "Elija Su Raza"
End Sub
 Private Sub lstGenero_GotFocus()
lblinfo.Caption = "Elija Su Genero"
End Sub
Private Sub Cabeza_GotFocus()
lblinfo.Caption = "Elija Su Cabeza"
End Sub
Private Sub lstHogar_Click()
lblinfo.Caption = "Hogar elejido " & lstHogar.Text
End Sub
