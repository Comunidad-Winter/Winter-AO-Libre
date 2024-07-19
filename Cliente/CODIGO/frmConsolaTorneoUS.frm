VERSION 5.00
Begin VB.Form frmConsolaTorneoUS 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear torneo"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1185
      Left            =   150
      TabIndex        =   23
      Top             =   150
      Width           =   3165
      Begin VB.TextBox Txt_Cupo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   26
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox Txt_LvlMax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   25
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox Txt_LvlMin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   24
         Top             =   525
         Width           =   1275
      End
      Begin VB.Label Cup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de jugadores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel mínimo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel máximo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Clases válidas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3165
      Left            =   3480
      TabIndex        =   14
      Top             =   150
      Width           =   1395
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Guerrero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   22
         Top             =   315
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Mago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   21
         Top             =   690
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Paladín"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   20
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Clérigo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   19
         Top             =   1380
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Bardo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   18
         Top             =   1740
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Asesino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   2070
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Cazador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   16
         Top             =   2730
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Druida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   15
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Comenzar torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Facción / Alineación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   975
      Left            =   150
      TabIndex        =   8
      Top             =   1410
      Width           =   3165
      Begin VB.CheckBox Check10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Criminal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1680
         TabIndex        =   12
         Top             =   525
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox Check11 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Ciudadano"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1680
         TabIndex        =   11
         Top             =   210
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox Check12 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Armada Caos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   525
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox Check13 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Armada Real"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   105
         TabIndex        =   9
         Top             =   210
         Value           =   1  'Checked
         Width           =   1590
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Summon automático"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   855
      Left            =   150
      TabIndex        =   0
      Top             =   2460
      Width           =   3165
      Begin VB.TextBox TxtMap 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   105
         MaxLength       =   3
         TabIndex        =   4
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TxtX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   690
         MaxLength       =   2
         TabIndex        =   3
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TxtY 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   2
         Top             =   420
         Width           =   435
      End
      Begin VB.CheckBox Check9 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Activado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Left            =   1950
         TabIndex        =   1
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   735
         TabIndex        =   6
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1260
         TabIndex        =   5
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmConsolaTorneous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()

If Not CheckDatos Then Exit Sub
Call SendData("/TOR " & Txt_LvlMin & " " & Txt_LvlMax & " " & Txt_Cupo & " " & Check1.Value & " " & Check2.Value & " " & Check3.Value & " " & Check4.Value & " " & Check5.Value & " " & Check6.Value & " " & Check7.Value & " " & Check8.Value & " " & Check9.Value & " " & TxtMap & " " & TxtX & " " & TxtY & " " & Check10.Value & " " & Check11.Value & " " & Check12.Value & " " & Check13.Value)
Unload Me

End Sub

Function CheckDatos() As Boolean

CheckDatos = True

If Txt_LvlMax = "" Then
    CheckDatos = False
    MsgBox "Falta completa el nivel máximo."
    Exit Function
End If

If Txt_LvlMin = "" Then
    MsgBox "Falta completa el nivel mínimo."
    CheckDatos = False
    Exit Function
End If

If Txt_Cupo = "" Then
    MsgBox "Falta completa el cupo."
    CheckDatos = False
    Exit Function
End If

If Not IsNumeric(Txt_LvlMax) Then
    CheckDatos = False
    MsgBox "Nivel máximo no numérico."
    Exit Function
End If

If Not IsNumeric(Txt_LvlMin) Then
    MsgBox "Nivel mínimo no numérico."
    CheckDatos = False
    Exit Function
End If

If Not IsNumeric(Txt_Cupo) Then
    MsgBox "Cupo no numérico."
    CheckDatos = False
    Exit Function
End If

End Function
