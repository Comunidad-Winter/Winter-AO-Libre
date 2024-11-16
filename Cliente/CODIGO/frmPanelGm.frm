VERSION 5.00
Begin VB.Form frmPanelGm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Comandos que afectan a los usuarios"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   4215
      Begin VB.CommandButton cmdAccion 
         Caption         =   "DONDE"
         Height          =   195
         Index           =   6
         Left            =   1560
         TabIndex        =   38
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "HORA"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "N.ENE."
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "INFO"
         Height          =   195
         Index           =   8
         Left            =   3120
         TabIndex        =   35
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Ver oro en banco"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Usuarios que enviaron GM"
         Height          =   195
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Desactivar Denuncias"
         Height          =   195
         Left            =   1680
         TabIndex        =   32
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Activar Denuncias"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Hacer Torneo"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Global"
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
         Height          =   495
         Left            =   1080
         TabIndex        =   27
         Top             =   1800
         Width           =   2895
         Begin VB.CommandButton Command5 
            Caption         =   "Desactivar Global"
            Height          =   195
            Left            =   1320
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Activar Global"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "ECHAR"
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "BAN"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "INV"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "SKILLS"
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "CARCEL"
         Height          =   195
         Index           =   11
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "UNBAN"
         Height          =   195
         Index           =   12
         Left            =   2400
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "NICK 2 IP"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "IP 2 NICK"
         Height          =   195
         Index           =   14
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Penas"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Ban X ip"
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Boveda"
         Height          =   195
         Index           =   17
         Left            =   2040
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "Show SOS"
         Height          =   195
         Index           =   18
         Left            =   3000
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Teletransporte"
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
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   4215
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Ciudades"
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
         Height          =   1335
         Left            =   0
         TabIndex        =   41
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton Command23 
            Caption         =   "Ir a Rinkel"
            Height          =   255
            Left            =   960
            TabIndex        =   51
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Ir a Arghal"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Ir a Ulla"
            Height          =   255
            Left            =   3360
            TabIndex        =   49
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Ir a Banderbill"
            Height          =   255
            Left            =   2280
            TabIndex        =   48
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Ir a Shakoud"
            Height          =   255
            Left            =   1200
            TabIndex        =   47
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Ir a Winderbill"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Ir a Feinur"
            Height          =   255
            Left            =   3120
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Ir a Winley"
            Height          =   255
            Left            =   2160
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Ir a Windell"
            Height          =   255
            Left            =   1200
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ir a Ramx"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "SUM"
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   40
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "IRA"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   39
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ir a Arenas de Torneo"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Clima"
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
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "Hacer de Noche"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Hacer de Dia"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Hacer de Tarde"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Lluvia"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Hacer de Mañana"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Hacer Niebla"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "Usuario"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar comentario"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Width           =   4035
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   120
      Y1              =   1020
      Y2              =   1860
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   4440
      Y1              =   540
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2280
      Y1              =   1200
      Y2              =   1620
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2280
      Y1              =   1620
      Y2              =   1620
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccion_Click(index As Integer)
Dim ok As Boolean, Tmp As String, Tmp2 As String
Dim Nick As String

Nick = cboListaUsus.Text

Select Case index
Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    Tmp = InputBox("Motivo ?", "")
    If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/BAN " & Tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    Tmp = InputBox("Comentario ?", "")
    Call SendData("/REM " & Tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    Tmp = InputBox("Mapa ?", "")
    Call SendData("/NENE " & Trim(Tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & cboListaUsus.Text)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    Tmp = InputBox("Minutos ? (hasta 30)", "")
    Tmp2 = InputBox("Razon ?", "")
    If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/CARCEL " & Nick & "@" & Tmp2 & "@" & Tmp)
    End If
Case 12 '/unban nick
    If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/ip2nick ip
    Call SendData("/IP2NICK " & Nick)
Case 15 '/penas
    Call SendData("/PENAS " & cboListaUsus.Text)
Case 16 'Ban X ip
    Tmp = InputBox("Ingrese el motivo del ban", "Ban X IP")
    If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
        Nick = Replace(Nick, " ", "+")
        Call SendData("/BANIP " & Nick & Tmp)
    End If
Case 17 ' MUESTA BOBEDA
    Call SendData("/BOV " & Nick)
Case 18 ' Sos
    Call SendData("/SHOW SOS")
Case 19 ' Balance
    Call SendData("/BAL " & cboListaUsus.Text)
End Select


End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 118 70 66")
End Sub

Private Sub Command10_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 94 50 50")
End Sub

Private Sub Command11_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TARDE")
End Sub

Private Sub Command12_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/HACERTORNEO")
End Sub

Private Sub Command13_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/SHOW SOS")
End Sub


Private Sub Command15_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 268 50 50")
End Sub

Private Sub Command18_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/MAÑANA")
End Sub

Private Sub Command2_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 1 50 50")
End Sub

Private Sub Command20_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 170 50 50")
End Sub

Private Sub Command21_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 233 50 50")
End Sub

Private Sub Command22_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 183 50 50")
End Sub

Private Sub Command23_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 246 50 50")
End Sub

Private Sub Command3_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 43 50 50")
End Sub

Private Sub Command4_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/GLOB AC")
End Sub

Private Sub Command5_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/GLOB DES")
End Sub

Private Sub Command6_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/NOCHE")
End Sub

Private Sub Command7_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/MAÑANA")
End Sub

Private Sub Command8_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 70 50 50")
End Sub

Private Sub Command9_Click()
Call Audio.PlayWave(SND_CLICK)
Call SendData("/TELEP YO 112 50 50")
End Sub

Private Sub Form_Load()
    Call cmdActualiza_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

