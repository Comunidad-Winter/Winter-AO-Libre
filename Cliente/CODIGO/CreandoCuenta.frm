VERSION 5.00
Begin VB.Form CreandoCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   4605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Text            =   "CreandoCuenta.frx":0000
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Mail 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox RePass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Nombre 
      Height          =   285
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Verificación"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esperando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Password"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuenta"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "CreandoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Pass <> RePass Then
    MsgBox "Lass passwords que tipeo no coinciden", , "Winter AO"
    Exit Sub
End If

If Not CheckMailString(Mail) Then
    MsgBox "Direccion de mail invalida."
    Exit Sub
End If

If Nombre = "" Or Pass = "" Or RePass = "" Or Mail = "" Then
    MsgBox "Completa todo!"
    Exit Sub
End If

Call SendData("NCUENT" & Nombre & "," & Pass & "," & Mail)

DoEvents

Cuenta = Nombre.Text

'EstadoLogin = Dados
'Load frmCrearPersonaje
'frmCrearPersonaje.Show vbModal

Unload Me

'Debug.Print "frm show"

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Mail_Click()
If Idioma = 1 Then
If Not RePass.Text = "" Then
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
Else
    Estado.ForeColor = vbBlack
    Estado.Caption = "Esperando..."
End If
Else
If Not RePass.Text = "" Then
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
Else
    Estado.ForeColor = vbBlack
    Estado.Caption = "Waiting..."
End If
End If
End Sub

Private Sub Nombre_Click()
If Idioma = 1 Then
If Pass.Text = "" Then
    Estado.ForeColor = vbBlack
    Estado.Caption = "Esperando..."
Else
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
Else
If Pass.Text = "" Then
    Estado.ForeColor = vbBlack
    Estado.Caption = "Waiting..."
Else
If Pass.Text = RePass.Text Then
    Estado.ForeColor = vbGreen
    Estado.Caption = "Password Correcto."
        Else
    Estado.ForeColor = vbRed
    Estado.Caption = "Password Incorrecto."
    End If
End If
End If
End Sub

Private Sub Pass_Click()
If Idioma = 1 Then
    Estado.ForeColor = vbBlack
    Estado.Caption = "Esperando..."
Else
    Estado.ForeColor = vbBlack
    Estado.Caption = "Waiting..."
End If
End Sub

Private Sub RePass_Click()
If Idioma = 1 Then
    Estado.ForeColor = vbBlack
    Estado.Caption = "Esperando..."
Else
    Estado.ForeColor = vbBlack
    Estado.Caption = "Waiting..."
End If
End Sub
