VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   5340
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Configurar Teclas"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Radio WAO"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   255
      Index           =   0
      Left            =   2520
      MouseIcon       =   "frmOpciones.frx":7746C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonidos Activados"
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmOpciones.frx":775BE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Creditos"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Rendimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   4230
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Lo que habla queda en consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1200
         Value           =   2  'Grayed
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Efecto de las FX"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Value           =   2  'Grayed
         Width           =   2415
      End
      Begin VB.CheckBox ActivarNoche 
         BackColor       =   &H00000000&
         Caption         =   "Activar / Desactivar Efecto Noche "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton Minimapa 
         Caption         =   "Desactivar Minimapa"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Ver Nombre de los mapas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Ver Fps"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00000000&
         Caption         =   "Ver Nombre de los usuarios"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Si no tienes un buen pc es recomendable desactivar las siguientes opciones."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4230
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "5"
         Top             =   240
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   225
      Left            =   120
      MouseIcon       =   "frmOpciones.frx":77710
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private Sub ActivarNoche_Click()
EfectosDiaY = Not EfectosDiaY
End Sub

Private Sub Check1_Click()

If MapNameY = True Then
    MapNameY = False
    frmMain.MapName.Visible = False
Else
    MapNameY = True
    frmMain.MapName.Visible = True
End If
    
End Sub

Private Sub Check2_Click()

If FpsY = True Then
    FpsY = False
    frmMain.fpps.Caption = "Desactivado"
Else
    FpsY = True
End If
    
End Sub

Private Sub Check3_Click()
'
'If EfectosAlphaY Then
'    EfectosAlphaY = False
'    frmMain.EfectosAlpha.Enabled = False
'    AlphaX = 150
'Else
'    EfectosAlphaY = True
'    frmMain.EfectosAlpha.Enabled = True
'End If
    
End Sub

Private Sub Check4_Click()
    Call SendData("/SEF")
    'ConsolaY = Not ConsolaY
End Sub

Private Sub chkop_Click()
Nombres = Not Nombres
End Sub

Private Sub Command1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        If Musica Then
            Musica = False
            Command1(0).Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            Command1(0).Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Command1(1).Caption = "Sonidos Desactivados"
            Call Audio.StopWave
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        Else
            Sound = True
            Command1(1).Caption = "Sonidos Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Me.Visible = False
Call GuardarOpciones
End Sub


Private Sub Command4_Click()
Call FrmCredits.Show(vbModeless, frmMain) 'Stand
End Sub


Private Sub Command5_Click()
Call frmCustomKeys.Show(vbModeless, frmMain) 'Stand
Unload Me
End Sub

Private Sub Command6_Click()
Call frmRadio.Show(vbModeless, frmMain) 'Stand
Unload Me
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Form_load()
    If Musica Then
        Command1(0).Caption = "Musica Activada"
    Else
        Command1(0).Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        Command1(1).Caption = "Sonidos Activados"
    Else
        Command1(1).Caption = "Sonidos Desactivados"
    End If

End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub
