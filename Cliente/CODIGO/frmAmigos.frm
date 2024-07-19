VERSION 5.00
Begin VB.Form frmAmigos 
   BorderStyle     =   0  'None
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAmigos.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameAgregarAmigo 
      Caption         =   "Agregar amigo"
      Height          =   1335
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese el nick del amigo:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox Marco 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4013
      Left            =   607
      ScaleHeight     =   4035.17
      ScaleMode       =   0  'User
      ScaleWidth      =   3085.098
      TabIndex        =   0
      Top             =   2400
      Width           =   3080
      Begin VB.OptionButton Amigo1 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton Amigo10 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   3375
      End
      Begin VB.OptionButton Amigo9 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   3375
      End
      Begin VB.OptionButton Amigo8 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   3375
      End
      Begin VB.OptionButton Amigo7 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   3375
      End
      Begin VB.OptionButton Amigo6 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   3375
      End
      Begin VB.OptionButton Amigo5 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   3375
      End
      Begin VB.OptionButton Amigo4 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton Amigo3 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton Amigo2 
         BackColor       =   &H00000000&
         Caption         =   "No hay amigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Label BotonCerrar 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Image BotonMP 
      Height          =   735
      Left            =   4800
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Image BotonBorrar 
      Height          =   615
      Left            =   4800
      Top             =   2880
      Width           =   2445
   End
   Begin VB.Image BotonAgregar 
      Height          =   615
      Left            =   4920
      Top             =   2040
      Width           =   2445
   End
   Begin VB.Label Mensajes 
      BackStyle       =   0  'Transparent
      Caption         =   "No hay acciones realizadas."
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4920
      TabIndex        =   15
      Top             =   5640
      Width           =   2295
   End
End
Attribute VB_Name = "frmAmigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BotonAgregar_Click()
FrameAgregarAmigo.Visible = True
End Sub

Private Sub BotonBorrar_Click()
If Amigo1.value = True Then
    If Amigo1.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo1.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo1", "No hay amigo")
            Amigo1.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo1.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo1.Caption
            End If
        End If
    End If
ElseIf Amigo2.value = True Then
    If Amigo2.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo2.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo2", "No hay amigo")
            Amigo2.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo2.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo2.Caption
            End If
        End If
    End If
ElseIf Amigo3.value = True Then
    If Amigo3.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo3.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo3", "No hay amigo")
            Amigo3.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo3.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo3.Caption
            End If
        End If
    End If
ElseIf Amigo4.value = True Then
    If Amigo4.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo4.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo4", "No hay amigo")
            Amigo4.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo4.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo4.Caption
            End If
        End If
    End If
ElseIf Amigo5.value = True Then
    If Amigo5.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo5.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo5", "No hay amigo")
            Amigo5.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo5.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo5.Caption
            End If
        End If
    End If
ElseIf Amigo6.value = True Then
    If Amigo6.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo6.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo6", "No hay amigo") 'Label de acciones
            Amigo6.Caption = "No hay amigo"
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo6.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo6.Caption
            End If
        End If
    End If
ElseIf Amigo7.value = True Then
    If Amigo7.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo7.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo7", "No hay amigo")
            Amigo7.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo7.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo7.Caption
            End If
        End If
    End If
ElseIf Amigo8.value = True Then
    If Amigo8.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo8.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo8", "No hay amigo")
            Amigo8.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo8.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo8.Caption
            End If
        End If
    End If
ElseIf Amigo9.value = True Then
    If Amigo9.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo9.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo9", "No hay amigo")
            Amigo9.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo9.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo9.Caption
            End If
        End If
    End If
ElseIf Amigo10.value = True Then
    If Amigo10.Caption = "No hay amigo" Then
        MsgBox "Imposible borrar a un amigo que ni siquiera esta en tu lista.", vbOKOnly, "Winter-AO"
    Else
        If MsgBox("¿Desea borrar al amigo " & Amigo10.Caption & "?", vbYesNo, "Winter-AO") = vbYes Then
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo10", "No hay amigo")
            Amigo10.Caption = "No hay amigo"
            'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Borrado el amigo " & Amigo10.Caption
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Borrado el amigo " & Amigo10.Caption
            End If
        End If
    End If
End If
End Sub

Private Sub BotonCerrar_Click()
Unload Me
End Sub



Private Sub Command1_Click()
FrameAgregarAmigo.Visible = False
Text1.Text = ""
End Sub

Private Sub Command2_Click()
If Amigo1.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo1.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo1.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo1", Amigo1.Caption)
    ElseIf Text1.Text = Amigo1.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo1.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo1.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo1", Amigo1.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo2.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo2.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo2.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo2", Amigo2.Caption)
    ElseIf Text1.Text = Amigo2.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo2.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo2.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo2", Amigo2.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo3.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo3.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo3.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo3", Amigo3.Caption)
    ElseIf Text1.Text = Amigo3.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo3.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo3.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo3", Amigo3.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo4.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo4.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo4.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo4", Amigo4.Caption)
    ElseIf Text1.Text = Amigo4.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo4.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo4.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo4", Amigo4.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo5.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo5.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo5.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo5", Amigo5.Caption)
    ElseIf Text1.Text = Amigo5.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo5.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo5.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo5", Amigo5.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo6.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo6.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo6.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo6", Amigo6.Caption)
    ElseIf Text1.Text = Amigo6.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo6.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo6.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo6", Amigo6.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo7.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo7.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo7.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo7", Amigo7.Caption)
    ElseIf Text1.Text = Amigo7.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo7.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo7.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo7", Amigo7.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo8.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo8.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo8.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo8", Amigo8.Caption)
    ElseIf Text1.Text = Amigo8.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo8.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo8.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo8", Amigo8.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo9.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo9.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo9.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo9", Amigo9.Caption)
    ElseIf Text1.Text = Amigo9.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo9.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo9.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo9", Amigo9.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
ElseIf Amigo10.value = True Then
If MsgBox("¿Esta seguro que desea agregar a " & Text1.Text & " como nuevo amigo?", vbYesNo, "Winter-AO") = vbYes Then
    If Amigo10.Caption = "No hay amigo" Then
        'Label de acciones
            If Mensajes.Caption = "No hay acciones realizadas." Then
                Mensajes.ForeColor = vbGreen
                Mensajes.Caption = "Agregado " & Text1.Text & " como nuevo amigo."
                    Else
                Mensajes.Caption = Mensajes.Caption & vbNewLine & "Agregado " & Text1.Text & " como nuevo amigo."
            End If
        'Grabamos al amigo
            Amigo10.Caption = Text1.Text
            Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo10", Amigo10.Caption)
    ElseIf Text1.Text = Amigo10.Caption Then
        MsgBox "Imposible reemplazar el mismo amigo.", vbOKOnly, "Winter-AO"
        Text1.Text = ""
    ElseIf Text1.Text = "" Then
        MsgBox "Nick inválido.", vbOKOnly, "Winter-AO"
    Else
        'Sobreescribir amigo
            If MsgBox("¿Desea sobreescribir al amigo " & Amigo10.Caption & " por " & Text1.Text & "?", vbYesNo, "Sobreescribir amigo | Winter-AO") = vbYes Then
                Amigo10.Caption = Text1.Text
                Call WriteVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo10", Amigo10.Caption)
                    Else
                FrameAgregarAmigo.Visible = False
                Text1.Text = ""
            End If
    End If
End If
FrameAgregarAmigo.Visible = False
End If
End Sub

Private Sub Form_Load()
'Cargamos los amigos
Amigo1.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo1")
Amigo2.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo2")
Amigo3.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo3")
Amigo4.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo4")
Amigo5.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo5")
Amigo6.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo6")
Amigo7.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo7")
Amigo8.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo8")
Amigo9.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo9")
Amigo10.Caption = GetVar(IniPath & "Amigos.ini", "AMIGOS", "Amigo10")
End Sub
