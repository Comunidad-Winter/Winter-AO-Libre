VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00004040&
   BorderStyle     =   0  'None
   Caption         =   "Cuentas"
   ClientHeight    =   9000
   ClientLeft      =   3615
   ClientTop       =   3150
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7260
      Picture         =   "Form2.frx":1BCF1
      ScaleHeight     =   480
      ScaleWidth      =   2055
      TabIndex        =   16
      Top             =   3720
      Width           =   2055
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   7320
      Picture         =   "Form2.frx":2018F
      ScaleHeight     =   30
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   5355
      Width           =   375
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   7320
      Picture         =   "Form2.frx":2071C
      ScaleHeight     =   0
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   5355
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   7440
      Picture         =   "Form2.frx":20CA9
      ScaleHeight     =   1305
      ScaleWidth      =   600
      TabIndex        =   13
      Top             =   3840
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   8
      Left            =   720
      Picture         =   "Form2.frx":23C43
      TabIndex        =   8
      Top             =   7320
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   7
      Left            =   720
      Picture         =   "Form2.frx":35523
      TabIndex        =   7
      Top             =   6825
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   6
      Left            =   720
      Picture         =   "Form2.frx":46E03
      TabIndex        =   6
      Top             =   6300
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   5
      Left            =   690
      Picture         =   "Form2.frx":586E3
      TabIndex        =   5
      Top             =   5700
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   4
      Left            =   690
      Picture         =   "Form2.frx":69FC3
      TabIndex        =   4
      Top             =   5160
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   705
      Picture         =   "Form2.frx":7B8A3
      TabIndex        =   3
      Top             =   4590
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   705
      Picture         =   "Form2.frx":8D183
      TabIndex        =   2
      Top             =   4020
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   7440
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   3720
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   705
      Picture         =   "Form2.frx":9EA63
      TabIndex        =   0
      Top             =   3480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   8760
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Mapa 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   0
      TabIndex        =   11
      Top             =   9000
      Width           =   15
   End
   Begin VB.Label Oro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   10
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Nivel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   5460
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   7560
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   6600
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Top             =   7800
      Width           =   345
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ddsd As DDSURFACEDESC2
Dim CharSurface As DirectDrawSurface7
Dim dck As DDCOLORKEY


Private Sub Form_Load()
'Unload frmConnect
Label1 = "Cargando... " & vbNewLine & " Clickea en un PJ"

    'Iniciamos la surface
    ddsd.lHeight = 152
    ddsd.lWidth = 152
    ddsd.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    Set CharSurface = DirectDraw.CreateSurface(ddsd)

    dck.Low = 0
    dck.High = 0
    CharSurface.SetColorKey DDCKEY_SRCBLT, dck


DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1 = ""
End Sub


Private Sub Image1_Click()
If MsgBox("¿Esta seguro/a que desea salir de la cuenta?", vbYesNo, "Winter AO 2.0") = vbYes Then
    #If UsarWrench = 1 Then
            If frmMain.Socket1.Connected Then
                frmMain.Socket1.Disconnect
                frmMain.Socket1.Cleanup
            End If
    #Else
            If frmMain.Winsock1.State <> sckClosed Then
                frmMain.Winsock1.Close
            End If
    #End If
    frmConnect.Show
    Unload Me
    End If
End Sub

Private Sub Image2_Click()
MP3P.stopMP3
    If NumPjs >= 8 Then
        MsgBox "No podes crear mas personajes con esta cuenta, si queres algun otro, borra uno y crea"
        Exit Sub
    End If
    
    frmCrearPersonaje.Show
     Set MP3P = New clsMP3Player
    Call Extract_File2(MP3, App.Path & "\ARCHIVOS\", "3.mp3", Windows_Temp_Dir, False)
    MP3P.mp3file = Windows_Temp_Dir & "3.mp3"
    MP3P.stopMP3
    MP3P.playMP3
    MP3P.Volume = 1000
    Form2.Hide
End Sub

Private Sub Image3_Click()
MP3P.stopMP3
    Dim Cosa As String
    If NumPjs <= 0 Then Exit Sub
    
    For i = 1 To NumPjs
        If Option1(i).Visible = True Then
            If Option1(i).value = True Then
                UserName = Option1(i).Caption
                Exit For
            End If
        End If
    Next i
    
    EstadoLogin = adentrocuenta
    Call Login(ValidarLoginMSG(CInt(bRK)))
    Form2.Visible = False

End Sub

Private Sub Labe2_Click()

End Sub

Private Sub Image4_Click()
Dim PjABorrar As String

For i = 1 To Option1.Count
        PjABorrar = Label2.Caption
        Exit For
Next i

Debug.Print PjABorrar

If MsgBox("Seguro que desea borrar el pj " & Chr(34) & PjABorrar & Chr(34) & " ?", vbYesNo, "Confirmar borrado de pj.") = vbNo Then Exit Sub
Call SendData("~BORRA" & PjABorrar & "," & Cuenta)
End Sub



Private Sub Option1_Click(index As Integer)
    Dim SR As RECT
    Dim GrhIndex As Long
    
    SR.Left = 0
    SR.Top = 0
    SR.Right = 152
    SR.Bottom = 152
    
    CharSurface.BltColorFill SR, vbBlack 'Limpiamos la surface, no se si es la mejor forma de hacerlo.
    Label2.Caption = PjCuenta.Nombre(index)
    Mapa.Caption = PjCuenta.Pos(index)
    Oro.Caption = PjCuenta.Oro(index)
    Nivel.Caption = PjCuenta.Nivel(index)
            
    If PjCuenta.Body(index) > 0 Then
        GrhIndex = GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)
        'Get Source Rect
        SR.Left = GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).sX
        SR.Top = GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).sY
        SR.Right = GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).pixelWidth
        SR.Bottom = GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).pixelHeight

        'CharSurface.Blt DR, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SR, DDBLT_DONOTWAIT
        CharSurface.BltFast 152 / 2 - SR.Right / 2, 152 / 2 - SR.Bottom / 2, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SR, DDBLTFAST_SRCCOLORKEY
    Else: Exit Sub
    End If
    
    If PjCuenta.Head(index) > 0 Then
        GrhIndex = HeadData(PjCuenta.Head(index)).Head(E_Heading.SOUTH).GrhIndex
        'Get Source Rect
        SR.Left = GrhData(GrhIndex).sX
        SR.Top = GrhData(GrhIndex).sY
        SR.Right = GrhData(GrhIndex).sX + GrhData(GrhIndex).pixelWidth
        SR.Bottom = GrhData(GrhIndex).sY + GrhData(GrhIndex).pixelHeight
        
        'file = GrhData(GrhData(BodyData(PjCuenta.Body(Index)).Walk(E_Heading.SOUTH).grhindex).Frames(1)).FileNum
        CharSurface.BltFast (152 / 2 - SR.Right / 2), 152 / 2 + GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).pixelHeight / 2 - 16 + BodyData(PjCuenta.Body(index)).HeadOffset.Y, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SR, DDBLTFAST_SRCCOLORKEY
    End If
    
    If PjCuenta.Casco(index) > 0 Then
        GrhIndex = CascoAnimData(PjCuenta.Casco(index)).Head(E_Heading.SOUTH).GrhIndex
        If GrhIndex > 0 Then
            'Get Source Rect
            SR.Left = GrhData(GrhIndex).sX
            SR.Top = GrhData(GrhIndex).sY
            SR.Right = GrhData(GrhIndex).sX + GrhData(GrhIndex).pixelWidth
            SR.Bottom = GrhData(GrhIndex).sY + GrhData(GrhIndex).pixelHeight
            
            'file = GrhData(GrhData(BodyData(PjCuenta.Body(Index)).Walk(E_Heading.SOUTH).grhindex).Frames(1)).FileNum
            CharSurface.BltFast (152 / 2 - SR.Right / 2), 152 / 2 + GrhData(GrhData(BodyData(PjCuenta.Body(index)).Walk(E_Heading.SOUTH).GrhIndex).Frames(1)).pixelHeight / 2 - 16 + BodyData(PjCuenta.Body(index)).HeadOffset.Y, SurfaceDB.Surface(GrhData(GrhIndex).FileNum), SR, DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    If PjCuenta.Escu(index) > 0 Then
        GrhIndex = ShieldAnimData(PjCuenta.Escu(index)).ShieldWalk(E_Heading.SOUTH).GrhIndex
        If GrhIndex > 0 Then
            'Get Source Rect
            SR.Left = GrhData(GrhIndex).sX
            SR.Top = GrhData(GrhIndex).sY
            SR.Right = GrhData(GrhIndex).sX + GrhData(GrhIndex).pixelWidth
            SR.Bottom = GrhData(GrhIndex).sY + GrhData(GrhIndex).pixelHeight
            
            'file = GrhData(GrhData(BodyData(PjCuenta.Body(Index)).Walk(E_Heading.SOUTH).grhindex).Frames(1)).FileNum
            CharSurface.BltFast (152 / 2 - SR.Right / 2), 152 / 2 - SR.Bottom / 2, SurfaceDB.Surface(GrhData(GrhData(GrhIndex).Frames(1)).FileNum), SR, DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    If PjCuenta.Arma(index) > 0 Then
        GrhIndex = WeaponAnimData(PjCuenta.Arma(index)).WeaponWalk(E_Heading.SOUTH).GrhIndex
        If GrhIndex > 0 Then
            'Get Source Rect
            SR.Left = GrhData(GrhIndex).sX
            SR.Top = GrhData(GrhIndex).sY
            SR.Right = GrhData(GrhIndex).sX + GrhData(GrhIndex).pixelWidth
            SR.Bottom = GrhData(GrhIndex).sY + GrhData(GrhIndex).pixelHeight
            
            'file = GrhData(GrhData(BodyData(PjCuenta.Body(Index)).Walk(E_Heading.SOUTH).grhindex).Frames(1)).FileNum
            CharSurface.BltFast (152 / 2 - SR.Right / 2), 152 / 2 - SR.Bottom / 2, SurfaceDB.Surface(GrhData(GrhData(GrhIndex).Frames(1)).FileNum), SR, DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    SR.Left = 0
    SR.Top = 0
    SR.Right = 152
    SR.Bottom = 152
    
    CharSurface.BltToDC Picture1.hdc, SR, SR
    Picture1.Refresh
    
End Sub

Private Sub Option1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Mori As String
If PjCuenta.muerto(index) = 1 Then Mori = "Sí"
If PjCuenta.muerto(index) = 0 Then Mori = "No"


End Sub

