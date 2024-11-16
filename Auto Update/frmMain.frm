VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "AutoUpdate"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   240
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin vbalProgBarLib6.vbalProgressBar Psb1 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "5"
      Top             =   1320
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   344
      Picture         =   "frmMain.frx":0945
      ForeColor       =   0
      BarPicture      =   "frmMain.frx":0961
      BarPictureMode  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Segments        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H0000FF00&
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H0000FF00&
      Height          =   495
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AutoUpdate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'>> Objetos
        Private Sub Command1_Click()
            If EnProceso Then Exit Sub
                Call addConsole("Conectando...", 0, 200, 0, False, False) '>> Informacion
                EnProceso = True
                Analizar 'Iniciamos la función Analizar =).
        End Sub
        Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            Command1.FontSize = Command1.FontSize - 1
        End Sub
        Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
            Command1.FontSize = Command1.FontSize + 1
        End Sub
        Private Sub Form_Load()
            'Posicionamos el formulario
            Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End Sub
        
Private Sub Label1_Click()
End
End Sub

        Private Sub Timer1_Timer()
            'Avance de la barra de descarga
            If Psb1.value = 90 Then
                Timer1.Enabled = False
            Else
                Psb1.value = Psb1.value + 5
            End If
            Psb1.Text = CLng(Psb1.Percent) & "%"
        End Sub
'<< End


'>> Funciones/Subs
        Function Analizar()
            On Error Resume Next
            
            Dim iX As Integer
            Dim tX As Integer
            Dim DifX As Integer
            Dim strsX As String
            
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
                iX = Inet1.OpenURL("http://wao.webcindario.com/VEREXE.txt")
            'Variable que contiene el numero de actualización del cliente
                tX = GetVar(App.Path & "\INIT\Update.ini", "INIT", "X")
            'Variable con la diferencia de actualizaciones servidor-cliente
                DifX = iX - tX
            
            If Not (DifX = 0) Then 'Si la diferencia no es nula,
            Call addConsole("Iniciando, se descargarán " & DifX & " actualizaciones.", 200, 200, 200, True, False)   '>> Informacion
                For i = 1 To DifX 'Descargamos todas las versiones de diferencia
'LINK2
                    strURL = "http://www.wao.net23.net/Parche" & CStr(i + tX) & ".zip" 'URL del parche .zip
                    Darchivo = App.Path & "\INIT\Parche" & i + tX & ".zip" 'Directorio del parche
                        Call addConsole("   Descargando parche nº " & i, 0, 0, 255, False, True)    '>> Informacion
                    Call AutoDownload(i + tX) 'Descargamos todas las versiones faltantes a partir de la nuestra
                        Call addConsole("   Parche nº " & i & " descargado satisfactoriamente.", 0, 0, 255, False, True)    '>> Informacion
                
                  Call addConsole(" Actualizaciones: " & i & "/" & DifX, 100, 100, 100, True, False)   '>> Informacion
                Next i
            Else
                Call addConsole("No hay actualizaciones pendientes", 200, 200, 200, True, False)    '>> Informacion
            End If
            
            
            Call WriteVar(App.Path & "\INIT\Update.ini", "INIT", "X", CStr(iX)) 'Avisamos al cliente que está actualizado
            
            EnProceso = False
            
            Call addConsole("El cliente ya está listo para jugar", 200, 200, 200, True, False)  '>> Informacion
            sRGY.Picture = sG.Picture
            
        End Function
        
        Public Sub AutoDownload(Numero As Integer)
            On Error Resume Next
            
            sRGY.Picture = sR.Picture
            
            Inet1.AccessType = icUseDefault
            Dim B() As Byte
            
            'Informacion...
            Psb1.value = 0
            Timer1.Enabled = True
            
            B() = Inet1.OpenURL(strURL, icByteArray)
            
            'Descargamos y guardamos el archivo
            Open Darchivo For Binary Access _
            Write As #1
            Put #1, , B()
            Close #1
            Timer1.Enabled = False
            Psb1.value = 100
            Timer1.Enabled = False
            
            'Informacion
            Call addConsole("   Instalando actualización.", 0, 100, 255, False, False)    '>> Informacion
            
            sRGY.Picture = sY.Picture
            
            'Unzipeamos
            UnZip Darchivo, App.Path & "\"
            
            'Borramos el zip
            Kill Darchivo
        End Sub
'<< End
