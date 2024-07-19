VERSION 5.00
Begin VB.Form frmCustomKeys 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Otros"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   40
      Top             =   2760
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   1920
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   1920
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Macro de Trabajo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Ver Mapa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Capturar Pantalla"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Modo Seguro"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Modo Combate"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar y Salir"
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Teclas por defecto"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hablar"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   1920
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Hablar al Clan"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Hablar a Todos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   2280
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   2280
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   2280
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   2280
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   2280
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Atacar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Usar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Tirar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocultar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Robar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Domar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Agarrar"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones Personales"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   2280
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar/Ocultar Nombres"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Corregir Posicion"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Activar/Desactivar Musica"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movimiento"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Derecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Izquierda"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Abajo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Arriba"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private Sub Command1_Click()
Call CustomKeys.LoadDefaults
Dim i As Long

For i = 1 To CustomKeys.Count
    Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
Next i
End Sub

Private Sub Command2_Click()

Dim i As Long

For i = 1 To CustomKeys.Count
    If LenB(Text1(i).Text) = 0 Then
        Call MsgBox("Hay una o mas teclas no validas, por favor verifique.", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Argentum Online")
        Exit Sub
    Else
        Call CustomKeys.SaveCustomKeys
    End If
Next i

Unload Me
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub
Private Sub Form_Load()
    Dim i As Long
    
    For i = 1 To CustomKeys.Count
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Frame5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
    For i = 1 To CustomKeys.Count
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = "" 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub
