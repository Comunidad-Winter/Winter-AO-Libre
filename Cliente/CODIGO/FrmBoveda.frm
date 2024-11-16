VERSION 5.00
Begin VB.Form frmBoveda 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Operación Bancaria"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4170
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Cantidad 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.ListBox LstBanco 
      Height          =   645
      ItemData        =   "FrmBoveda.frx":0000
      Left            =   120
      List            =   "FrmBoveda.frx":000D
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenido a la Cadena de Finanzas de Winter-AO... ¿En que le puedo servir?"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Depo 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "FrmBoveda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Acerptar_Click()

End Sub

Private Sub Aceptar_Click()
Select Case LstBanco.listIndex
    Case 0, -1 'Depositar
    
        'Negativos y ceros
        If (Val(Cantidad.Text) <= 0 And (UCase$(Cantidad.Text) <> "TODO")) Then Depo.Caption = "Cantidad inválida."
    
        If Val(Cantidad.Text) <= UserGLD Or UCase$(Cantidad.Text) = "TODO" Then
            Call SendData("/DEPOSITAR " & IIf(Val(Cantidad.Text) > 0, Val(Cantidad.Text), UserGLD))
            SendData ("FINBAN")
            Unload Me
        Else
            Depo.Caption = "No tienes esa cantidad. Escríbela nuevamente."
        End If
        
        Case 1 'Retirar
    
        'Negativos y ceros
        If (Val(Cantidad.Text) <= 0 And (UCase$(Cantidad.Text) <> "TODO")) Then Depo.Caption = "Cantidad inválida."
            Call SendData("/RETIRAR " & Cantidad.Text)
            SendData ("FINBAN")
            Unload Me
        
        
    Case 2 'Bóveda
        Unload Me
        End Select
End Sub



Private Sub LstBanco_Click()
Select Case LstBanco.listIndex
    Case 0 'Depositar
        Depo.Caption = "¿Cuánto deseas depositar?"
    Case 1 'Retirar
        Depo.Caption = "¿Cuánto deseas retirar?"
    Case 2 'Bóveda
        frmBancoObj.Show , frmMain
        Unload FrmBoveda
    Case 3 'Transferir
        Depo.Caption = "¿Qué cantidad deseas transferir?"
End Select
End Sub

