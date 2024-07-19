VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_InterGTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Intervalo As Long
Private Tiempo As Long
Private Puede As Boolean
Public Sub init(Inter As Long)
Intervalo = Inter
Tiempo = GetTickCount + Intervalo
End Sub
Public Function Puedo() As Boolean
If GetTickCount < Intervalo - 1 Then 'me fijo que el tiempo no vuelva a cerro
Tiempo = GetTickCount + Intervalo
End If
If GetTickCount >= Tiempo Then
Puede = True
Tiempo = GetTickCount + Intervalo
End If
Puedo = Puede
Puede = False
End Function
