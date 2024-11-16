Attribute VB_Name = "modDX8Requires"
Option Explicit

Public vertList(3) As TLVERTEX

Public Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public DX As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    x As Single
    y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu1 As Single
    tv1 As Single
    tu2 As Single
    tv2 As Single
End Type

Public Const PI As Single = 3.14159265358979
Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte
Type luzxhora
    r As Long
    g As Long
    b As Long
End Type
Public luz_dia(0 To 24) As luzxhora '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0

'JOJOJO
Public engine As New clsDX8Engine
'JOJOJO

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

'To get free bytes in RAM

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
Dim dblAns As Double
dblAns = (Bytes / 1024) / 1024
General_Bytes_To_Megabytes = format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Free_Ram() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

Sub jojoparticulas()
'Lorwik - Como poner Particulas
'Las N significa que no tengo ni puta idea de para que sirve... Las que tienen una F al final significa Final.
'Grav, significa Gravedad, Cant, signficia cantidad, Text, significa textura.
'-------------------------------------------------------------------------------------------------------------------
'                              N, Z , X ,Y ,N,Alpha,Red,Green,Blue,AlpF,RedF,GreenF,BlueF,Cant,Grav,Text,Tamaño,Vida
    engine.Particle_Group_Make 1, 0, 50, 70, 0, 100, 120, 120, 255, 100, 0, 0, 255, 0, 255, 50, -1, 21505, 50, 800
    'Engine.Particle_Group_Make 1, 0, 50, 49, 0, 20, 1, 255, 200, 80, 0, 10, 40, 40, 40, 200, -10, 609, 30, 100
    engine.Particle_Group_Make 2, 0, 44, 45, 0, 100, 120, 120, 255, 100, 0, 0, 255, 0, 255, 50, -1, 21505, 50, 800
    'Luces
    '                   X,   X,     Color, Tamaño
    engine.Light_Create 50, 50, &HFFCCCCCC, 10
    engine.Light_Create 50, 70, &HFFFFFFFF, 5
    engine.Light_Create 76, 48, &HFFFFFFFF, 4
    engine.Light_Render_All
End Sub

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal A As Long) As Long
        
    Dim C As Long
        
    If A > 127 Then
        A = A - 128
        C = A * 2 ^ 24 Or &H80000000
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    Else
        C = A * 2 ^ 24
        C = C Or r * 2 ^ 16
        C = C Or g * 2 ^ 8
        C = C Or b
    End If
    
    ARGB = C

End Function


