Attribute VB_Name = "Mod_MODOS_DE_VIDEO"
Option Explicit

'Testea si la maquina soporta un modo de video ;-)
Function SoportaDisplay(DD As DirectDraw7, DDSDaTestear As DDSURFACEDESC2) As Boolean
Dim ddsd As DDSURFACEDESC2
Dim DDEM As DirectDrawEnumModes

Set DDEM = DD.GetDisplayModesEnum(DDEDM_DEFAULT, ddsd)

Dim loopc As Integer
Dim flag As Boolean
loopc = 1
   
Do While loopc <> DDEM.GetCount And Not flag

    DDEM.GetItem loopc, ddsd
    flag = ddsd.lHeight = DDSDaTestear.lHeight _
    And ddsd.lWidth = DDSDaTestear.lWidth _
    And ddsd.ddpfPixelFormat.lRGBBitCount = _
    DDSDaTestear.ddpfPixelFormat.lRGBBitCount
    loopc = loopc + 1
Loop
SoportaDisplay = flag
End Function

Function ModosDeVideoIguales(dd1 As DDSURFACEDESC2, dd2 As DDSURFACEDESC2) As Boolean
ModosDeVideoIguales = _
    dd1.lHeight = dd2.lHeight _
    And dd1.lWidth = dd2.lWidth _
    And dd1.ddpfPixelFormat.lRGBBitCount = _
    dd2.ddpfPixelFormat.lRGBBitCount
End Function


