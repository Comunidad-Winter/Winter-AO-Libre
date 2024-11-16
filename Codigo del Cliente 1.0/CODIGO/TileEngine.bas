Attribute VB_Name = "Mod_TileEngine"
Option Explicit

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    C       O       N       S      T
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'    T       I      P      O      S
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tama�o y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type


'Lista de cuerpos
Public Type FxData
    Fx As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public UserGuerra As Boolean 'Guerras

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tama�o del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tama�o muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tama�o de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?�?�?�?�?�?�?�?�?�?�Totales?�?�?�?�?�?�?�?�?�?�?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?

Public lastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'�?�?�?�?�?�?�?�?�?�Graficos�?�?�?�?�?�?�?�?�?�?�?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Mapa?�?�?�?�?�?�?�?�?�?�?�?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�Usuarios?�?�?�?�?�?�?�?�?�?�?�?�?
'
'epa ;)
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

'�?�?�?�?�?�?�?�?�?�?�API?�?�?�?�?�?�?�?�?�?�?�?�?�?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?


'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'est� raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?

#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long
        
Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()
On Error Resume Next
Dim N As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #N, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #N, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #N, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #N

End Sub
Sub CargarFxs()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #N, , MisFxs(i)
    Call InitGrh(FxData(i).Fx, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #N

End Sub

Sub CargarTips()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumTips As Integer

N = FreeFile
Open App.Path & "\init\Tips.ayu" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumTips

'Resize array
ReDim Tips(1 To NumTips) As String * 255

For i = 1 To NumTips
    Get #N, , Tips(i)
Next i

Close #N

End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, i As Integer
Dim Nu As Integer

N = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #N, , bLluvia(i)
Next i

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.Y + CY

End Sub






Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(CharIndex).iHead = Head
charlist(CharIndex).iBody = Body
charlist(CharIndex).Head = HeadData(Head)
charlist(CharIndex).Body = BodyData(Body)
charlist(CharIndex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(CharIndex).Arma.WeaponAttack = 0

charlist(CharIndex).Escudo = ShieldAnimData(Escudo)
charlist(CharIndex).Casco = CascoAnimData(Casco)

charlist(CharIndex).Heading = Heading

'Reset moving stats
charlist(CharIndex).Moving = 0
charlist(CharIndex).MoveOffset.X = 0
charlist(CharIndex).MoveOffset.Y = 0

'Update position
charlist(CharIndex).Pos.X = X
charlist(CharIndex).Pos.Y = Y

'Make active
charlist(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).Fx = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).invisible = False

#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(CharIndex)
#End If

    charlist(CharIndex).Moving = 0
    charlist(CharIndex).muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).pie = False
    charlist(CharIndex).Pos.X = 0
    charlist(CharIndex).Pos.Y = 0
    charlist(CharIndex).UsandoArma = False

End Sub


Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurr�a debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
'[END]'

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

charlist(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Debug.Print UserCharIndex
    Call EraseChar(CharIndex)
End If

End Sub

Public Sub DoFogataFx()
If Sound Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata()
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
 
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) Then
                .pie = Not .pie
        'Si esta en una superficie de pasto?
                If MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 6000 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 6307 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS3)
                    Else
                        Call Audio.PlayWave(SND_PASOS4)
                    End If
                Else
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS1)
                    Else
                        Call Audio.PlayWave(SND_PASOS2)
                    End If
                End If
            End If
   
   
    If UserNavegando Then
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO)
    End If
   
    End With
    End If
End Sub



Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex


charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY

charlist(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(CharIndex).Fx
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(CharIndex).Fx = 0
    charlist(CharIndex).FxLoopTimes = 0
End If

If Not EstaPCarea(CharIndex) Then Dialogos.QuitarDialogo (CharIndex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        Y = -1

    Case E_Heading.EAST
        X = 1

    Case E_Heading.SOUTH
        Y = 1
    
    Case E_Heading.WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer




'Resize arrays
ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '�Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function




Function InMapLegalBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If
'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, Hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hWnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC Hdc, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
 
 
On Error Resume Next
 
If UserCiego Then Exit Sub
 
Dim Y        As Integer 'Keeps track of where on map we are
Dim X        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim TempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim PP       As RECT
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer
 
'Configuramos el tama�o de la pantalla
'Arreglado bug de agrandar mainviewshp
minY = (tiley - (WindowTileHeight \ 2)) - TileBufferSize
maxY = (tiley + (WindowTileHeight \ 2)) + TileBufferSize
minX = (tilex - (WindowTileWidth \ 2)) - TileBufferSize
maxX = (tilex + (WindowTileWidth \ 2)) + TileBufferSize
 
 
'Draw floor layer
ScreenY = 8
For Y = (minY + 8) To maxY - 8
ScreenX = 8
For X = minX + 8 To maxX - 8
If X > 100 Or Y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                            .FrameCounter = 1
                    End If
                End If
            End If
 
            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With
 
        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight
 
        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
 
            'Layer 2
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(2), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, _
                        1)
            End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y
 
'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)
 
'Mostrar capas transparentes
ScreenY = 8
For Y = minY + 8 To maxY - 1
ScreenX = 5
For X = minX + 5 To maxX - 5
If X > 100 Or X < -3 Then Exit For
iPPx = 32 * ScreenX - 32 + PixelOffsetX
iPPy = 32 * ScreenY - 32 + PixelOffsetY
 
 
 
        'Object Layer **********************************
        If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).ObjGrh, _
                        iPPx, iPPy, 1, 1)
        End If
        '***********************************************
        'Char layer ************************************
        If MapData(X, Y).CharIndex <> 0 Then
            TempChar = charlist(MapData(X, Y).CharIndex)
            PixelOffsetXTemp = PixelOffsetX
            PixelOffsetYTemp = PixelOffsetY
 
            Moved = 0
            'If needed, move left and right
            If TempChar.MoveOffset.X <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                TempChar.MoveOffset.X = TempChar.MoveOffset.X - (8 * Sgn(TempChar.MoveOffset.X))
                Moved = 1
            End If
            'If needed, move up and down
            If TempChar.MoveOffset.Y <> 0 Then
                TempChar.Body.Walk(TempChar.Heading).Started = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                TempChar.MoveOffset.Y = TempChar.MoveOffset.Y - (8 * Sgn(TempChar.MoveOffset.Y))
                Moved = 1
            End If
            'If done moving stop animation
            If Moved = 0 And TempChar.Moving = 1 Then
                TempChar.Moving = 0
                TempChar.Body.Walk(TempChar.Heading).FrameCounter = 1
                TempChar.Body.Walk(TempChar.Heading).Started = 0
                TempChar.Arma.WeaponWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                TempChar.Escudo.ShieldWalk(TempChar.Heading).FrameCounter = 1
                TempChar.Escudo.ShieldWalk(TempChar.Heading).Started = 0
            End If
           
            '[ANIM ATAK]
            If TempChar.Arma.WeaponAttack > 0 Then
                TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - 1
                If TempChar.Arma.WeaponAttack = 0 Then
                    TempChar.Arma.WeaponWalk(TempChar.Heading).Started = 0
                End If
            End If
            '[/ANIM ATAK]
           
            'Dibuja solamente players
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
 
            If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Or (UCase$(TempChar.Nombre) = UCase$(UserName) And UserNavegando = True) Then
                If Not charlist(MapData(X, Y).CharIndex).invisible Then
    #If (ConAlfaB = 1) Then
                        If TempChar.muerto Then
                            Call SurfaceConAB(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        Else
    #End If
                            Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
    #If ConAlfaB = 1 Then
                        End If
    #End If
                        '[END]'
                        '[CABEZA]'
    #If ConAlfaB = 1 Then
                        If TempChar.muerto Then
                            Call SurfaceConAB( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        Else
    #End If
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
    #If ConAlfaB = 1 Then
                        End If
    #End If
                        '[END]'
                        '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[END]'
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[END]'
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[END]'
                End If  'end if ~in
                 
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.X), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If
 
               
                 If Nombres Then
                    If TempChar.invisible = False Then
                        If TempChar.Nombre <> "" Then
                           Dim lCenter As Long
                            If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                lCenter = (frmMain.TextWidth(Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                Dim sClan As String: sClan = mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                Dim ColorClan As Long
                                ColorClan = RGB(231, 202, 157)
                               
                                Select Case TempChar.priv
                                Case 0
                                    If TempChar.Criminal Then
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                    Else
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                    End If
                                Case 25  'admin
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                Case Else 'el resto
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                End Select
                            Else
                                lCenter = (frmMain.TextWidth(TempChar.Nombre) / 2) - 16
 
                                Select Case TempChar.priv
                                Case 0
                                    If TempChar.Criminal Then
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
                                    Else
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b))
                                    End If
                                Case 7
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
                                Case Else
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
                                End Select
                            End If
                        End If
 
                'clan
 
            ElseIf TempChar.Nombre <> "" Then
                                             
                sClan = mid(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                Dim UserClan As String
                UserClan = Trim(mid(charlist(UserCharIndex).Nombre, InStr(charlist(UserCharIndex).Nombre, "<")))
                If charlist(UserCharIndex).Nombre = TempChar.Nombre Or UCase(UserClan) = UCase(Trim(sClan)) Then
                       
                       
                        If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                         lCenter = (frmMain.TextWidth(Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                         Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), RGB(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b))
                         lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                         Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, vbYellow)
                        Else
                         lCenter = (frmMain.TextWidth(TempChar.Nombre) / 2) - 16
                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, TempChar.Nombre, RGB(ColoresPJ(TempChar.priv).r, ColoresPJ(TempChar.priv).G, ColoresPJ(TempChar.priv).b))
                     
                       
                        End If
                       
                         'cuerpo
                        Call SurfaceConAB(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                         
                        'cabeza
                        Call SurfaceConAB( _
                                    BackBufferSurface, _
                                    TempChar.Head.Head(TempChar.Heading), _
                                    iPPx + TempChar.Body.HeadOffset.X, _
                                    iPPy + TempChar.Body.HeadOffset.Y, _
                                    1, 0)
 
           
                       '[Casco]'
                            If TempChar.Casco.Head(TempChar.Heading).GrhIndex <> 0 Then
                                Call SurfaceConAB( _
                                        BackBufferSurface, _
                                        TempChar.Casco.Head(TempChar.Heading), _
                                        iPPx + TempChar.Body.HeadOffset.X, _
                                        iPPy + TempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[END]'
                        '[ARMA]'
                            If TempChar.Arma.WeaponWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call SurfaceConAB( _
                                        BackBufferSurface, _
                                        TempChar.Arma.WeaponWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[END]'
                        '[Escudo]'
                            If TempChar.Escudo.ShieldWalk(TempChar.Heading).GrhIndex <> 0 Then
                                Call SurfaceConAB( _
                                        BackBufferSurface, _
                                        TempChar.Escudo.ShieldWalk(TempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                End If
             End If
          End If 'if Nombres
         
            Else 'If TempChar
 
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + TempChar.Body.HeadOffset.X), _
                            (iPPy + TempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If
 
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        TempChar.Body.Walk(TempChar.Heading), _
                        iPPx, iPPy, 1, 1)
                       
            End If 'TempChar
            charlist(MapData(X, Y).CharIndex) = TempChar
           
            If charlist(MapData(X, Y).CharIndex).Fx <> 0 Then
#If (ConAlfaB = 0) Then
 
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            FxData(TempChar.Fx).Fx, _
                            iPPx + FxData(TempChar.Fx).OffsetX, _
                            iPPy + FxData(TempChar.Fx).OffsetY, _
                            1, 1, MapData(X, Y).CharIndex)
#Else
                    Call SurfaceConAB( _
                            BackBufferSurface, _
                            FxData(TempChar.Fx).Fx, _
                            iPPx + FxData(TempChar.Fx).OffsetX, _
                            iPPy + FxData(TempChar.Fx).OffsetY, _
                            1, 1, MapData(X, Y).CharIndex)
#End If
            End If
        End If 'If MapData
        'Layer 3
        If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, Y).Graphic(3), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, 1)
        End If
       
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y
 
 
' TECHOS
If bTecho Then
        ScreenY = 5
    For Y = minY + 5 To maxY - 1
        If Y > 0 And Y < 101 Then 'In map bounds
            ScreenX = 5
        For X = minX + 5 To maxX
            If Y > 0 And Y < 101 Then 'In map bounds
    If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
        'Draw
        Call DDrawTransGrhtoSurfaceAlpha( _
        BackBufferSurface, _
        MapData(X, Y).Graphic(4), _
        ((32 * ScreenX) - 32) + PixelOffsetX, _
        ((32 * ScreenY) - 32) + PixelOffsetY, _
        1, 1)
    End If
End If
        ScreenX = ScreenX + 1
    Next X
End If
        ScreenY = ScreenY + 1
    Next Y
 
Else
 
        ScreenY = 5
            For Y = minY + 5 To maxY - 1
        If Y > 0 And Y < 101 Then 'In map bounds
            ScreenX = 5
            For X = minX + 5 To maxX
        If Y > 0 And Y < 101 Then 'In map bounds
If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
            BackBufferSurface, _
            MapData(X, Y).Graphic(4), _
            ((32 * ScreenX) - 32) + PixelOffsetX, _
            ((32 * ScreenY) - 32) + PixelOffsetY, _
            1, 1)
        End If
    End If
ScreenX = ScreenX + 1
    Next X
End If
    ScreenY = ScreenY + 1
    Next Y
End If
 
'LLuvia
If bLluvia(UserMap) = 1 And bRain Then
    If llTick < DirectX.TickCount - 50 Then
        iFrameIndex = iFrameIndex + 1
    If iFrameIndex > 7 Then iFrameIndex = 0
        llTick = DirectX.TickCount
End If
 
For Y = 0 To 4
    For X = 0 To 4
        Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.Surface(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
    Next X
Next Y
End If
 
    With PP
        .Left = 0
        .Top = 0
        .Right = WindowTileWidth * TilePixelWidth
        .Bottom = WindowTileHeight * TilePixelHeight
    End With
     '   If MiniMapY = True Then
        Call DibujarMiniMapa(frmMain.MiniMap)
 '   End If
    
Call CargarMacros

#If ConAlfaB Then
'Efectos
If EfectosDiaY = True Then

    If Anochecer = 1 Then
        EfectoNoche BackBufferSurface
    End If

    If Atardecer = 1 Then
        EfectoTarde BackBufferSurface
    End If

    If Amanecer = 1 Then
        EfectoAmanecer BackBufferSurface
    End If

    If Niebla = 1 Then
        EfectoNiebla BackBufferSurface
    End If
End If
#End If

End Sub
Public Function RenderSounds()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 4/22/2006
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 And Sound Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.Y <= Y
        
End If
End Function

Function PixelPos(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * X) - TilePixelWidth
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
    AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
     .Right = TilePixelWidth * (WindowTileWidth * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.Low = 0
ddck.High = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.status, "Cargando Gr�ficos....", 0, 0, 0, , , True
Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer
    
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
        End If

        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

        If IScombate Then Call Dialogos.DrawText(260, 260, "MODO COMBATE", vbRed)
        
'Estado del Dia
#If ConAlfaB Then
        If Amanecer Then
            Call Dialogos.DrawText(259, 625, "Ma�ana", &H80C0FF)
        ElseIf Atardecer Then
            Call Dialogos.DrawText(259, 625, "Tarde", &HC0C0C0)
        ElseIf Anochecer Then
            Call Dialogos.DrawText(259, 625, "Noche", &H808080)
        ElseIf Niebla Then
            Call Dialogos.DrawText(259, 625, "Niebla", vbWhite)
        Else
            Call Dialogos.DrawText(259, 625, "Dia", vbGreen)
        End If
#End If
        If CartelInvisibilidad Then Call Dialogos.DrawText(260, 520, CartelInvisibilidad & " segundos restantes de Invisibilidad", vbCyan)
        If CartelInvisibilidad Then Call Dialogos.DrawText(260, 520, CartelInvisibilidad & " segundos restantes de Invisibilidad", vbCyan)
                If UserGuerra Then Call Dialogos.DrawText(260, 270, "�Estas En Guerra!", vbRed) 'Guerras
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        Call Dialogos.DrawUsers
        Call DialogosClanes.Draw(Dialogos)
         
Call Dialogos.DrawCoordMapa

        Call DrawBackBufferSurface
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
'Saao Minimapa.
    Dim DDm As DDSURFACEDESC2
    DDm.lHeight = 101
    DDm.lWidth = 101
    DDm.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    DDm.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    Set SupMiniMap = DirectDraw.CreateSurface(DDm)
    Set SupBMiniMap = DirectDraw.CreateSurface(DDm)
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function

#If ConAlfaB Then

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT

Surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
Modo)

HayErrorAlpha:

If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If

End Sub

Public Sub EfectoTarde(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT

Surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()


Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 60, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 0, 0)

HayErrorAlpha:

If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If

End Sub

Public Sub EfectoAmanecer(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT

Surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()


Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 50, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 128, 64, 64)
HayErrorAlpha:

If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If

End Sub
Public Sub EfectoNiebla(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT

Surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()


Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 60, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 255, 255, 255)

HayErrorAlpha:

If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If

End Sub
#End If
Sub SurfaceConAB(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)

Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then

                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If

                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0

Surface.GetLockedArray dArray()
src.GetLockedArray sArray()
       
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de v�deo no esta en 555 o 565 o algo fall�."
  End
End If



vbDABLalphablend16 Modo, 1, sArray(SourceRect.Left * 2, SourceRect.Top), dArray(X + X, Y), AlphaX, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 0


Surface.Unlock rDest
src.Unlock SourceRect



Exit Sub

HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub

Sub TonalidadParalisis(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
 
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If
 
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
If Grh.GrhIndex = 0 Then Exit Sub
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
On Local Error GoTo HayErrorAlpha
src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.GetLockedArray dArray()
src.GetLockedArray sArray()
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de v�deo no esta en 555 o 565 o algo fall�."
  End
End If
Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 120, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 255, 255, 255)
Surface.Unlock rDest
src.Unlock SourceRect
Exit Sub
HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest
End Sub
 
#If ConAlfaB Then
Sub SurfaceColor(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
 
Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
 
                            If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                            If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                                charlist(KillAnim).Fx = 0
                                Exit Sub
                            End If
 
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
 
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
 
Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long
 
Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
 
src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
 
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
 
Surface.GetLockedArray dArray()
src.GetLockedArray sArray()
       
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de v�deo no esta en 555 o 565 o algo fall�."
  End
End If
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 50, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 255, 255, 0)
Surface.Unlock rDest
src.Unlock SourceRect
 
Exit Sub
 
HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest
 
End Sub
#End If
