Attribute VB_Name = "Mod_General"
Option Explicit

Public bK As Long
Public bRK As Long


Public iplst As String
Public banners As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long
Public sHKeys() As String

Public Function DirPath(ByVal path As String) As String
'•Parra: Nuevo Engine v2.0
    Select Case path
        Case "Graficos"
            DirPath = App.path & "\GRAFICOS\"
            Exit Function
        
        Case "Sound"
            DirPath = App.path & "\WAV\"
            Exit Function
        
        Case "Midi"
            DirPath = App.path & "\MIDI\"
            Exit Function
        
        Case "Maps"
            DirPath = App.path & "\MAPAS\"
            Exit Function
    End Select
End Function

Public Function DirGraficos() As String
     DirGraficos = Windows_Temp_Dir 'compresion
    'Antiguo sistema de lectura de graficos
    'DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function SumaDigitos(ByVal numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (numero Mod 10)
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (numero Mod 10) - 1
        numero = numero \ 10
    Loop While (numero > 0)
End Function

Public Function Complex(ByVal numero As Integer) As Integer
    If numero Mod 2 <> 0 Then
        Complex = numero * SumaDigitos(numero)
    Else
        Complex = numero * SumaDigitosMenos(numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(numero)
    AuxInteger2 = SumaDigitosMenos(numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.y).CharIndex = loopc
        End If
    Next loopc
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next
Call ReleaseInstance
    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini
    'Call GenerarMiniMapa (SE HACE EN EL SWITCHMAP) - AGU
    
    'Unload the connect form
    Unload frmConnect
    
    frmMain.Label8.Caption = UserName
    Call SetMusicInfo("Jugando Winter-AO " & Chr(44) & "Nombre: " & UserName & Chr(44) & "Nivel: " & UserLvl & Chr(44) & " www.winter-ao.com.ar", "", "", , "{1}{0}")
    'Load main form
    frmMain.Visible = True
        FrmLanzador.Hide

End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub
Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    If UserParalizado Then Exit Sub
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.x, UserPos.y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.x + 1, UserPos.y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.x, UserPos.y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.x - 1, UserPos.y)
    End Select
    
    If LegalOk Then
        Call SendData("M" & Direccion)
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            engine.Char_Move_by_Head UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
    Call Audio.MoveListener(UserPos.x, UserPos.y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys() 'Stand
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
                If frmCustomKeys.Visible = True Then Exit Sub
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                End If
                Call MoveTo(NORTH)
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                End If
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                End If
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                End If
                Call MoveTo(WEST)
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            If kp Then Call RandomMove
                If frmMain.WorkMacro.Enabled Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                End If
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
    
        Case E_Heading.EAST
            x = 1
    
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim loopc As Long
    
    loopc = 1
    Do While charlist(loopc).active And loopc < UBound(charlist)
        loopc = loopc + 1
    Loop
    
    NextOpenChar = loopc
End Function

Sub SwitchMap(ByVal map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    handle = FreeFile()
    
    Open DirPath("Maps") & "Mapa" & map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            MapData(x, y).luz = 0
            MapData(x, y).particle_group = 0
            MapData(x, y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(x, y).Graphic(1).grhindex
            InitGrh MapData(x, y).Graphic(1), MapData(x, y).Graphic(1).grhindex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(x, y).Graphic(2).grhindex
                InitGrh MapData(x, y).Graphic(2), MapData(x, y).Graphic(2).grhindex
            Else
                MapData(x, y).Graphic(2).grhindex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(x, y).Graphic(3).grhindex
                InitGrh MapData(x, y).Graphic(3), MapData(x, y).Graphic(3).grhindex
            Else
                MapData(x, y).Graphic(3).grhindex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(x, y).Graphic(4).grhindex
                InitGrh MapData(x, y).Graphic(4), MapData(x, y).Graphic(4).grhindex
            Else
                MapData(x, y).Graphic(4).grhindex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(x, y).Trigger
            Else
                MapData(x, y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, y).CharIndex > 0 Then
                Call EraseChar(MapData(x, y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(x, y).ObjGrh.grhindex = 0
        Next x
    Next y
    
    Close handle
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = map
End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function



Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer
    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)
    End If
End Function

Public Function CurServerIp() As String
    If CurServer <> 0 Then
        CurServerIp = ServersLst(CurServer).Ip
    Else
        CurServerIp = frmConnect.IPTxt
    End If
End Function

Public Function CurServerPort() As Integer
    If CurServer <> 0 Then
        CurServerPort = ServersLst(CurServer).Puerto
    Else
        
    End If
End Function


Sub Main()
'TODO : Cambiar esto cuando se corrija el bug de los timers
'On Error GoTo ManejadorErrores
On Error Resume Next

    Call WriteClientVer

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8
    
        
    frmCargando.Show
    frmCargando.Refresh
    
    frmMain.Socket1.Startup
        
    Call InicializarNombres
    


UserMap = 1
    
    LoadGrhData
    CargarCabezas
    CargarCascos
    CargarCuerpos
    CargarArrayLluvia
    CargarFxs
    Call engine.Engine_Init
    Call engine.setup_ambient
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores

    
    Unload frmCargando
    
    'Inicializamos el sonido
     Call Audio.Initialize(frmMain.hWnd, DirPath("Sounds"), DirPath("Midi"))
    'Audio.MusicActivated = Not ClientSetup.bNoMusic
    'Audio.SoundActivated = Not ClientSetup.bNoSound
    'Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    
    If Musica Then
        Call Audio.PlayMIDI(MIdi_Inicio & ".mid")
    End If

    'frmPres.Picture = LoadPicture(App.path & "\Graficos\bosquefinal.jpg")
    'frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
        Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font

jojoparticulas
    
engine.Start
    
Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub
Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean

    HayAgua = MapData(x, y).Graphic(1).grhindex >= 1505 And _
                MapData(x, y).Graphic(1).grhindex <= 1520 And _
                MapData(x, y).Graphic(2).grhindex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ramx"

    CityDesc(1) = "Humilde ciudad de pescadores, unas de las ciudades mas antiguas de Winter-AO fundada por el Gran Dios Erwin. Desde su fundacion esta ciudad a vivido grandes batallas y a sido conquistada varias veces, ahora es una de las ciudades mas pacifica y neutral."

    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.Equitacion) = "Equitacion"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub

