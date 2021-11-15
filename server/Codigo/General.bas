Attribute VB_Name = "General"
Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader

Sub DarCuerpoDesnudo(ByVal userindex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(userindex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 32
                    Else
                        UserList(userindex).Char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 40
                    Else
                        UserList(userindex).Char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
      Case "ORCO"
Select Case UCase$(UserList(userindex).Genero)
Case "HOMBRE"
If Mimetizado Then
UserList(userindex).CharMimetizado.Body = 434
Else
UserList(userindex).Char.Body = 434
End If
Case "MUJER"
If Mimetizado Then
UserList(userindex).CharMimetizado.Body = 432
Else
UserList(userindex).Char.Body = 432
End If
End Select
    Case "GNOMO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    
End Select

UserList(userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function




Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer


For i = 1 To TrashCollector.Count
    Dim d As cGarbage
    Set d = TrashCollector(1)
    Call EraseObj(SendTarget.ToMap, 0, d.Map, 1, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(1)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista



End Sub

Sub EnviarSpawnList(ByVal userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.Toindex, userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar

Prision.Map = 127
Libertad.Map = 127

Prision.X = 52
Prision.Y = 44
Libertad.X = 52
Libertad.Y = 50
Denuncias = True

LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

If App.PrevInstance Then
Call MsgBox("Ya tenes abierto el server", vbInformation, "Error")
End
End If

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty
ReDim Guilds(1 To MAX_GUILDS) As clsClan

LoadEncrypt


IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "ORCO"

Torneo_Clases_Validas(1) = "Guerrero"
Torneo_Clases_Validas(2) = "Mago"
Torneo_Clases_Validas(3) = "Paladin"
Torneo_Clases_Validas(4) = "Clerigo"
Torneo_Clases_Validas(5) = "Bardo"
Torneo_Clases_Validas(6) = "Asesino"
Torneo_Clases_Validas(7) = "Druida"
Torneo_Clases_Validas(8) = "Cazador"
 
Torneo_Alineacion_Validas(1) = "Criminal"
Torneo_Alineacion_Validas(2) = "Ciudadano"
Torneo_Alineacion_Validas(3) = "Armada Caos"
Torneo_Alineacion_Validas(4) = "Armada Real"

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
ListaClases(13) = "Le�ador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apu�alar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Equitacion"

frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList
Call CargarForbidenWords
'�?�?�?�?�?�?�?� CARGAMOS DATOS DESDE ARCHIVOS �??�?�?�?�?�?�?�
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero
Call LoadQuests

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�

With frmMain
    .AutoSave.Enabled = True
    .tPiqueteC.Enabled = True
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenFORCOnnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'�?�?�?�?�?�?�?�?�?�?�?�?�?�?��?�?�?�?�?�?�?�?�?�?�?�?�?�?�




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(file, FileType) <> ""
End Function

Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
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
Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub
Public Sub LogSoportes(ByVal GM As String, ByVal Respuesta As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile
Open App.Path & "\logs\Soportes.log" For Append Shared As #nfile
Print #nfile, " "
'Print #nfile, "** " & Date & " ** " & Time & " " & Desc
Print #nfile, "** " & GM
Print #nfile, "** " & Respuesta
Print #nfile, " "
Close #nfile

Exit Sub

errhandler:
End Sub
Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If Consejero Then
    Open App.Path & "\logs\consejeros\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenFORCOnnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal userindex As Integer) As Boolean
    
    If MapInfo(UserList(userindex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 1 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 2 And _
           MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function
Public Sub TiempoInvocacion(ByVal userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userindex As Integer)

Dim modifi As Integer

If UserList(userindex).Counters.Frio < IntervaloFrio Then
  UserList(userindex).Counters.Frio = UserList(userindex).Counters.Frio + 1
Else

  If MapInfo(UserList(userindex).Pos.Map).Terreno = Desierto Then
  Call SendData(SendTarget.Toindex, userindex, UserList(userindex).Pos.Map, "CLX")
  UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinAGU - 5
  End If


  If MapInfo(UserList(userindex).Pos.Map).Terreno = Nieve Then
    Call SendData(SendTarget.Toindex, userindex, 0, "PRE9")
    Call SendData(SendTarget.Toindex, userindex, UserList(userindex).Pos.Map, "FRX")
    UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 5
    modifi = Porcentaje(UserList(userindex).Stats.MaxHP, 5)
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - modifi
    If UserList(userindex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.Toindex, userindex, 0, "||��Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(userindex).Stats.MinHP = 0
             
            If userindex = GranPoder Then
                Call SendData(SendTarget.ToAll, 0, 0, "PRE8," & UserList(userindex).name)
                Call OtorgarGranPoder(0)
            End If
            Call UserDie(userindex)
    End If
    Call SendData(SendTarget.Toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(userindex).Stats.MaxSta, 5)
    Call QuitarSta(userindex, modifi)
    Call SendData(SendTarget.Toindex, userindex, 0, "ASS" & UserList(userindex).Stats.MinSta)
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||��Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(userindex).Counters.Frio = 0
  
  
End If
End Sub


Public Sub EfectoMimetismo(ByVal userindex As Integer)

If UserList(userindex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(userindex).Counters.Mimetismo = UserList(userindex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.Toindex, userindex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        
    
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, userindex, UserList(userindex).Pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
End If
            
End Sub
Public Sub EfectoInvisibilidad(ByVal userindex As Integer)
Dim TiempoTranscurrido As Long
 
If UserList(userindex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(userindex).Counters.Invisibilidad = UserList(userindex).Counters.Invisibilidad + 1
    TiempoTranscurrido = (UserList(userindex).Counters.Invisibilidad * frmMain.GameTimer.Interval)
   
    If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
        If TiempoTranscurrido = 40 Then
            Call SendData(SendTarget.Toindex, userindex, 0, "INVI" & ((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000))
        Else
            Call SendData(SendTarget.Toindex, userindex, 0, "INVI" & (((IntervaloInvisible * frmMain.GameTimer.Interval) / 1000) - (TiempoTranscurrido / 1000)))
        End If
    End If
Else
    UserList(userindex).Counters.Invisibilidad = 0
    UserList(userindex).flags.Invisible = 0
    If UserList(userindex).flags.Oculto = 0 Then
        Call SendData(SendTarget.Toindex, userindex, 0, "PRB39")
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).Pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
        Call SendData(SendTarget.Toindex, userindex, 0, "INVI0")
    End If
End If
 
End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal userindex As Integer)

If UserList(userindex).Counters.Ceguera > 0 Then
    UserList(userindex).Counters.Ceguera = UserList(userindex).Counters.Ceguera - 1
Else
    If UserList(userindex).flags.Ceguera = 1 Then
        UserList(userindex).flags.Ceguera = 0
        Call SendData(SendTarget.Toindex, userindex, 0, "NSEGUE")
    End If
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call SendData(SendTarget.Toindex, userindex, 0, "NESTUP")
    End If

End If


End Sub
Public Sub EfectoParalisisUser(ByVal userindex As Integer)
Dim TiempoTranscurrido As Long
 
If UserList(userindex).Counters.Paralisis > 0 Then
    UserList(userindex).Counters.Paralisis = UserList(userindex).Counters.Paralisis - 1
    TiempoTranscurrido = (IntervaloParalizado * frmMain.GameTimer.Interval) - (UserList(userindex).Counters.Paralisis * frmMain.GameTimer.Interval)
   
    If TiempoTranscurrido Mod 1000 = 0 Or TiempoTranscurrido = 40 Then
        If TiempoTranscurrido = 40 Then
            Call SendData(SendTarget.Toindex, userindex, 0, "INMO" & ((IntervaloParalizado * frmMain.GameTimer.Interval) / 1000))
        Else
            Call SendData(SendTarget.Toindex, userindex, 0, "INMO" & (((IntervaloParalizado * frmMain.GameTimer.Interval) / 1000) - (TiempoTranscurrido / 1000)))
        End If
    End If
Else
    UserList(userindex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(SendTarget.Toindex, userindex, 0, "PARADOK")
    Call SendData(SendTarget.Toindex, userindex, 0, "INMO0")
    Call SendData(SendTarget.Toindex, userindex, 0, "PU" & UserList(userindex).Pos.X & "," & UserList(userindex).Pos.Y)
    Call SendData(SendTarget.Toindex, userindex, 0, "||Has recuperado la movilidad." & FONTTYPE_INFO)
End If
 
End Sub

Public Sub RecStamina(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 1 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 2 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta Then
   If UserList(userindex).Counters.STACounter < Intervalo Then
       UserList(userindex).Counters.STACounter = UserList(userindex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(userindex).Counters.STACounter = 0
       massta = RandomNumber(1, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
       UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta + massta
       If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then
            UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(userindex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(userindex).Counters.Veneno = UserList(userindex).Counters.Veneno + 1
Else
Call SendData(SendTarget.Toindex, userindex, 0, "PRE10")
  UserList(userindex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - N
  If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
    If userindex = GranPoder And UserList(userindex).Stats.MinHP <= 0 Then
            Call SendData(SendTarget.ToAll, userindex, 0, "PRE8," & UserList(userindex).name)
            Call OtorgarGranPoder(0)
    End If
  Call SendData(SendTarget.Toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
End If

End Sub

Public Sub DuracionPociones(userindex As Integer)

'Controla la duracion de las pociones
If UserList(userindex).flags.DuracionEfecto > 0 Then
   UserList(userindex).flags.DuracionEfecto = UserList(userindex).flags.DuracionEfecto - 1
   If UserList(userindex).flags.DuracionEfecto = 0 Then
        UserList(userindex).flags.TomoPocion = False
        UserList(userindex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(userindex).Stats.UserAtributos(loopX) = UserList(userindex).Stats.UserAtributosBackUP(loopX)
            Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "PF" & UserList(userindex).Stats.UserAtributos(Fuerza))
Call SendData(ToPCArea, userindex, UserList(userindex).Pos.Map, "PG" & UserList(userindex).Stats.UserAtributos(Agilidad))
        Next
   End If
End If

End Sub

Public Sub HambreYSed(userindex As Integer, fenviarAyS As Boolean)
'Sed
If UserList(userindex).Stats.MinAGU > 0 Then
    If UserList(userindex).Counters.AGUACounter < IntervaloSed Then
          UserList(userindex).Counters.AGUACounter = UserList(userindex).Counters.AGUACounter + 1
    Else
          UserList(userindex).Counters.AGUACounter = 0
          UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
                            
          If UserList(userindex).Stats.MinAGU <= 0 Then
               UserList(userindex).Stats.MinAGU = 0
               UserList(userindex).flags.Sed = 1
          End If
                            
          fenviarAyS = True
                            
    End If
End If

'hambre
If UserList(userindex).Stats.MinHam > 0 Then
   If UserList(userindex).Counters.COMCounter < IntervaloHambre Then
        UserList(userindex).Counters.COMCounter = UserList(userindex).Counters.COMCounter + 1
   Else
        UserList(userindex).Counters.COMCounter = 0
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 10
        If UserList(userindex).Stats.MinHam <= 0 Then
               UserList(userindex).Stats.MinHam = 0
               UserList(userindex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 1 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 2 And _
   MapData(UserList(userindex).Pos.Map, UserList(userindex).Pos.X, UserList(userindex).Pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
   If UserList(userindex).Counters.HPCounter < Intervalo Then
      UserList(userindex).Counters.HPCounter = UserList(userindex).Counters.HPCounter + 1
   Else
      mashit = RandomNumber(2, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
                           
      UserList(userindex).Counters.HPCounter = 0
      UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + mashit
      If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
      Call SendData(SendTarget.Toindex, userindex, 0, "PRB33")
      EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub

Sub PasarSegundo()
    Dim i As Integer
    For i = 1 To LastUser
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1

                Call SendData(SendTarget.Toindex, i, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
                Call SendData(SendTarget.Toindex, i, 0, "FINOK")
                
                Call CloseSocket(i)
                Exit Sub
            End If
        
        'ANTIEMPOLLOS
        ElseIf UserList(i).flags.EstaEmpo = 1 Then
             UserList(i).EmpoCont = UserList(i).EmpoCont + 1
             If UserList(i).EmpoCont = 30 Then
                 
                 'If FileExist(CharPath & UserList(Z).Name & ".chr", vbNormal) Then
                 'esto siempre existe! sino no estaria logueado ;p
                 
                 'TmpP = val(GetVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant"))
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant", TmpP + 1)
                 'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "P" & TmpP + 1, LCase$(UserList(Z).Name) & ": CARCEL " & 30 & "m, MOTIVO: Empollando" & " " & Date & " " & Time)

                 'Call Encarcelar(Z, 30, "El sistema anti empollo")
                 Call SendData(SendTarget.Toindex, i, 0, "!! Fuiste expulsado por permanecer muerto sobre un item")
                 'Call SendData(SendTarget.ToAdmins, Z, 0, "|| " & UserList(Z).Name & " Fue encarcelado por empollar" & FONTTYPE_INFO)
                 UserList(i).EmpoCont = 0
                 Call CloseSocket(i)
                 Exit Sub
             ElseIf UserList(i).EmpoCont = 15 Then
                 Call SendData(SendTarget.Toindex, i, 0, "|| LLevas 15 segundos bloqueando el item, mu�vete o ser�s desconectado." & FONTTYPE_WARNING)
             End If
         End If
    Next i
    
If CuentaRegresiva > 0 Then
If CuentaRegresiva > 1 Then
Call SendData(SendTarget.ToAll, 0, 0, "||Contando..." & CuentaRegresiva - 1 & FONTTYPE_GUILD)
Else
Call SendData(SendTarget.ToAll, 0, 0, "||YA!!!!!!!!!" & "~255~0~0~1~1")
End If
CuentaRegresiva = CuentaRegresiva - 1
End If

If AdiosRegresiva > 0 Then
If AdiosRegresiva > 1 Then
Call SendData(SendTarget.ToAll, 0, 0, "||Cerrando...Se cerrar� el juego en " & AdiosRegresiva - 1 & " segundos..." & FONTTYPE_INFO)
Else
Call SendData(SendTarget.ToAll, 0, 0, "||Gracias por jugar a Winter-AO Return" & "~255~0~0~1~1")
End If
AdiosRegresiva = AdiosRegresiva - 1
End If

End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Grabando Personajes" & FONTTYPE_SERVER)
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Personajes Grabados" & FONTTYPE_SERVER)
    Call SendData(SendTarget.ToAll, 0, 0, "BKW")

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub
Public Sub Recordatorios()
Call SendData(SendTarget.ToAll, 0, 0, "||Servidor >>>  No esta permitido: insultos, propagandas de otros servidores, u ofertas de items en el chat, de lo contrario, seran advertidos por algun Administrador." & FONTTYPE_SERVER)
Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> No esta permitido: insultos, propagandas de otros servidores, mandar denuncias pidiendole cosas a los GM's o llamandolos por su Nick, piquetes, u otras cosas que no estan permitidas, de lo contrario, ser�n advertidos. Para evitar todo este tipo de inconvenientes, visitese el manual en nuestra web accediendo a www.winter-ao.com.ar o contactandose con el email de algun GM. Muchas Gracias." & FONTTYPE_SERVER)
End Sub
Function ZonaCura(ByVal userindex As Integer) As Boolean
Dim X As Integer, Y As Integer
For Y = UserList(userindex).Pos.Y - MinYBorder + 1 To UserList(userindex).Pos.Y + MinYBorder - 1
        For X = UserList(userindex).Pos.X - MinXBorder + 1 To UserList(userindex).Pos.X + MinXBorder - 1
       
            If MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex).NPCtype = 1 Then
                    If Distancia(UserList(userindex).Pos, Npclist(MapData(UserList(userindex).Pos.Map, X, Y).NpcIndex).Pos) < 10 Then
                        ZonaCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
ZonaCura = False
End Function
 Public Sub SwapObjects(ByVal userindex As Integer)
Dim tmpUserObj As UserOBJ
 
    With UserList(userindex)
               
        'Cambiamos si alguno es una herramienta
        If .Invent.HerramientaEqpSlot = ObjSlot1 Then
            .Invent.HerramientaEqpSlot = ObjSlot2
        ElseIf .Invent.HerramientaEqpSlot = ObjSlot2 Then
            .Invent.HerramientaEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un armor
        If .Invent.ArmourEqpSlot = ObjSlot1 Then
            .Invent.ArmourEqpSlot = ObjSlot2
        ElseIf .Invent.ArmourEqpSlot = ObjSlot2 Then
            .Invent.ArmourEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un barco
        If .Invent.BarcoSlot = ObjSlot1 Then
            .Invent.BarcoSlot = ObjSlot2
        ElseIf .Invent.BarcoSlot = ObjSlot2 Then
            .Invent.BarcoSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un casco
        If .Invent.CascoEqpSlot = ObjSlot1 Then
            .Invent.CascoEqpSlot = ObjSlot2
        ElseIf .Invent.CascoEqpSlot = ObjSlot2 Then
            .Invent.CascoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un escudo
        If .Invent.EscudoEqpSlot = ObjSlot1 Then
            .Invent.EscudoEqpSlot = ObjSlot2
        ElseIf .Invent.EscudoEqpSlot = ObjSlot2 Then
            .Invent.EscudoEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es munici�n
        If .Invent.MunicionEqpSlot = ObjSlot1 Then
            .Invent.MunicionEqpSlot = ObjSlot2
        ElseIf .Invent.MunicionEqpSlot = ObjSlot2 Then
            .Invent.MunicionEqpSlot = ObjSlot1
        End If
       
        'Cambiamos si alguno es un arma
        If .Invent.WeaponEqpSlot = ObjSlot1 Then
            .Invent.WeaponEqpSlot = ObjSlot2
        ElseIf .Invent.WeaponEqpSlot = ObjSlot2 Then
            .Invent.WeaponEqpSlot = ObjSlot1
        End If
       
        'Hacemos el intercambio propiamente dicho
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        'Actualizamos los 2 slots que cambiamos solamente
        Call UpdateUserInv(False, userindex, ObjSlot1)
        Call UpdateUserInv(False, userindex, ObjSlot2)
    End With
End Sub
 
    Public Sub StatisticaLwK()

STAT_MAXELV = val(GetVar(IniPath & "Server.ini", "STAT", "Nivel"))
STAT_MAXHP = val(GetVar(IniPath & "Server.ini", "STAT", "MaxHP"))
STAT_MAXSTA = val(GetVar(IniPath & "Server.ini", "STAT", "MaxEnergia"))
STAT_MAXMAN = val(GetVar(IniPath & "Server.ini", "STAT", "MaxMana"))
STAT_MAXHIT_UNDER36 = val(GetVar(IniPath & "Server.ini", "STAT", "MAXHITUNDER"))
STAT_MAXHIT_OVER36 = val(GetVar(IniPath & "Server.ini", "STAT", "MAXHITOVER"))
STAT_MAXDEF = val(GetVar(IniPath & "Server.ini", "STAT", "MaxDEF"))
EXPI = val(GetVar(IniPath & "Server.ini", "STAT", "Exp"))
OROI = val(GetVar(IniPath & "Server.ini", "STAT", "Oro"))
 End Sub
