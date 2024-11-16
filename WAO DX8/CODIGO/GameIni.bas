Attribute VB_Name = "GameIni"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    Fx As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.desc = "Winter-AO Un mod de Argentum Online"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim N As Integer
    Dim GameIni As tGameIni
    N = FreeFile
    Open App.Path & "\init\Inicio.con" For Binary As #N
    Get #N, , MiCabecera
    
    Get #N, , GameIni
    
    Close #N
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
On Local Error Resume Next

Dim N As Integer
N = FreeFile
Open App.Path & "\init\Inicio.con" For Binary As #N
Put #N, , MiCabecera
Put #N, , GameIniConfiguration
Close #N
End Sub

