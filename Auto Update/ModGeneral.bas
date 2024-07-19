Attribute VB_Name = "ModGeneral"
'declaraciones generales
Public Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal _
cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Darchivo As String
Public strURL As String
Public i As Integer
Public EnProceso As Boolean
'end

'declaraciones zip
Private Type CBChar
    ch(4096) As Byte
End Type
Private Type UNZIPUSERFUNCTION
    UNZIPPrntFunction As Long
    UNZIPSndFunction As Long
    UNZIPReplaceFunction  As Long
    UNZIPPassword As Long
    UNZIPMessage  As Long
    UNZIPService  As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumFiles As Long
    Comment As Integer
End Type
Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    FPrivilege As Long
    Zip As String
    extractdir As String
End Type
Private Type ZIPnames
    s(0 To 99) As String
End Type
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long
'end

'subs zip
Public Sub UnZip(Zip As String, extractdir As String)
On Error GoTo err_Unzip

Dim Resultado As Long
Dim intContadorFicheros As Integer

Dim FuncionesUnZip As UNZIPUSERFUNCTION
Dim OpcionesUnZip As UNZIPOPTIONS

Dim NombresFicherosZip As ZIPnames, NombresFicheros2Zip As ZIPnames

NombresFicherosZip.s(0) = vbNullChar
NombresFicheros2Zip.s(0) = vbNullChar
FuncionesUnZip.UNZIPMessage = 0&
FuncionesUnZip.UNZIPPassword = 0&
FuncionesUnZip.UNZIPPrntFunction = DevolverDireccionMemoria(AddressOf UNFuncionParaProcesarMensajes)
FuncionesUnZip.UNZIPReplaceFunction = DevolverDireccionMemoria(AddressOf UNFuncionReplaceOptions)
FuncionesUnZip.UNZIPService = 0&
FuncionesUnZip.UNZIPSndFunction = 0&
OpcionesUnZip.ndflag = 1 'Carpetas incluidas >> [Bug Fixed]
OpcionesUnZip.C_flag = 1
OpcionesUnZip.fQuiet = 2
OpcionesUnZip.noflag = 1
OpcionesUnZip.Zip = Zip
OpcionesUnZip.extractdir = extractdir

Resultado = Wiz_SingleEntryUnzip(0, NombresFicherosZip, 0, NombresFicheros2Zip, OpcionesUnZip, FuncionesUnZip)

Exit Sub
err_Unzip:
    MsgBox "Unzip: " + Err.Description, vbExclamation
    Err.Clear
End Sub

Private Function UNFuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal X As Long) As Long
On Error GoTo err_UNFuncionParaProcesarMensajes

    UNFuncionParaProcesarMensajes = 0

Exit Function
err_UNFuncionParaProcesarMensajes:
    MsgBox "UNFuncionParaProcesarMensajes: " + Err.Description, vbExclamation
    Err.Clear
End Function

Private Function UNFuncionReplaceOptions(ByRef p As CBChar, ByVal L As Long, ByRef m As CBChar, ByRef Name As CBChar) As Integer
On Error GoTo err_UNFuncionReplaceOptions

    UNFuncionParaProcesarPassword = 0

Exit Function
err_UNFuncionReplaceOptions:
    MsgBox "UNFuncionParaProcesarPassword: " + Err.Description, vbExclamation
    Err.Clear
End Function
Public Function DevolverDireccionMemoria(Direccion As Long) As Long
On Error GoTo err_DevolverDireccionMemoria

    DevolverDireccionMemoria = Direccion

Exit Function
err_DevolverDireccionMemoria:
    MsgBox "DevolverDireccionMemoria: " + Err.Description, vbExclamation
    Err.Clear
End Function
'end

'subs generales
Sub WriteVar(File As String, Main As String, Var As String, value As String)
writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(File As String, Main As String, Var As String) As String
Dim L As Integer
Dim Char As String
Dim sSpaces As String
Dim szReturn As String
szReturn = ""
sSpaces = Space(5000)
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub addConsole(Texto As String, Rojo As Byte, Verde As Byte, Azul As Byte, Bold As Boolean, Italic As Boolean, Optional ByVal Enter As Boolean = False)
    With frmMain.RichTextBox1
        If (Len(.Text)) > 700 Then .Text = ""
        
        .SelStart = Len(.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        .SelColor = RGB(Rojo, Verde, Azul)
        
        .SelText = IIf(Enter, Texto, Texto & vbCrLf)
        
        .Refresh
    End With
frmMain.Caption = "Aut" & "oup" & "date" & " Winter" & "AO, v" & App.Major & "." & App.Minor
End Sub
'end
