Attribute VB_Name = "Mod_TCP"
Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, index As Integer, M As Integer
    Dim T() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    sData = UCase$(Rdata)
     If Left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
     If Left$(sData, 4) = "INMO" Then CartelParalisis = Right$(sData, Len(sData) - 4)
     
    Select Case sData
            Case "CREADACUENTA"
            frmConnect.Show
            Exit Sub
            
        Case "BUENO"
            TimerPing(2) = GetTickCount()
            Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & (TimerPing(2) - TimerPing(1)) & " ms", 255, 0, 0, True, False, False)
            
        Case "LOKYED"            ' >>>>> LOGIN :: LOGGED
            logged = True
            UserCiego = False
            EngineRun = True
            IScombate = False
            UserDescansar = False
            Nombres = True
            If frmCrearPersonaje.Visible Then
                Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje
                Unload frmConnect
                frmMain.Show
            End If
            Call SetConnected
            Call DibujarMiniMapa
        
            'Mostramos el Tip
            If tipf = "1" And PrimeraVez Then
                 Call CargarTip
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "ET"
            UserEquitando = Not UserEquitando
            Exit Sub
        Case "FYNOC" ' Graceful exit ;))
Unload Form2
        Call SetMusicInfo("", "", "", "Games", , False)
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            IScombate = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            UserEquitando = False
            Call Reloguear
            Call Audio.StopWave
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
            
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            
            bK = 0
            
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
               Case "FINSUBOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmSubasta.List1(1).Clear
            Unload frmSubasta
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
                             Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        Case "INITSUB"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmSubasta.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmSubasta.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            'Comerciando = True
            frmSubasta.Show , frmMain
            Exit Sub

        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim II As Integer
            II = 1
            Do While II <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(II) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(II)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                II = II + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            FrmBoveda.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False

        Case "BORROK"
            Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "REAU" '<--- Requiere AutoUpdate
            Call frmMain.DibujarSatelite
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call frmMain.DibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call frmMain.DesDibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
        Case "M!"     ' <--- Usa meditando
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
    Case "³"
      Rdata = Right$(Rdata, Len(Rdata) - 1)
      NumUsers = Rdata
      Exit Sub
        Case "*"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)

#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveCharDecrypt(Rdata, CharIndex, X, Y)
#Else
            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
#End If

            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If

            Call MoveCharbyPos(CharIndex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
                        Case "="              ' Propagandas by SheKme 31/12/08
           Rdata = Right$(Rdata, Len(Rdata) - 1)
            If Not Rdata = "Apagar" Then
             frmMain.Propagandas.Visible = True
            frmMain.Propagandas.Navigate (Rdata)
            frmMain.ItemName.Top = 320
            Else
          frmMain.Propagandas.Visible = False
           frmMain.ItemName.Top = 344
            End If
           Exit Sub
        Case "+", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveNPCDecrypt(Rdata, CharIndex, X, Y, Left$(sData, 1) <> "*")
#Else
            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
#End If
            
            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4
            
'            If charlist(CharIndex).Body.Walk(1).GrhIndex = 4747 Then
'                Debug.Print "hola"
'            End If
            
            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, X, Y)
            'Call MoveCharbyPos(CharIndex, Asc(Mid$(Rdata, 3, 1)), Asc(Mid$(Rdata, 4, 1)))
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case Left$(sData, 2)
    Case "PJ"
            sData = Right$(sData, Len(sData) - 2)
            
            Dim Datos As String
            
            Dim NroDePj As Integer
            
            Unload frmConnect
            Form2.Show
            
            ReDim PjCuenta.Clase(1 To NumPjs)
            ReDim PjCuenta.muerto(1 To NumPjs)
            ReDim PjCuenta.Nivel(1 To NumPjs)
            ReDim PjCuenta.Nombre(1 To NumPjs)
            ReDim PjCuenta.Oro(1 To NumPjs)
            ReDim PjCuenta.Pos(1 To NumPjs)
            ReDim PjCuenta.status(1 To NumPjs)
            ReDim PjCuenta.Arma(1 To NumPjs)
            ReDim PjCuenta.Escu(1 To NumPjs)
            ReDim PjCuenta.Body(1 To NumPjs)
            ReDim PjCuenta.Head(1 To NumPjs)
            ReDim PjCuenta.Casco(1 To NumPjs)
            
            Datos = Right$(Datos, Len(Datos) - 1) 'saco el nro
            
            'Debug.Print "datos sin el nrito " & Datos
            For i = 1 To NumPjs
            
                Datos = ReadField(i, sData, Asc("~"))
                NroDePj = Val(ReadField(1, sData, Asc(","))) 'aca me tira el nro de pj que va
            
                If NroDePj <= 0 Then 'controlo que nose chotie
                    Exit Sub
                End If
            
                PjCuenta.Nombre(i) = ReadField(2, Datos, Asc(","))
                                
                PjCuenta.Oro(i) = ReadField(10, Datos, Asc(","))
                PjCuenta.Nivel(i) = ReadField(11, Datos, Asc(","))
                PjCuenta.Pos(i) = ReadField(9, Datos, Asc(","))
                
                PjCuenta.Head(i) = ReadField(3, Datos, Asc(","))
                PjCuenta.Body(i) = ReadField(4, Datos, Asc(","))
                PjCuenta.Casco(i) = ReadField(5, Datos, Asc(","))
                
                PjCuenta.Escu(i) = ReadField(6, Datos, Asc(","))
                PjCuenta.Arma(i) = ReadField(7, Datos, Asc(","))
                
                If ReadField(8, Datos, Asc(",")) <= 0 Then
                    Form2.Label2.ForeColor = vbRed
                Else
                    Form2.Label2.ForeColor = vbWhite
                End If
                
                'pongo en mayus la primer letra nomas
                PjCuenta.Clase(i) = Left$(PjCuenta.Clase(i), 1) & _
                Right$(LCase$(PjCuenta.Clase(i)), Len(PjCuenta.Clase(i)) - 1)
                PjCuenta.Nombre(i) = Left$(PjCuenta.Nombre(i), 1) & _
                Right$(LCase$(PjCuenta.Nombre(i)), Len(PjCuenta.Nombre(i)) - 1)
                'fin mayus
            
            Next i
        
            For i = 1 To NumPjs
                With Form2.Option1(i)
                    .Enabled = True
                    .Visible = True
                    .Caption = PjCuenta.Nombre(i)
                End With
            Next i
        Exit Sub
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.ExperienciaShp.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 112)
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            ' Dasi
            If UserLvl = 50 Then
            frmMain.exp.Caption = "¡Nivel Maximo!"
            End If
' /Dasi

            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            Exit Sub
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa
            If UserMap <> 1 Then
                Call Audio.StopWave
            End If
                 
#If SeguridadAlkon Then
            Call InitMI
#End If
            
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                   
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).CharIndex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            Exit Sub
            
          Case "MN"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            frmMain.MapName.Caption = Rdata
            
            Exit Sub
            
            Exit Sub
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        'Standelf
        Case "|?"
            Call frmGMAyuda.Show(vbModeless, frmMain)
            Exit Sub
        Case "\~"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Exit Sub
                    Case "|G" 'Guerras
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            If Rdata = 1 Then UserGuerra = True
            If Rdata = 0 Then UserGuerra = False
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub
            
        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        'Standelf
        Case "AF"
            If Val(Right$(Rdata, 1)) = 1 Then
                charlist(UserCharIndex).Criminal = True
            ElseIf Val(Right$(Rdata, 1)) = 0 Then
                charlist(UserCharIndex).Criminal = False
            End If
            Exit Sub
            
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            'Debug.Print "CC"
            'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
            '    Debug.Print "CHAR DUPLICADO: " & CharIndex
            '    Call EraseChar(CharIndex)
           ' End If
            
            charlist(CharIndex).Fx = Val(ReadField(9, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            charlist(CharIndex).Nombre = ReadField(12, Rdata, 44)
            charlist(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(CharIndex).priv = Val(ReadField(14, Rdata, 44))
            
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            
            Call RefreshAllChars
            Exit Sub
            
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            
            If charlist(CharIndex).Fx >= 40 And charlist(CharIndex).Fx <= 49 Then   'si esta meditando
                charlist(CharIndex).Fx = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).muerto = Val(ReadField(3, Rdata, 44)) = 500
            charlist(CharIndex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            charlist(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            charlist(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            charlist(CharIndex).Fx = Val(ReadField(7, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Casco = CascoAnimData(tempint)

            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            If Musica Then
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Sound Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")
            End If
            Exit Sub
            
            'Ladder 09/7/07

Case "PF" ' >>>>> Label de fuerza
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Fuerza.Caption = Rdata

Exit Sub

Case "PG" ' >>>>> Label de Agilidad
Rdata = Right$(Rdata, Len(Rdata) - 2)
frmMain.Agilidad.Caption = Rdata

Exit Sub
'Ladder 09/7/07
            
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)
     Case "AUR"
Rdata = Right$(Rdata, Len(Rdata) - 3)
CharIndex = Val(ReadField(1, Rdata, 44))
charlist(CharIndex).AuraIndex = Val(ReadField(2, Rdata, 44))
Call InitGrh(charlist(CharIndex).Aura, Val(ReadField(2, Rdata, 44)))
        Case "VAL"                  ' >>>>> Validar Cliente :: VAL
            Dim ValString As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            bK = CLng(ReadField(1, Rdata, Asc(",")))
            bRK = ReadField(2, Rdata, Asc(","))
            ValString = ReadField(3, Rdata, Asc(","))
            CargarCabezas
            
        If EstadoLogin = BorrarPj Then
                Call SendData("BORR" & "," & ValidarLoginMSG(CInt(Rdata)))
            ElseIf EstadoLogin = LogCuenta Or EstadoLogin = adentrocuenta Or EstadoLogin = CrearNuevoPj Then
                Call Login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            End If
            Exit Sub
            
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
      
            
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).Fx = Val(ReadField(2, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
            Exit Sub
            
            Case "NUB"
Rdata = Right$(Rdata, Len(Rdata) - 3)
If Rdata = 1 Then
    Amanecer = 0
            Atardecer = 0
        Anochecer = 1
    Niebla = 0
End If
If Rdata = 0 Then
Anochecer = 0
End If
Exit Sub

Case "MAÑ" 'Mañana
Rdata = Right$(Rdata, Len(Rdata) - 3)
If Rdata = True Then
    Amanecer = 1
        Atardecer = 0
        Anochecer = 0
    Niebla = 0
End If
If Rdata = False Then
Amanecer = False
End If
Exit Sub

Case "TAR" 'Tarde
Rdata = Right$(Rdata, Len(Rdata) - 3)
If Rdata = 1 Then
    Amanecer = 0
        Atardecer = 1
        Anochecer = 0
    Niebla = 0
End If
If Rdata = 0 Then
Atardecer = 0
End If
Exit Sub

Case "MDI" 'Dia
Rdata = Right$(Rdata, Len(Rdata) - 3)
If Rdata = 1 Then
    Amanecer = 0
        Atardecer = 0
        Anochecer = 0
    Niebla = 0
End If
Exit Sub

Case "NIE" 'Niebla
Rdata = Right$(Rdata, Len(Rdata) - 3)
If Rdata = 1 Then
    Amanecer = 0
        Atardecer = 0
        Anochecer = 0
    Niebla = 1
End If
If Rdata = 0 Then
Niebla = 0
End If
Exit Sub
            Case "ARM" ' fuerza y armaduras/escus/cascos en labels
Rdata = Right$(Rdata, Len(Rdata) - 3)

With frmMain
.Arma = ReadField(1, Rdata, Asc(","))
.Armadura = ReadField(2, Rdata, Asc(","))
.Casco = ReadField(3, Rdata, Asc(","))
.Escudo = ReadField(4, Rdata, Asc(","))
End With
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
frmMain.ManaBar.Caption = "" & UserMinMAN & " / " & UserMaxMAN & ""
frmMain.StaBar.Caption = "" & UserMinSTA & " / " & UserMaxSTA & ""
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
                   
            frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.HpBar.Caption = "" & UserMinHP & " / " & UserMaxHP & ""
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
        
            Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
            
            Exit Sub
            Case "PRE"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceived(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRB"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedB(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRT"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedT(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For M = 0 To UBound(ArmasHerrero)
                ArmasHerrero(M) = 0
            Next M
            i = 1
            M = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(M) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                M = M + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For M = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(M) = 0
            Next M
            i = 1
            M = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(M) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                M = M + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For M = 0 To UBound(ObjCarpintero)
                ObjCarpintero(M) = 0
            Next M
            i = 1
            M = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(M) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                M = M + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmPasswdSinPadrinos.MousePointer = 1
            If Not frmCrearPersonaje.Visible Then
#If UsarWrench = 1 Then
                frmMain.Socket1.Disconnect
#Else
                If frmMain.Winsock1.State <> sckClosed Then _
                    frmMain.Winsock1.Close
#End If
            End If
            MsgBox Rdata
            Exit Sub
    End Select
    
    
    Select Case Left$(sData, 4)
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
            Case "PCGN"
Dim Proceso As String
Dim Nombre As String
Rdata = Right$(Rdata, Len(Rdata) - 4)
Proceso = ReadField(1, Rdata, 44)
Nombre = ReadField(2, Rdata, 44)
Call frmProcesos.Show
frmProcesos.List1.AddItem Proceso
frmProcesos.Caption = "Procesos de " & Nombre
Case "PCGR" ' >>>>> Ver procesos
frmProcesos.List1.Clear
frmProcesos.Caption = ""
Rdata = Right$(Rdata, Len(Rdata) - 4)
CharIndex = Val(ReadField(1, Rdata, 44))
Call enumProc(CharIndex)
Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmMain.hambar.Caption = "" & UserMinHAM & " / " & UserMaxHAM & ""
frmMain.agubar.Caption = "" & UserMinAGU & " / " & UserMaxAGU & ""
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, Rdata, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(CharIndex)
            Else
                Call MI(CualMI).ResetInvisible(CharIndex)
            End If
#End If

            Exit Sub
        Case "ZMOTD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
        Case "DADOS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select
    
    Select Case Left(sData, 6)
    
        Case "NUMPJS"
            sData = Right$(sData, Len(sData) - 6)
            NumPjs = Val(sData)
            If NumPjs = 0 Then
                'Form2.Label1.Visible = False
                'Form2.Picture1.Visible = False
                Form2.Show
                Call MsgBox("No tienes PJs en esta cuenta, crea uno.")
            End If
        '---------------------------------
        'coco cuentas


        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
            Case "MUERTEE"
FrmHogar.Show , frmMain
Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Call SendData("RPU")
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).listIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).listIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                II = 1
                Do While II <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(II).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(II).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    II = II + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).listIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).listIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
            Call frmPanelGm.Show(vbModeless, frmMain)
            Exit Sub
            Case "PANTOR"
Call frmConsolaTorneo.Show(vbModeless, frmMain)
Exit Sub
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0
            End If
            Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
    
#If SeguridadAlkon Then
    If HandleCryptedData(Rdata) Then Exit Sub
    
    If HandleDataEx(Rdata) Then Exit Sub
#End If
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    If AuxCmd = "/PING" Then TimerPing(1) = GetTickCount()
    
    'Debug.Print ">> " & sdData

#If SeguridadAlkon Then
    bK = CheckSum(bK, sdData)


    'Agregamos el fin de linea
    sdData = sdData & "~" & bK & ENDC
#Else
    sdData = sdData & ENDC
#End If

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

Sub Login(ByVal valcode As Integer)

    If EstadoLogin = LogCuenta Then
        SendData ("XIKLON" & UserPassword & "," & Cuenta & "," & SegCode)
    ElseIf EstadoLogin = adentrocuenta Then
        SendData ("MIKOLO" & UserName & "," & SegCode)
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("NLOGIN" & frmCrearPersonaje.txtNombre & "," & UserPassword _
                & "," & UserRaza & "," & UserSexo & "," & UserClase _
                & "," & UserSkills(1) & "," & UserSkills(2) _
                & "," & UserSkills(3) & "," & UserSkills(4) _
                & "," & UserSkills(5) & "," & UserSkills(6) _
                & "," & UserSkills(7) & "," & UserSkills(8) _
                & "," & UserSkills(9) & "," & UserSkills(10) _
                & "," & UserSkills(11) & "," & UserSkills(12) _
                & "," & UserSkills(13) & "," & UserSkills(14) _
                & "," & UserSkills(15) & "," & UserSkills(16) _
                & "," & UserSkills(17) & "," & UserSkills(18) _
                & "," & UserSkills(19) & "," & UserSkills(20) _
                & "," & UserSkills(21) & "," & UserEmail _
                & "," & UserHogar & "," & Cuenta & "," & Actual)
            
    End If
End Sub

