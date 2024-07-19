Attribute VB_Name = "TCP_HandleData2"
Option Explicit

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim OfertaSUB As Long
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rData)
        Case "/POWAH"
            Dim IkDex
            Dim O
            Dim PkDex
                        O = 0
        For IkDex = 1 To LastUser
            If UserList(IkDex).Stats.ELV > O Then
                If UserList(UserIndex).flags.Privilegios = 0 Then
                O = UserList(IkDex).Stats.ELV
            PkDex = UserList(IkDex).name
                End If
            End If
        IkDex = IkDex + 1
            Next IkDex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario mas fuerte online es: " & PkDex & ", quien es lvl " & O & " " & FONTTYPE_VENENO)
        Exit Sub
                Case "/TORNEO"
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
           
           'Si esta en la carcel no se va
If UserList(UserIndex).Pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If
           
            If Hay_Torneo = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Torneo.Longitud >= Torneo_Cantidad Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El torneo está lleno." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            For i = 1 To 8
                If UCase$(UserList(UserIndex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Next
           
            If Not Torneo.Existe(UserList(UserIndex).name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás en la lista de espera del torneo. Eres el participante nº " & Torneo.Longitud + 1 & FONTTYPE_INFO)
                Call Torneo.Push("", UserList(UserIndex).name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(UserIndex).name & "]" & FONTTYPE_INFOBOLD)
                If Torneo.Longitud = Torneo_Cantidad Then Call SendData(SendTarget.ToAll, 0, 0, "||El torneo se ha llenado!." & FONTTYPE_CELESTE_NEGRITA)
                If Torneo_SumAuto = 1 Then
                    FuturePos.Map = Torneo_Map
                    FuturePos.X = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
            End If
           
            Exit Sub
        Case "/ONLINE"
            N = 0
            tStr = ""
            For LoopC = 1 To LastUser
                If (UserList(LoopC).name <> "") And UserList(LoopC).flags.Privilegios <= 1 Then
                    N = N + 1
                    tStr = tStr & UserList(LoopC).name & ", "
                End If
            Next LoopC
            If Len(tStr) > 2 Then
                tStr = Left(tStr, Len(tStr) - 2)
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & "~200~200~200~0~0")
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Número de usuarios: " & N & FONTTYPE_INFO)
            Exit Sub
        'Stand
        Case "/PING"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "BUENO")
            Exit Sub
                Case "/GUERRA"
                
                'Si esta en la carcel no se va
If UserList(UserIndex).Pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If
                
            EntrarGuerra UserIndex
            Exit Sub
        Case "/INICIARGUERRA"
            If UserList(UserIndex).flags.Privilegios <> User Then
                IniciarGuerra UserIndex
            End If
            Exit Sub
        Case "/TERMINARGUERRA"
            If UserList(UserIndex).flags.Privilegios <> User Then
                EmpatarGuerra UserIndex
            End If
            Exit Sub
 
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            'Lorwik - Con esto solucionabamos el problema que si salias montado no podiamos entrar
            'pero como pusimos un nuevo sistema ya no es necesario.
                        'If UserList(UserIndex).flags.Equitando = 1 Then
               ' Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Bajate de la Montura para salir." & FONTTYPE_WARNING)
               ' Exit Sub
           ' End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            Call Cerrar_Usuario(UserIndex)
            Exit Sub
            '[ILUSION]
Case "/DUELO"

'Se asegura que el target es un npc
If UserList(UserIndex).flags.TargetNPC = 0 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If
' Verificamos que sea el npc de duelo
If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 10 Then Exit Sub

' Si esta muy lejos no actua
If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If

' Si esta muerto no puede entrar.
If UserList(UserIndex).flags.Muerto = 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muerto, solo los vivos pueden jugar!!!" & FONTTYPE_VENENO)
Exit Sub
End If

If MapInfo(118).NumUsers = 2 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La sala de duelos está llena." & FONTTYPE_VENENO)
Exit Sub
End If

' Transportamos al usuario
            Call WarpUserChar(UserIndex, 118, 26, 83, True) 'Aca pones el mapa y la posicion.
            UserList(UserIndex).flags.EnDuelo = 1
Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Bienvenido a la sala de duelos." & FONTTYPE_VENENO)
If MapInfo(35).NumUsers = 1 Then
Call SendData(SendTarget.ToAll, 0, 0, "||Duelos> " & UserList(UserIndex).name & " espera contricante en la sala de duelos." & FONTTYPE_TALK)
Else
Call SendData(SendTarget.ToAll, 0, 0, "||Duelos> " & UserList(UserIndex).name & " ha aceptado el duelo." & FONTTYPE_TALK)
End If
'[/ILUSION]
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub
                        
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
            Case "/LIMPIARMUNDO"
        If UserList(UserIndex).flags.Privilegios > 0 Then
       Call SendData(ToAll, 0, 0, "||Servidor> Limpiando Mundo." & FONTTYPE_SERVER)
Dim MapaActual As Integer
MapaActual = 1
For MapaActual = 1 To NumMaps
For Y = YMinMapSize To YMaxMapSize
For X = XMinMapSize To XMaxMapSize
If MapData(MapaActual, X, Y).OBJInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then _

If ItemNoEsDeMapa(MapData(MapaActual, X, Y).OBJInfo.ObjIndex) Then Call EraseObj(ToMap, UserIndex, MapaActual, 10000, MapaActual, X, Y)
                End If
               

Next X
Next Y
Next MapaActual
End If
Exit Sub
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendUserStatsBox(UserIndex) 'para que ande lo de mana
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(val(UserIndex))
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
            Else
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
           UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te estás concentrando. En " & TIEMPO_INICIOMEDITAR & " segundos comenzarás a meditar." & FONTTYPE_INFO)
                
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 45 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARGRANDE
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARXGRANDE
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
            
            
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           Call RevivirUsuario(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If UserList(UserIndex).flags.Envenenado = True Then
           UserList(UserIndex).flags.Envenenado = False
           End If
           
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
           Call SendUserStatsBox(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Notas como las heridas se te van cerrando poco a poco¡¡Hás sido curado!!" & FONTTYPE_INFO)
           Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 26 & "," & 0)
           Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "TW" & 211)
           Exit Sub
           Case "/HOGAR"
           'Si esta en la carcel no se va
If UserList(UserIndex).Pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If
           
 If UserList(UserIndex).Pos.Map = 118 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 132 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 128 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la Mansion Sagrada, mas respeto" & FONTTYPE_INFO)
Exit Sub
End If
          
           
If UserList(UserIndex).flags.Muerto Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido llevado a Ramx" & FONTTYPE_INFO)
Call WarpUserChar(UserIndex, 1, 50, 50, True)
Else
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes estar muerto para utilizar el comando" & FONTTYPE_INFO)
End If
Exit Sub


Case "/CASTILLO SUR"
'Si esta en la carcel no se va
If UserList(UserIndex).Pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 118 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 132 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 128 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la Mansion Sagrada, mas respeto" & FONTTYPE_INFO)
Exit Sub
End If


If UserList(UserIndex).flags.Muerto Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los muertos no pueden participar en la conquista de castillos !!" & FONTTYPE_INFO)
End If

If UserList(UserIndex).Stats.GLD < 100000 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes monedas de oro!." & FONTTYPE_INFO)
Else
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
Call SendUserStatsBox(UserIndex)
End If
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Bienvenido al Castillo sur. Se te restaron 100000 monedas de oro por el viaje." & FONTTYPE_INFO)
Call WarpUserChar(UserIndex, 79, 53, 74, True)

Exit Sub

Case "/CASTILLO NORTE"
'Si esta en la carcel no se va
If UserList(UserIndex).Pos.Map = 127 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 118 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 132 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en Torneo" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Pos.Map = 128 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la Mansion Sagrada, mas respeto" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).flags.Muerto Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Los muertos no pueden participar en la conquista de castillos !!" & FONTTYPE_INFO)
End If
If UserList(UserIndex).Stats.GLD < 100000 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes monedas de oro!." & FONTTYPE_INFO)
Else
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
Call SendUserStatsBox(UserIndex)
End If
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Bienvenido al Castillo norte. Se te restaron 100000 monedas de oro por el viaje." & FONTTYPE_INFO)
Call WarpUserChar(UserIndex, 62, 53, 63, True)

Exit Sub

Case "/CASTILLO"
Call SendData(ToIndex, UserIndex, 0, "||El Castillo Sur esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "SUR") & FONTTYPE_FENIX)
Call SendData(ToIndex, UserIndex, 0, "||El Castillo Norte esta en manos del clan: " & GetVar(App.Path & "\Castillos.ini", "CLANES", "NORTE") & FONTTYPE_FENIX)
Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
        Case "/SEF"
        
           If UserList(UserIndex).flags.Consola = 1 Then
                UserList(UserIndex).flags.Consola = 0
            Else
                UserList(UserIndex).flags.Consola = 1
            End If
           
            Exit Sub
   
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
            '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
Case "/ENCAMAR"
            If UserList(UserIndex).flags.TargetUser > 0 Then
             
              Dim TempIndex As Integer
             
 If UserList(TempIndex).flags.Muerto = 1 Then
                 Call SendData(ToIndex, UserIndex, 0, "||No podes ecamarte con un muerto!!!" & FONTTYPE_INFO)
                 Exit Sub
               End If
 
              If UserList(UserIndex).flags.Muerto = 1 Then
                 Call SendData(ToIndex, UserIndex, 0, "||Primero encamate al cura!!!" & FONTTYPE_INFO)
                 Exit Sub
               End If
 
             
              If TempIndex = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No podes encamarte a vos mismo, de ultima echate una paja!!!" & FONTTYPE_INFO)
                Exit Sub
              End If
             
              'y si no tiene 15 tipo como que la tiene chica
             
               If UserList(UserIndex).Stats.ELV <= 15 Then
                 If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
                   Call SendData(ToIndex, UserIndex, 0, "||La tenes chica todabia, pajeate hasta que tengas 15!!!" & FONTTYPE_INFO)
                   Exit Sub
               
                 Else
                   Call SendData(ToIndex, UserIndex, 0, "||Sos chiquita para tener sexo, cumpli los 15!!!" & FONTTYPE_INFO)
                   Exit Sub
                 End If
              End If
              'onda que garchate una mineeta a menos que seas GAY
              If UserList(UserIndex).Genero = UserList(TempIndex).Genero Then
                 Call SendData(ToIndex, UserIndex, 0, "||Garchate a tu sexo opuesto eehhh!!!" & FONTTYPE_INFO)
                 Exit Sub
              End If
             
              If Distancia(UserList(TempIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                 Call SendData(ToIndex, UserIndex, 0, "||Tenes que estar cerca de la mina para ponerla!!!" & FONTTYPE_INFO)
                 Exit Sub
              End If
             
              UserList(UserIndex).flags.cojiendo = TempIndex
             
              If UserList(TempIndex).flags.cojiendo <> UserIndex And _
                 UserList(TempIndex).flags.cojiendo <> 0 Then
                 Call SendData(ToIndex, UserIndex, 0, "||A la mina se la estan encamando!!!" & FONTTYPE_INFO)
                 Exit Sub
              End If
             
              If UserList(TempIndex).flags.cojiendo = UserIndex Then
                 UserList(TempIndex).flags.tCoje = UserList(UserIndex).name
                 UserList(UserIndex).flags.tCoje = UserList(TempIndex).name
                 UserList(UserIndex).flags.cojiendo = 0
                 UserList(TempIndex).flags.cojiendo = 0
                 Call SendData(ToIndex, UserIndex, 0, "||Te estas cojiendo con" & UserList(TempIndex).name & "!!" & FONTTYPE_INFO)
                 Call SendData(ToIndex, TempIndex, 0, "||Te estas cojiendo con" & UserList(UserIndex).name & "!!" & FONTTYPE_INFO)
                 
                  'Sonidito de putasa
                  If UCase$(UserList(UserIndex).Genero) = "MUJER" Then
                    'la emboca el chabon y grita
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, TempIndex, UserList(TempIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
                    'la minita pega el gritón
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
                    'Le clavamos la panza al toque
                                     
                    UserList(UserIndex).Char.Body = iCuerpoEmbarazada
             
                    Call SendToUserArea(UserIndex, "CP" & UserList(UserIndex).Char.CharIndex & "," & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & NingunArma & "," & NingunEscudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & NingunCasco)
                   
                 ElseIf UCase$(UserList(TempIndex).Genero) = "MUJER" Then
                 
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, TempIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_HOMBRE)
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, e_SoundIndex.MUERTE_MUJER)
                 
                     UserList(TempIndex).Char.Body = iCuerpoEmbarazada
             
                    Call SendToUserArea(TempIndex, "CP" & UserList(TempIndex).Char.CharIndex & "," & UserList(TempIndex).Char.Body & "," & UserList(TempIndex).Char.Head & "," & UserList(TempIndex).Char.Heading & "," & NingunArma & "," & NingunEscudo & "," & UserList(TempIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & NingunCasco)
                 End If
              Else
                Call SendData(ToIndex, TempIndex, 0, "||" & UserList(UserIndex).name & "Quiere hacerte el amor, escribi /encamar y haz click sobre él y pasala a pleno!!" & FONTTYPE_INFO)
                End If
              End If
              Exit Sub
 
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
   
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
    End Select

    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "|+" & UserList(UserIndex).name & "> " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & rData & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        Exit Sub
    End If
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(UserIndex, tInt)
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)
        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
    
    ' ILUSION
    If UCase$(Left$(rData, 8)) = "/DARORO " Then
    rData = Right$(rData, Len(rData) - 8)
    name = ReadField(1, rData, Asc(" "))
    tStr = ReadField(2, rData, Asc(" "))
   
    If name = "" Or tStr = "" Then
        Exit Sub
    End If
   
    If UserList(UserIndex).Stats.GLD < tStr Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes " & tStr & " monedas de oro." & FONTTYPE_INFO)
        Exit Sub
    End If
   
     
    tIndex = NameIndex(name)
           
    If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
   
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - tStr
        Call SendUserStatsBox(val(UserIndex))
        UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + tStr
        Call SendUserStatsBox(val(tIndex))
        Call SendData(ToIndex, UserIndex, 0, "||Le has regalado " & tStr & " monedas de oro a " & UserList(tIndex).name & "." & FONTTYPE_INFO)
        Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).name & " te ha regalado " & tStr & " monedas de oro." & FONTTYPE_VENENO)
    End If
' /ILUSION
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(UserIndex).name, "Mensaje a Gms:" & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).name & "> " & rData & "~0~255~0~1~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|?")
            Exit Sub
         Case "/GX"
            If Not Ayuda.Existe(UserList(UserIndex).name) Then
                Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).name)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
                Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase(Left(rData, 5))
        Case "/BUG "
            rData = Right$(rData, Len(rData) - 5)
            
            Dim CantBugs As Integer
            Dim Bug As Integer
            Dim NuevoBug As String
            Dim Mensaje As String
                
            CantBugs = GetVar(App.Path & "\BUGS\BUG.INI", "BUGS", "CANTIDAD")
                Bug = val(CantBugs) + 1
            NuevoBug = "Bug" & Bug
            Mensaje = UserList(UserIndex).name & " Reporto el siguiente Bug: " & rData

            Call WriteVar(App.Path & "\BUGS\BUG.INI", "Bugs", "Cantidad", Bug)
            Call WriteVar(App.Path & "\BUGS\BUG.INI", "Reportes", NuevoBug, Mensaje)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Bug ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Mensaje & FONTTYPE_TALK)
            
            Exit Sub
            
        Case "/SUG "
            rData = Right$(rData, Len(rData) - 5)
            
            Dim CantSugs As Integer
            Dim Sug As Integer
            Dim NuevaSug As String
            Dim Sugerencia As String
                
            CantSugs = GetVar(App.Path & "\SUGS\SUG.INI", "SUGS", "CANTIDAD")
                Sug = val(CantSugs) + 1
            NuevaSug = "Sug" & Sug
            Sugerencia = UserList(UserIndex).name & " Reporto la siguiente Sugerencia: " & rData

            Call WriteVar(App.Path & "\SUGS\SUG.INI", "Sugs", "Cantidad", Sug)
            Call WriteVar(App.Path & "\SUGS\SUG.INI", "Sugerencias", NuevaSug, Sugerencia)
            
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La Sugerencia ha sido reportado exitosamente!" & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & Sugerencia & FONTTYPE_TALK)
            
            Exit Sub
            
               Case "/CARA"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
        ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
            'Se asegura que el target es un npc
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
        ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
        ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Cirujano Then
        Exit Sub
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Espero que te guste tu nueva cara!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            If UserList(UserIndex).Raza = "Humano" Then
                If UserList(UserIndex).Genero = "Hombre" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 30)
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 30)
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
               
                If UserList(UserIndex).Genero = "Mujer" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 7) + 69
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 7) + 69
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
            ElseIf UserList(UserIndex).Raza = "Elfo" Then
                If UserList(UserIndex).Genero = "Hombre" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 13) + 100
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 13) + 100
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
                If UserList(UserIndex).Genero = "Mujer" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 7) + 169
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 7) + 169
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
            ElseIf UserList(UserIndex).Raza = "Elfo oscuro" Then
                If UserList(UserIndex).Genero = "Hombre" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 8) + 201
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 8) + 201
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
                If UserList(UserIndex).Genero = "Mujer" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 11) + 269
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 11) + 269
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
            ElseIf UserList(UserIndex).Raza = "Enano" Then
                    If UserList(UserIndex).Genero = "Hombre" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 5) + 300
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 5) + 300
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
                If UserList(UserIndex).Genero = "Mujer" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 3) + 369
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 3) + 369
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
            ElseIf UserList(UserIndex).Raza = "Gnomo" Then
            If UserList(UserIndex).Genero = "Hombre" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 6) + 400
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 6) + 400
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
           
                If UserList(UserIndex).Genero = "Mujer" Then
                    UserList(UserIndex).Char.Head = RandomNumber(1, 5) + 469
                    UserList(UserIndex).OrigChar.Head = RandomNumber(1, 5) + 469
                    Call WriteVar(CharPath & UCase(UserList(UserIndex).name) & ".chr", "INIT", "Head", str(UserList(UserIndex).OrigChar.Head))
                    Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
                End If
                End If
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
         Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            ' [GS] Seguro anti /passwd - v2 (idea de Aereal)
            ' /passwd passnuevo passviejo
            Arg1 = ReadField(1, rData, 32)
            Arg2 = ReadField(2, rData, 32)
            If Arg2 <> UserList(UserIndex).Password Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No se ha podido completar el cambio de la contraseña. Utilize /passwd <password_nuevo> <password_anterior>" & FONTTYPE_INFO)
            ElseIf Len(Arg1) < 6 Then
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(UserIndex).Password = Arg1
            End If
            ' [/GS] Seguro anti /passwd - v2 (idea de Aereal)
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(UserIndex)
            End If
            Exit Sub
    
                
    End Select
    
    Select Case UCase$(Left$(rData, 10))

        'Standelf - Advertencias
        Case "/VERADVER "
       
            Dim Adverts As Integer
            Dim Advert As String
            rData = UCase$(Right$(rData, Len(rData) - 10))
            tStr = Replace$(ReadField(1, rData, 32), "+", " ") 'Nick
 
            If UserList(UserIndex).flags.Privilegios = User And UserList(UserIndex).name <> tStr Then Exit Sub
           
            Adverts = val(GetVar(CharPath & tStr & ".chr", "Advertencias", "Number"))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario " & tStr & " tiene " & Adverts & FONTTYPE_INFO)
           
            Dim loopX As Integer
                For loopX = 1 To Adverts
                   
                        Advert = GetVar(CharPath & tStr & ".chr", "Advertencias", "Adv" & loopX)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Advertencia" & loopX & " : " & Advert & FONTTYPE_INFO)
                Next loopX
                   
            Exit Sub
            
        Case "/ADVERTIR "
            If UserList(UserIndex).flags.Privilegios = User Then Exit Sub
            Dim TotalAdvert As Integer
            rData = UCase$(Right$(rData, Len(rData) - 10))
            tStr = Replace$(ReadField(1, rData, 32), "+", " ") 'Nick
                tIndex = NameIndex(tStr)
                    Arg1 = ReadField(2, rData, 32)
                   
            TotalAdvert = val(GetVar(CharPath & tStr & ".chr", "Advertencias", "Number"))
            TotalAdvert = val(TotalAdvert) + 1
            Call WriteVar(CharPath & tStr & ".chr", "Advertencias", "Number", val(TotalAdvert))
 
            Call WriteVar(CharPath & tStr & ".chr", "Advertencias", "Adv" & TotalAdvert, Arg1)
            'Call WriteVar(CharPath & tStr & ".bwpj", "Advertencias", "Adv" & TotalAdvert, Arg1)
           
            'Notificamos A los usuarios Que el GM advirtio a un usuario
            Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(UserIndex).name & " advirtio a: " & tStr & FONTTYPE_ADVERTENCIAS)
 
 
            'Notificamos al usuarios que Fue Advertido, el motivo, quien lo advirtio y la cantidad de advertencias que tiene
             If tIndex <= 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Personaje esta Offline." & FONTTYPE_ADVERTENCIAS)
                Exit Sub
            Else
                Call SendData(SendTarget.ToIndex, tIndex, 0, "||Has sido Advertido por: " & UserList(UserIndex).name & ". El Motivo de la Advertencias es: " & Arg1 & " .Con esta llevas " & TotalAdvert & FONTTYPE_ADVERTENCIAS)
            End If
 
            'Encarcelamos el total de advertencias x 5 Ejemplo 2 Advertencias, lo encarcela por 10 Minutos.
            If Not val(TotalAdvert) >= 5 Then
                Call Encarcelar(tIndex, TotalAdvert * 5, UserList(UserIndex).name)
            End If
 
            'Si llego al Maximo de Advertencias?
            If val(TotalAdvert) >= 5 Then
                Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & tStr & " ha sido Baneado Automaticamente por llegar a su Maximo de advertencias." & FONTTYPE_ADVERTENCIAS)
                    tInt = val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
                    Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, "El Servidor te ha Baneado Automaticamente. El Motivo es: Acumulacion de Advertencias. " & Date & " " & Time)
                   
                'Desconectamos al usuario
                If Not tIndex <= 0 Then Call CloseSocket(tIndex)
               
                'Baneamos ^^
                Call WriteVar(CharPath & tStr & ".chr", "FLAGS", "Ban", "1")
            End If
           
            Exit Sub
            
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rData))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Sub
        'Standelf
        Case "/DENUNCIAR "
                If Denuncias = False Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Las denuncias no estan activadas." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub
            
        Case "/FUNDARCLAN"
        
        If Not TieneObjetos(1212, 1, UserIndex) Then
Call SendData(ToIndex, UserIndex, 0, "||Necesitas tener el Anillo del Clan. Puedes solicitarla en el Foro." & FONTTYPE_GUILD)
Exit Sub
End If
        
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHOWFUN")
            Else
                UserList(UserIndex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
