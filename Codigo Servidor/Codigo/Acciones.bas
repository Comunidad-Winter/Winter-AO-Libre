Attribute VB_Name = "Acciones"
Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Error Resume Next

'�Posicion valida?
If InMapBounds(Map, X, Y) Then
   
    Dim FoundChar As Byte
    Dim FoundSomething As Byte
    Dim TempCharIndex As Integer
       
    '�Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
        
        Select Case ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType
            
            Case eOBJType.otPuertas 'Es una puerta
                Call AccionParaPuerta(Map, X, Y, UserIndex)
            Case eOBJType.otCarteles 'Es un cartel
                Call AccionParaCartel(Map, X, Y, UserIndex)
            Case eOBJType.otForos 'Foro
                Call AccionParaForo(Map, X, Y, UserIndex)
            Case eOBJType.otLe�a    'Le�a
                If MapData(Map, X, Y).OBJInfo.ObjIndex = FOGATA_APAG And UserList(UserIndex).flags.Muerto = 0 Then
                    Call AccionParaRamita(Map, X, Y, UserIndex)
                End If
        End Select
    '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
            
        End Select
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.ObjIndex
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).name & "," & "OBJ")
        Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType
            
            Case 6 'Es una puerta
                Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
            
        End Select
    ElseIf MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
        'Set the target NPC
        UserList(UserIndex).flags.TargetNPC = MapData(Map, X, Y).NpcIndex
        
        If Npclist(MapData(Map, X, Y).NpcIndex).Comercia = 1 Then
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Iniciamos la rutina pa' comerciar.
            Call IniciarCOmercioNPC(UserIndex)
        
        ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Banquero Then
            If Distancia(Npclist(MapData(Map, X, Y).NpcIndex).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Standelf
            Call SendUserStatsBox(UserIndex)
                SendData SendTarget.ToIndex, UserIndex, 0, "INITBANKO"
        
        ElseIf Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = eNPCType.Revividor Then
            If Distancia(UserList(UserIndex).Pos, Npclist(MapData(Map, X, Y).NpcIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
                Exit Sub
            End If
           
           'Revivimos si es necesario
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call RevivirUsuario(UserIndex)
            End If
            
            'curamos totalmente
            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
            Call SendUserStatsBox(UserIndex)
        End If
    Else
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
    End If
End If

End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
    Exit Sub
End If

'�Hay mensajes?
Dim f As String, tit As String, men As String, base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = base & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = ""
        auxcad = ""
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "FMSG" & tit & Chr(176) & men)
        
    Next
End If
Call SendData(SendTarget.ToIndex, UserIndex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexAbierta
                    
                    Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                     
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, X, Y, 0)
                    Call Bloquear(SendTarget.ToMap, 0, Map, Map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).OBJInfo.ObjIndex = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).IndexCerrada
                
                Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhIndex & "," & X & "," & Y)
                
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(SendTarget.ToMap, 0, Map, Map, X, Y, 1)
                
                SendData SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PUERTA
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.ObjIndex
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La puerta esta cerrada con llave." & FONTTYPE_INFO)
    End If
Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto) > 0 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "MCAR" & _
        ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).texto & _
        Chr(176) & ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).GrhSecundario)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||En zona segura no puedes hacer fogatas." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(UserIndex).Pos.Map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.Amount = 1
        
        Call SendData(ToIndex, UserIndex, 0, "||Has prendido la fogata." & FONTTYPE_INFO)
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "FO")
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||La ley impide realizar fogatas en las ciudades." & FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has podido hacer fuego." & FONTTYPE_INFO)
End If

'Sino tiene hambre o sed quizas suba el skill supervivencia
If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
    Call SubirSkill(UserIndex, Supervivencia)
End If

End Sub
Sub OtorgarGranPoder(UserIndex As Integer)
Dim LoopC As Integer
Dim EncontroIdeal As Boolean
If LastUser = 0 Then Exit Sub
If UserIndex = 0 Then
    Do While EncontroIdeal = False And LoopC < 500
        LoopC = LoopC + 1
        UserIndex = RandomNumber(1, LastUser)
        If UserList(UserIndex).flags.UserLogged = True And UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.Privilegios = User Then
            EncontroIdeal = True
            Exit Do
        End If
    Loop
    If Not EncontroIdeal Then
        UserIndex = 0
        GranPoder = 0
    End If
End If
If UserIndex > 0 Then
    If UserList(UserIndex).flags.Muerto <> 0 Then Call OtorgarGranPoder(0)
    GranPoder = UserIndex
    Call SendData(SendTarget.ToAll, UserIndex, 0, "||Los dioses le otorgan el Gran Poder a " & UserList(UserIndex).name & " en el mapa " & UserList(UserIndex).Pos.Map & "." & FONTTYPE_GUILD)
    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
End Sub
