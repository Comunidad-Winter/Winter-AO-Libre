Attribute VB_Name = "modHechizos"
Option Explicit

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userindex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim da�o As Integer

If Hechizos(Spell).SubeHP = 1 Then

    da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + da�o
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    Call SendUserStatsBox(val(userindex))

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(userindex).flags.Privilegios = PlayerType.User Then
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
            da�o = da�o - RandomNumber(ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        If da�o < 0 Then da�o = 0
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - da�o
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
        Call SendUserStatsBox(val(userindex))
        
        'Muere
        If UserList(userindex).Stats.MinHP < 1 Then
            UserList(userindex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (userindex)
            End If
             

 
            Call UserDie(userindex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
            
                         If userindex = GranPoder Then
                    Call SendData(SendTarget.toall, 0, 0, "PRE8," & UserList(userindex).name)
                    Call OtorgarGranPoder(0)
                End If
            
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(userindex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(userindex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(userindex).flags.Paralizado = 1
          UserList(userindex).Counters.Paralisis = IntervaloParalizado

#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.ToIndex, userindex, 0, "PARADOK")
        Else
#End If
            Call SendData(SendTarget.ToIndex, userindex, 0, "PARADOK")
#If SeguridadAlkon Then
        End If
#End If
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim da�o As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "CFX" & Npclist(TargetNPC).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - da�o
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal userindex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
        Call SonidosMapas.ReproducirSonido(SendTarget.ToIndex, userindex, UserList(userindex).pos.Map, 158)
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal userindex As Integer)
On Error Resume Next
 
    Dim ind As String
        ind = UserList(userindex).Char.CharIndex
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &H8080& & "�" & S & "�" & ind)
    Exit Sub
   
End Sub

Function PuedeLanzar(ByVal userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(userindex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(userindex).flags.TargetMap
    wp2.x = UserList(userindex).flags.TargetX
    wp2.Y = UserList(userindex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu B�culo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes lanzar este conjuro sin la ayuda de un b�culo." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(userindex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Est�s muy cansado para lanzar este hechizo." & FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
           Call SendData(SendTarget.ToIndex, userindex, 0, "PRB22")
            PuedeLanzar = False
        End If
    Else
           Call SendData(SendTarget.ToIndex, userindex, 0, "PRB23")
            Call SonidosMapas.ReproducirSonido(SendTarget.ToIndex, userindex, UserList(userindex).pos.Map, RandomNumber(189, 190))
            PuedeLanzar = False
    End If
Else
   Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal userindex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(userindex).flags.TargetX
    PosCasteadaY = UserList(userindex).flags.TargetY
    PosCasteadaM = UserList(userindex).flags.TargetMap
    
    H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).userindex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(userindex)
    End If

End Sub

Sub HechizoInvocacion(ByVal userindex As Integer, ByRef b As Boolean)

If UserList(userindex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(userindex).pos.Map).Pk = False Or MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
       Call SendData(SendTarget.ToIndex, userindex, 0, "PRE34")
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(userindex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
            
            Index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(Index) = ind
            UserList(userindex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(userindex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(userindex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(userindex, b)
        
        Case uMaterializa
        Call HechizoMaterializar(userindex, b)
    
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(userindex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
    Call SendUserStatsBox(UserList(userindex).flags.TargetUser)
    UserList(userindex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userindex).flags.TargetNPC, uh, b, userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNPC, userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNPC = 0
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
End If

End Sub


Sub LanzarHechizo(Index As Integer, userindex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(userindex).Stats.UserHechizos(Index)

If PuedeLanzar(userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case TargetType.uNPC
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(userindex, uh)
    End Select
    
End If

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal userindex As Integer, ByRef b As Boolean)



Dim H As Integer, TU As Integer
H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
TU = UserList(userindex).flags.TargetUser


If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||�Est� muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Criminal(TU) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call VolverCriminal(userindex)
        End If
    End If
    
    UserList(TU).flags.Invisible = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call InfoHechizo(userindex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(userindex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Privilegios >= PlayerType.Consejero Then
        Exit Sub
    End If
    
    If UserList(userindex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(userindex)
        .CharMimetizado.Body = .Char.Body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.Body = UserList(TU).Char.Body
        .Char.Head = UserList(TU).Char.Head
        .Char.CascoAnim = UserList(TU).Char.CascoAnim
        .Char.ShieldAnim = UserList(TU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(TU).Char.WeaponAnim
    
        Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(userindex)
   b = True
End If


If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).ArmorBreak = 1 Then
        If UserList(TU).flags.Desnudo = 1 Then
            Call SendData(ToIndex, userindex, 0, "||�El usuario ya est� desnudo!." & FONTTYPE_INFO)
        Exit Sub
        End If
       
        Call Desequipar(userindex, UserList(TU).Invent.ArmourEqpSlot)
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then

If userindex = TU Then
 Call SendData(SendTarget.ToIndex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
Exit Sub
End If
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(userindex, TU) Then Exit Sub
            
            If userindex <> TU Then
                Call UsuarioAtacadoPorUsuario(userindex, TU)
            End If
            
            Call InfoHechizo(userindex)
            b = True
            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.ToIndex, TU, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.ToIndex, userindex, 0, "|| �El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                Call SendCryptedData(SendTarget.ToIndex, TU, 0, "PARADOK")
            Else
#End If
                Call SendData(SendTarget.ToIndex, TU, 0, "PARADOK")
#If SeguridadAlkon Then
            End If
#End If
            
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
        If Criminal(TU) And Not Criminal(userindex) Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(userindex)
            End If
        End If
       
        UserList(TU).flags.Paralizado = 0
         UserList(TU).Counters.Paralisis = 0
        Call SendData(SendTarget.ToIndex, TU, 0, "INMO0")
        Call SendData(SendTarget.ToIndex, TU, 0, "||Has recuperado la movilidad." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToIndex, TU, 0, "PARADOK")
        Call SendData(SendTarget.ToIndex, userindex, 0, "PU" & UserList(userindex).pos.x & "," & UserList(userindex).pos.Y)
        Call InfoHechizo(userindex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If Not UserList(TU).flags.Estupidez = 0 Then
                UserList(TU).flags.Estupidez = 0
                'no need to crypt this
                Call SendData(SendTarget.ToIndex, TU, 0, "NESTUP")
                Call InfoHechizo(userindex)
                b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        If Criminal(TU) And Not Criminal(userindex) Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
                Exit Sub
            Else
                Call VolverCriminal(userindex)
            End If
        End If

        'revisamos si necesita vara
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
            
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Necesitas un mejor b�culo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UCase$(UserList(userindex).Clase) = "BARDO" Then
            If UserList(userindex).Invent.HerramientaEqpObjIndex <> LAUDMAGICO Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Necesitas un instrumento m�gico para devolver la vida" & FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        'Pablo Toxic Waste
        UserList(TU).Stats.MinAGU = UserList(TU).Stats.MinAGU - 25
        UserList(TU).Stats.MinHam = UserList(TU).Stats.MinHam - 25
        'Juan Maraxus
        If UserList(TU).Stats.MinAGU <= 0 Then
                UserList(TU).Stats.MinAGU = 0
                UserList(TU).flags.Sed = 1
        End If
        If UserList(TU).Stats.MinHam <= 0 Then
                UserList(TU).Stats.MinHam = 0
                UserList(TU).flags.Hambre = 1
        End If
        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> userindex Then
                UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep + 500
                If UserList(userindex).Reputacion.NobleRep > MAXREP Then _
                    UserList(userindex).Reputacion.NobleRep = MAXREP
               Call SendData(SendTarget.ToIndex, userindex, 0, "PRB26")
            End If
        End If
        UserList(TU).Stats.MinMAN = 0
        Call EnviarHambreYsed(TU)
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(userindex)
        Call RevivirUsuario(TU)
    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3
#If SeguridadAlkon Then
        Call SendCryptedData(SendTarget.ToIndex, TU, 0, "CEGU")
#Else
        Call SendData(SendTarget.ToIndex, TU, 0, "CEGU")
#End If
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.ToIndex, TU, 0, "DUMB")
        Else
#End If
            Call SendData(SendTarget.ToIndex, TU, 0, "DUMB")
#If SeguridadAlkon Then
        End If
#End If
        Call InfoHechizo(userindex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userindex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
            Exit Sub
        Else
            UserList(userindex).Reputacion.NobleRep = 0
            UserList(userindex).Reputacion.PlebeRep = 0
            UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 200
            If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
        
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
            Exit Sub
        Else
            UserList(userindex).Reputacion.NobleRep = 0
            UserList(userindex).Reputacion.PlebeRep = 0
            UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 200
            If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                UserList(userindex).Reputacion.AsesinoRep = MAXREP
        End If
    End If
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
                Exit Sub
            Else
                UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Call InfoHechizo(userindex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = userindex Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(SendTarget.ToIndex, userindex, 0, "||Este hechizo solo afecta NPCs que tengan amo." & FONTTYPE_WARNING)
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes quitarte el seguro para de poder atacar guardias" & FONTTYPE_WARNING)
                Exit Sub
            Else
                UserList(userindex).Reputacion.NobleRep = 0
                UserList(userindex).Reputacion.PlebeRep = 0
                UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
                If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(userindex).Reputacion.AsesinoRep = MAXREP
            End If
        End If
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(userindex)
        b = True
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userindex As Integer, ByRef b As Boolean)

Dim da�o As Long


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(userindex).Stats.ELV)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "�" & da�o & "�" & str(UserList(userindex).Char.CharIndex))
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + da�o
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Has curado " & da�o & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Npclist(NpcIndex).NPCtype = 2 And UserList(userindex).flags.Seguro Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Debes sacarte el seguro para atacar guardias del imperio." & FONTTYPE_FIGHT)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(userindex).Stats.ELV)

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
             
                da�o = (da�o * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta da�o segun el staff-
                'Da�o = (Da�o* (80 + BonifB�culo)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 80% del original
            End If
        End If
    End If
    If UserList(userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If


    Call InfoHechizo(userindex)
    b = True
    Call NpcAtacado(NpcIndex, userindex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
    SendData SendTarget.ToIndex, userindex, 0, "PRB27," & da�o
    Call CalcularDarExp(userindex, NpcIndex, da�o)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userindex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal userindex As Integer)


    Dim H As Integer
    H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, userindex)
    
    If UserList(userindex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(UserList(userindex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, UserList(userindex).pos.Map, "TW" & Hechizos(H).WAV)
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).pos.Map, "CFX" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, UserList(userindex).pos.Map, "TW" & Hechizos(H).WAV)
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
        If userindex <> UserList(userindex).flags.TargetUser Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.ToIndex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal userindex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim da�o As Integer
Dim tempChr As Integer
    
    
H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tempChr = UserList(userindex).flags.TargetUser
      
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| le tiro el hechizo " & H & " a " & UserList(tempChr).Name & FONTTYPE_VENENO)
'End If
      
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(userindex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + da�o
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has restaurado " & da�o & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(userindex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - da�o
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has quitado " & da�o & " puntos de hambre a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + da�o
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If userindex <> tempChr Then
      Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has restaurado " & da�o & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
      Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - da�o
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has quitado " & da�o & " puntos de sed a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(userindex)
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    Call InfoHechizo(userindex)
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    
    If Criminal(tempChr) And Not Criminal(userindex) Then
        If UserList(userindex).flags.Seguro Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos" & FONTTYPE_INFO)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    End If
    
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(userindex).Stats.ELV)
    
    Call InfoHechizo(userindex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + da�o
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If userindex = tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| danio, minhp, maxhp " & da�o & " " & Hechizos(H).MinHP & " " & Hechizos(H).MaxHP & FONTTYPE_VENENO)
'End If
    
    
    da�o = da�o + Porcentaje(da�o, 3 * UserList(userindex).Stats.ELV)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| da�o, ELV " & da�o & " " & UserList(UserIndex).Stats.ELV & FONTTYPE_VENENO)
'End If
    
    
    If Hechizos(H).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                da�o = (da�o * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                da�o = da�o * 0.7 'Baja da�o a 70% del original
            End If
        End If
    End If
    
    If UserList(userindex).Invent.HerramientaEqpObjIndex = LAUDMAGICO Then
        da�o = da�o * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If da�o < 0 Then da�o = 0
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - da�o
    
    Call SendData(SendTarget.ToIndex, userindex, 0, "PRB28," & da�o & "," & UserList(tempChr).name)
    Call SendData(SendTarget.ToIndex, tempChr, 0, "PRB29," & UserList(userindex).name & "," & da�o)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userindex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + da�o
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has restaurado " & da�o & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has quitado " & da�o & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - da�o
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + da�o
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has restaurado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Le has quitado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToIndex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - da�o
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.ToIndex, userindex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(SendTarget.ToIndex, userindex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, userindex, CualHechizo)

End Sub


Public Sub DisNobAuBan(ByVal userindex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'Si estamos en la arena no hacemos nada
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.x, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep - NoblePts
    If UserList(userindex).Reputacion.NobleRep < 0 Then
        UserList(userindex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + BandidoPts
    If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(userindex).Reputacion.BandidoRep = MAXREP
    Call SendData(SendTarget.ToIndex, userindex, 0, "PN")
    If Criminal(userindex) Then If UserList(userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)
End Sub
Sub HechizoMaterializar(userindex As Integer, b As Boolean)
 
Dim TU As Integer
Dim H As Integer
Dim i As Integer
 
Dim PosTIROTELEPORT As WorldPos 'matute
 
H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
 
 
If Hechizos(H).Materializa = 1 Then 'matute

    If UserList(userindex).pos.Map = 127 Then Exit Sub 'zona espera
    If UserList(userindex).pos.Map = 118 Then Exit Sub 'zonas de "torneo"
    If UserList(userindex).pos.Map = 125 Then Exit Sub 'Carcel
    'If UserList(userindex).Counters.TimeTeleport <> 0 Then Exit Sub 'Ya invocó.
   
    If UserList(userindex).flags.TiroPortalL = True Then
        Call SendData(ToIndex, userindex, 0, "||��Ya tienes un portal oscuro invocado.!!" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    PosTIROTELEPORT.x = UserList(userindex).flags.TargetX
    PosTIROTELEPORT.Y = UserList(userindex).flags.TargetY
    PosTIROTELEPORT.Map = UserList(userindex).flags.TargetMap
   
    UserList(userindex).flags.DondeTiroMap = PosTIROTELEPORT.Map
    UserList(userindex).flags.DondeTiroX = PosTIROTELEPORT.x
    UserList(userindex).flags.DondeTiroY = PosTIROTELEPORT.Y
   
    If MapData(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).OBJInfo.ObjIndex Then 'si hay algo...
        Exit Sub
    End If
   
    If MapData(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Map Then
        Exit Sub
    End If
   
    If MapData(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked Then
        Exit Sub
    End If
    
        If MapData(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).userindex Then
        Exit Sub
    End If
   
    If Not MapaValido(UserList(userindex).pos.Map) Or Not InMapBounds(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then Exit Sub
   
   
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 1269 'veamos asd - [Primer FX que se ve en la imagen 1] -VER OBJ.DAT
   
   
    Call MakeObj(ToMap, userindex, UserList(userindex).pos.Map, ET, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY)
    b = True
                       
    UserList(userindex).Counters.TimeTeleport = 0
    UserList(userindex).Counters.CreoTeleport = True
    UserList(userindex).flags.TiroPortalL = True
End If
 
Call InfoHechizo(userindex)
 
End Sub
