Attribute VB_Name = "ModFacciones"
Option Explicit

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarqu�a
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer

Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/03/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineaci�n neutral.
'Handles the entrance of users to the "Caballeros del Temple"
'***************************************************
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "���Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "���Maldito insolente!!! vete de aqui seguidor de las sombras", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "���No se permiten criminales en el ej�rcito imperial!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 30 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "���Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call WriteChatOverHead(UserIndex, "�Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call WriteChatOverHead(UserIndex, "�Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Reputacion.NobleRep < 1000000 Then
    Call WriteChatOverHead(UserIndex, "Necesitas ser a�n m�s Noble para integrar el Ej�rcito del Rey, solo tienes " & UserList(UserIndex).Reputacion.NobleRep & "/1.000.000 Puntos de Nobleza", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .Name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "���Eres el fundador de un clan neutro!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1

Call WriteChatOverHead(UserIndex, "���Bienvenido al Ej�rcito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargar� de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.amount = 20
    MiObj2.amount = 10
        
    If UserList(UserIndex).raza = eRaza.Enano Or UserList(UserIndex).raza = eRaza.Gnomo Then
        MiObj.ObjIndex = VestimentaImperialEnano
        Select Case UserList(UserIndex).clase
            Case eClass.Mage
                MiObj2.ObjIndex = TunicaConspicuaEnano
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaImperialHumano
        Select Case UserList(UserIndex).clase
            Case eClass.Mage
                MiObj2.ObjIndex = TunicaConspicuaHumano
            Case eClass.Cleric, eClass.Druid, eClass.Bard
                MiObj2.ObjIndex = ArmaduraGranSacerdote
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaHumano
        End Select
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj2) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj2)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date
    'Esto por ahora es in�til, siempre va a ser cero, pero bueno, despues va a servir.
    UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.CiudadanosMatados

End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AsignarExperiencia(UserIndex, ExpAlUnirse)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    UserList(UserIndex).Faccion.RecompensasReal = 0
    UserList(UserIndex).Faccion.NextRecompensa = 70
End If

'Agregado para que no hayan armadas en un clan Neutro
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan por tu nueva facci�n.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

Call LogEjercitoReal(UserList(UserIndex).Name & " ingres� el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Caballeros del Temple"
'***************************************************
Dim Crimis As Long
Dim Lvl As Byte
Dim NextRecom As Long
Dim Nobleza As Long
Lvl = UserList(UserIndex).Stats.ELV
Crimis = UserList(UserIndex).Faccion.CriminalesMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa
Nobleza = UserList(UserIndex).Reputacion.NobleRep

If Crimis < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " Criminales m�s para recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 70:
        UserList(UserIndex).Faccion.RecompensasReal = 1
        UserList(UserIndex).Faccion.NextRecompensa = 130
    
    Case 130:
        UserList(UserIndex).Faccion.RecompensasReal = 2
        UserList(UserIndex).Faccion.NextRecompensa = 210
    
    Case 210:
        UserList(UserIndex).Faccion.RecompensasReal = 3
        UserList(UserIndex).Faccion.NextRecompensa = 320
    
    Case 320:
        UserList(UserIndex).Faccion.RecompensasReal = 4
        UserList(UserIndex).Faccion.NextRecompensa = 460
    
    Case 460:
        UserList(UserIndex).Faccion.RecompensasReal = 5
        UserList(UserIndex).Faccion.NextRecompensa = 640
    
    Case 640:
        If Lvl < 27 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 27 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 6
        UserList(UserIndex).Faccion.NextRecompensa = 870
    
    Case 870:
        UserList(UserIndex).Faccion.RecompensasReal = 7
        UserList(UserIndex).Faccion.NextRecompensa = 1160
    
    Case 1160:
        UserList(UserIndex).Faccion.RecompensasReal = 8
        UserList(UserIndex).Faccion.NextRecompensa = 2000
    
    Case 2000:
        If Lvl < 30 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 30 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 9
        UserList(UserIndex).Faccion.NextRecompensa = 2500
    
    Case 2500:
        If Nobleza < 2000000 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 2000000 - Nobleza & " puntos de Nobleza para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 10
        UserList(UserIndex).Faccion.NextRecompensa = 3000
    
    Case 3000:
        If Nobleza < 3000000 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 3000000 - Nobleza & " puntos de Nobleza para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 11
        UserList(UserIndex).Faccion.NextRecompensa = 3500
    
    Case 3500:
        If Lvl < 35 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 35 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If Nobleza < 4000000 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 4000000 - Nobleza & " puntos de Nobleza para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 12
        UserList(UserIndex).Faccion.NextRecompensa = 4000
    
    Case 4000:
        If Lvl < 36 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 36 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If Nobleza < 5000000 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 5000000 - Nobleza & " puntos de Nobleza para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 13
        UserList(UserIndex).Faccion.NextRecompensa = 5000
    
    Case 5000:
        If Lvl < 37 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 37 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If Nobleza < 6000000 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & 6000000 - Nobleza & " puntos de Nobleza para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasReal = 14
        UserList(UserIndex).Faccion.NextRecompensa = 10000
    
    Case 10000:
        Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Crimis & ", sigue asi. Ya no tengo m�s recompensa para darte que mi agradescimiento. �Felicidades!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    
    Case Else:
        Exit Sub
End Select

Call WriteChatOverHead(UserIndex, "���Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
Call AsignarExperiencia(UserIndex, ExpX100)


End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "���Has sido expulsado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "���Te has retirado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si est� equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si est� equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "���Has sido expulsado de la Legi�n Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "���Te has retirado de la Legi�n Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura de caos si est� equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si est� equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Caballeros del Temple"
'***************************************************
Select Case UserList(UserIndex).Faccion.RecompensasReal
'Rango 1: Aprendiz (30 Criminales)
'Rango 2: Escudero (70 Criminales)
'Rango 3: Soldado (130 Criminales)
'Rango 4: Sargento (210 Criminales)
'Rango 5: Caballero (320 Criminales)
'Rango 6: Comandante (460 Criminales)
'Rango 7: Capit�n (640 Criminales + > lvl 27)
'Rango 8: Senescal (870 Criminales)
'Rango 9: Mariscal (1160 Criminales)
'Rango 10: Condestable (2000 Criminales + > lvl 30)
'Rangos de Honor de los Caballeros del Temple: (Consejo de Bander)
'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
'Rango 14: Guardi�n del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
'Rango 15: Campe�n de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
    Case 0
        TituloReal = "Aprendiz"
    Case 1
        TituloReal = "Escudero"
    Case 2
        TituloReal = "Soldado"
    Case 3
        TituloReal = "Sargento"
    Case 4
        TituloReal = "Teniente"
    Case 5
        TituloReal = "Comandante"
    Case 6
        TituloReal = "Capit�n"
    Case 7
        TituloReal = "Senescal"
    Case 8
        TituloReal = "Mariscal"
    Case 9
        TituloReal = "Condestable"
    Case 10
        TituloReal = "Ejecutor Imperial"
    Case 11
        TituloReal = "Protector del Reino"
    Case 12
        TituloReal = "Avatar de la Justicia"
    Case 13
        TituloReal = "Guardi�n del Bien"
    Case Else
        TituloReal = "Campe�n de la Luz"
End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 15/3/2009
'15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineaci�n neutral.
'Handles the entrance of users to the "Legi�n Oscura"
'***************************************************
If Not criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "���L�rgate de aqui, buf�n!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "���Ya perteneces a la legi�n oscura!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "Las sombras reinar�n en Argentum. ���Fuera de aqui insecto Real!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

'[Barrin 17-12-03] Si era miembro de los Caballeros del Temple no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ah�: �Recibio la experiencia para entrar?
    Call WriteChatOverHead(UserIndex, "No permitir� que ning�n insecto real ingrese a mis tropas.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
'[/Barrin]

If Not criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "��Ja ja ja!! Tu no eres bienvenido aqui asqueroso Ciudadano", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 70 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 70 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "���Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .Name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "���Eres el fundador de un clan neutro!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With


If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
        Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeld�a has atacado a mi ej�rcito. �Vete de aqu�!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "�Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    End If
    Exit Sub
End If

UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.FuerzasCaos = 1

Call WriteChatOverHead(UserIndex, "���Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre Ciudadana y Real y ser�s recompensado, lo prometo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.amount = 20
    MiObj2.amount = 10
    
    If UserList(UserIndex).raza = eRaza.Enano Or UserList(UserIndex).raza = eRaza.Gnomo Then
        MiObj.ObjIndex = VestimentaLegionEnano
        Select Case UserList(UserIndex).clase
            Case eClass.Mage
                MiObj2.ObjIndex = TunicaEgregiaEnano
            Case Else
                MiObj2.ObjIndex = TunicaLobregaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaLegionHumano
        Select Case UserList(UserIndex).clase
            Case eClass.Mage
                MiObj2.ObjIndex = TunicaEgregiaHumano
            Case eClass.Cleric, eClass.Druid, eClass.Bard
                MiObj2.ObjIndex = SacerdoteDemoniaco
            Case Else
                MiObj2.ObjIndex = TunicaLobregaHumano
        End Select
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj2) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj2)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
    UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
    UserList(UserIndex).Faccion.FechaIngreso = Date

End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AsignarExperiencia(UserIndex, ExpAlUnirse)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    UserList(UserIndex).Faccion.RecompensasCaos = 0
    UserList(UserIndex).Faccion.NextRecompensa = 160
End If

'Agregado para que no hayan armadas en un clan Neutro
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan por tu nueva facci�n.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

Call LogEjercitoCaos(UserList(UserIndex).Name & " ingres� el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Legi�n Oscura"
'***************************************************
Dim Ciudas As Long
Dim Lvl As Byte
Dim NextRecom As Long
Lvl = UserList(UserIndex).Stats.ELV
Ciudas = UserList(UserIndex).Faccion.CiudadanosMatados
NextRecom = UserList(UserIndex).Faccion.NextRecompensa

If Ciudas < NextRecom Then
    Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " Cuidadanos m�s para recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

Select Case NextRecom
    Case 160:
        UserList(UserIndex).Faccion.RecompensasCaos = 1
        UserList(UserIndex).Faccion.NextRecompensa = 300
    
    Case 300:
        UserList(UserIndex).Faccion.RecompensasCaos = 2
        UserList(UserIndex).Faccion.NextRecompensa = 490
    
    Case 490:
        UserList(UserIndex).Faccion.RecompensasCaos = 3
        UserList(UserIndex).Faccion.NextRecompensa = 740
    
    Case 740:
        UserList(UserIndex).Faccion.RecompensasCaos = 4
        UserList(UserIndex).Faccion.NextRecompensa = 1100
    
    Case 1100:
        UserList(UserIndex).Faccion.RecompensasCaos = 5
        UserList(UserIndex).Faccion.NextRecompensa = 1500
    
    Case 1500:
        If Lvl < 27 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 27 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 6
        UserList(UserIndex).Faccion.NextRecompensa = 2010
    
    Case 2010:
        UserList(UserIndex).Faccion.RecompensasCaos = 7
        UserList(UserIndex).Faccion.NextRecompensa = 2700
    
    Case 2700:
        UserList(UserIndex).Faccion.RecompensasCaos = 8
        UserList(UserIndex).Faccion.NextRecompensa = 4600
    
    Case 4600:
        If Lvl < 30 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 30 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 9
        UserList(UserIndex).Faccion.NextRecompensa = 5800
    
    Case 5800:
        If Lvl < 31 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 31 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 10
        UserList(UserIndex).Faccion.NextRecompensa = 6990
    
    Case 6990:
        If Lvl < 33 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 33 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 11
        UserList(UserIndex).Faccion.NextRecompensa = 8100
    
    Case 8100:
        If Lvl < 35 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 35 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 12
        UserList(UserIndex).Faccion.NextRecompensa = 9300
    
    Case 9300:
        If Lvl < 36 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 36 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 13
        UserList(UserIndex).Faccion.NextRecompensa = 11500
    
    Case 11500:
        If Lvl < 37 Then
            Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & 37 - Lvl & " Niveles para poder recibir la pr�xima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        UserList(UserIndex).Faccion.RecompensasCaos = 14
        UserList(UserIndex).Faccion.NextRecompensa = 23000
    
    Case 23000:
        Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Ciudas & ". Tu �nica recompensa ser� la sangre derramada. ��Continua as�!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Exit Sub
    
    Case Else:
        Exit Sub
        
End Select

Call WriteChatOverHead(UserIndex, "���Bien hecho " + TituloCaos(UserIndex) + ", aqu� tienes tu recompensa!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
Call AsignarExperiencia(UserIndex, ExpX100)


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Legi�n Oscura"
'***************************************************
'Rango 1: Ac�lito (70)
'Rango 2: Alma Corrupta (160)
'Rango 3: Paria (300)
'Rango 4: Condenado (490)
'Rango 5: Esbirro (740)
'Rango 6: Sanguinario (1100)
'Rango 7: Corruptor (1500 + lvl 27)
'Rango 8: Heraldo Impio (2010)
'Rango 9: Caballero de la Oscuridad (2700)
'Rango 10: Se�or del Miedo (4600 + lvl 30)
'Rango 11: Ejecutor Infernal (5800 + lvl 31)
'Rango 12: Protector del Averno (6990 + lvl 33)
'Rango 13: Avatar de la Destrucci�n (8100 + lvl 35)
'Rango 14: Guardi�n del Mal (9300 + lvl 36)
'Rango 15: Campe�n de la Oscuridad (11500 + lvl 37)

Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Ac�lito"
    Case 1
        TituloCaos = "Alma Corrupta"
    Case 2
        TituloCaos = "Paria"
    Case 3
        TituloCaos = "Condenado"
    Case 4
        TituloCaos = "Esbirro"
    Case 5
        TituloCaos = "Sanguinario"
    Case 6
        TituloCaos = "Corruptor"
    Case 7
        TituloCaos = "Heraldo Imp�o"
    Case 8
        TituloCaos = "Caballero de la Oscuridad"
    Case 9
        TituloCaos = "Se�or del Miedo"
    Case 10
        TituloCaos = "Ejecutor Infernal"
    Case 11
        TituloCaos = "Protector del Averno"
    Case 12
        TituloCaos = "Avatar de la Destrucci�n"
    Case 13
        TituloCaos = "Guardi�n del Mal"
    Case Else
        TituloCaos = "Campe�n de la Oscuridad"
End Select

End Function

