Attribute VB_Name = "InvNpc"
Option Explicit
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo Inv & Obj
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Modulo para controlar los objetos y los inventarios.
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
Public Function TirarItemAlPiso(Pos As WorldPos, Obj As Obj) As WorldPos
On Error GoTo errhandler

    Dim NuevaPos As WorldPos
    NuevaPos.X = 0
    NuevaPos.Y = 0
    Call Tilelibre(Pos, NuevaPos, Obj)
    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
          Call MakeObj(SendTarget.ToMap, 0, Pos.Map, _
                Obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
          TirarItemAlPiso = NuevaPos
    End If

Exit Function
errhandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next

If npc.Invent.NroItems > 0 Then
Dim i As Byte
Dim MiObj As Obj

For i = 1 To MAX_INVENTORY_SLOTS
If npc.Invent.Object(i).ObjIndex > 0 Then
If RandomNumber(1, 100) <= npc.Invent.Object(i).ProbTirar Then
MiObj.Amount = npc.Invent.Object(i).Amount
MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
Call TirarItemAlPiso(npc.Pos, MiObj)
End If
End If
Next i
End If
End Sub
Public Sub NpcTiraEstrella(ByRef npc As npc)
'TIRA TODOS LOS ITEMS DEL NPC
On Error Resume Next
   
Dim EstrellaNacimiento As Obj
   
Dim EstrellaRandom
Dim Pos1A
Dim Pos10B
'NO VALORES NEGATIVOS NI NULOS
Pos1A = 2 'PROBABILIDAD MINIMA
Pos10B = 10  'PROBABILIDAD MAXIMA
'Ecuacion de probabilidad
'Pos1A sobre Pos2B.
'ej: 1 de cada 5
EstrellaRandom = RandomNumber(Pos1A, Pos10B) 'RANDOM DE PROBABILIDADES
If EstrellaRandom = 1 Then
EstrellaNacimiento.Amount = 1 'CANTIDAD
EstrellaNacimiento.ObjIndex = 1000 'NUMERO DEL ITEM EN EL OBJ.DAT
Call TirarItemAlPiso(npc.Pos, EstrellaNacimiento)
End If

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error Resume Next
'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

Dim i As Integer
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
            QuedanItems = True
            Exit Function
        End If
    Next
End If
QuedanItems = False
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
On Error Resume Next
'Devuelve la cantidad original del obj de un npc

Dim ln As String, npcfile As String
Dim i As Integer

If Npclist(NpcIndex).Numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If
 
For i = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)
    If ObjIndex = val(ReadField(1, ln, 45)) Then
        EncontrarCant = val(ReadField(2, ln, 45))
        Exit Function
    End If
Next
                   
EncontrarCant = 50

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
On Error Resume Next

Dim i As Integer

Npclist(NpcIndex).Invent.NroItems = 0

For i = 1 To MAX_INVENTORY_SLOTS
   Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
   Npclist(NpcIndex).Invent.Object(i).Amount = 0
Next i

Npclist(NpcIndex).InvReSpawn = 0

End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = Npclist(NpcIndex).Invent.Object(Slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(Slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(Slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                   
                   Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
                   Npclist(NpcIndex).Invent.Object(Slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                   Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
            
            End If
            
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
               Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    
    
    
    End If
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)

'Vuelve a cargar el inventario del npc NpcIndex
Dim LoopC As Integer
Dim ln As String
Dim npcfile As String

If Npclist(NpcIndex).Numero > 499 Then
    npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "NPCs.dat"
End If

Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
Next LoopC

End Sub


