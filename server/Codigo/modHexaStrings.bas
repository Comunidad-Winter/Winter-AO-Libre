Attribute VB_Name = "modHexaStrings"
Option Explicit

Public Function hexMd52Asc(ByVal md5 As String) As String
    Dim i As Integer, l As String
    
    md5 = UCase$(md5)
    If Len(md5) Mod 2 = 1 Then md5 = "0" & md5
    
    For i = 1 To Len(md5) \ 2
        l = mid(md5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr(hexHex2Dec(l))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    Dim i As Integer, l As String
    For i = 1 To Len(hex)
        l = mid(hex, i, 1)
        Select Case l
            Case "A": l = 10
            Case "B": l = 11
            Case "C": l = 12
            Case "D": l = 13
            Case "E": l = 14
            Case "F": l = 15
        End Select
        
        hexHex2Dec = (l * 16 ^ ((Len(hex) - i))) + hexHex2Dec
    Next i
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Integer, l As String
    For i = 1 To Len(Text)
        l = mid(Text, i, 1)
        txtOffset = txtOffset & Chr((Asc(l) + off) Mod 256)
    Next i
End Function
