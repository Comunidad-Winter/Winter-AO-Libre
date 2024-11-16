Attribute VB_Name = "modExtraSecurity"
Option Explicit
Option Base 0

Private m_KeyS As String
Private m_sBoxRC4(0 To 255) As Integer

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' *********************
' * EXPORTED FUNCTION *
' *********************

Public Function mpModExp(strBaseHex As String, strExponentHex As String, strModulusHex As String) As String
    
    Dim abBase() As Byte
    Dim abExponent() As Byte
    Dim abModulus() As Byte
    Dim abResult() As Byte
    Dim nLen As Integer
    Dim n As Integer
    
    ' Convert hex strings to arrays of bytes
    abBase = mpFromHex(strBaseHex)
    abExponent = mpFromHex(strExponentHex)
    abModulus = mpFromHex(strModulusHex)
    
    ' We require all byte arrays to be the same length with the first byte left as zero
    nLen = UBound(abModulus) + 1
    n = UBound(abExponent) + 1
    If n > nLen Then nLen = n
    n = UBound(abBase) + 1
    If n > nLen Then nLen = n
    Call FixArrayDim(abModulus, nLen)
    Call FixArrayDim(abExponent, nLen)
    Call FixArrayDim(abBase, nLen)
    
    ' Do the business
    abResult = aModExp(abBase, abExponent, abModulus, nLen)
    
    ' Convert result to hex
    mpModExp = mpToHex(abResult)
    
    ' Strip leading zeroes
    For n = 1 To Len(mpModExp)
        If mid$(mpModExp, n, 1) <> "0" Then
            Exit For
        End If
    Next
    If n >= Len(mpModExp) Then
        ' Answer is zero
        mpModExp = "0"
    ElseIf n > 1 Then
        ' Zeroes to strip
        mpModExp = mid$(mpModExp, n)
    End If
    
End Function

Public Function RC4_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

    'Convert the data into a byte array
    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call RC4_EncryptByte(ByteArray(), Key)

    'Convert the byte array back into a string
    RC4_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

' **********************
' * INTERNAL FUNCTIONS *
' **********************

Private Function aModExp(abBase() As Byte, abExponent() As Byte, abModulus() As Byte, nLen As Integer) As Variant
' Computes a = b^e mod m and returns the result in a byte array as a VARIANT
    Dim a() As Byte
    Dim e() As Byte
    Dim s() As Byte
    Dim nBits As Long
    
    ' Perform right-to-left binary exponentiation
    ' 1. Set A = 1, S = b
    ReDim a(nLen - 1)
    a(nLen - 1) = 1
    ' NB s and e are trashed so use copies
    s = abBase
    e = abExponent
    ' 2. While e != 0 do:
    For nBits = nLen * 8 To 1 Step -1
        ' 2.1 if e is odd then A = A*S mod m
        If (e(nLen - 1) And &H1) <> 0 Then
            a = aModMult(a, s, abModulus, nLen)
        End If
        ' 2.2 e = e / 2
        Call DivideByTwo(e)
        ' 2.3 if e != 0 then S = S*S mod m
        If aIsZero(e, nLen) Then Exit For
        s = aModMult(s, s, abModulus, nLen)
        DoEvents
    Next
    
    ' 3. Return(A)
    aModExp = a
    
End Function

Private Function aModMult(abX() As Byte, abY() As Byte, abMod() As Byte, nLen As Integer) As Variant
' Returns w = (x * y) mod m
    Dim w() As Byte
    Dim x() As Byte
    Dim y() As Byte
    Dim nBits As Integer
    
    ' 1. Set w = 0, and temps x = abX, y = abY
    ReDim w(nLen - 1)
    x = abX
    y = abY
    ' 2. From LS bit to MS bit of X do:
    For nBits = nLen * 8 To 1 Step -1
        ' 2.1 if x is odd then w = (w + y) mod m
        If (x(nLen - 1) And &H1) <> 0 Then
            Call aModAdd(w, y, abMod, nLen)
        End If
        ' 2.2 x = x / 2
        Call DivideByTwo(x)
        ' 2.3 if x != 0 then y = (y + y) mod m
        If aIsZero(x, nLen) Then Exit For
        Call aModAdd(y, y, abMod, nLen)
    Next
    aModMult = w
    
End Function

Private Function aIsZero(a() As Byte, ByVal nLen As Integer) As Boolean
' Returns true if a is zero
    aIsZero = True
    Do While nLen > 0
        If a(nLen - 1) <> 0 Then
            aIsZero = False
            Exit Do
        End If
        nLen = nLen - 1
    Loop
End Function

Private Sub aModAdd(a() As Byte, b() As Byte, m() As Byte, nLen As Integer)
' Computes a = (a + b) mod m
    Dim i As Integer
    Dim d As Long
    ' 1. Add a = a + b
    d = 0
    For i = nLen - 1 To 0 Step -1
        d = CLng(a(i)) + CLng(b(i)) + d
        a(i) = CByte(d And &HFF)
        d = d \ &H100
    Next
    ' 2. If a > m then a = a - m
    For i = 0 To nLen - 2
        If a(i) <> m(i) Then
            Exit For
        End If
    Next
    If a(i) >= m(i) Then
        Call aSubtract(a, m, nLen)
    End If
    ' 3. Return a in-situ
            
End Sub

Private Sub aSubtract(a() As Byte, b() As Byte, nLen As Integer)
' Computes a = a - b
    Dim i As Integer
    Dim borrow As Long
    Dim d As Long   ' NB d is signed
    
    borrow = 0
    For i = nLen - 1 To 0 Step -1
        d = CLng(a(i)) - CLng(b(i)) - borrow
        If d < 0 Then
            d = d + &H100
            borrow = 1
        Else
            borrow = 0
        End If
        a(i) = CByte(d And &HFF)
    Next
    
End Sub

Private Sub DivideByTwo(ByRef x() As Byte)
' Divides multiple-precision integer x by 2 by shifting to right by one bit
    Dim d As Long
    Dim i As Integer
    d = 0
    For i = 0 To UBound(x)
        d = d Or x(i)
        x(i) = CByte((d \ 2) And &HFF)
        If (d And &H1) Then
            d = &H100
        Else
            d = 0
        End If
    Next
End Sub

Private Function mpToHex(abNum() As Byte) As String
' Returns a string containg the mp number abNum encoded in hex
' with leading zeroes trimmed.
    Dim i As Integer
    Dim sHex As String
    sHex = ""
    For i = 0 To UBound(abNum)
        If abNum(i) < &H10 Then
            sHex = sHex & "0" & Hex(abNum(i))
        Else
            sHex = sHex & Hex(abNum(i))
        End If
    Next
    mpToHex = sHex
End Function

Private Function mpFromHex(ByVal strHex As String) As Variant
' Converts number encoded in hex in big-endian order to a multi-precision integer
' Returns an array of bytes as a VARIANT
' containing number in big-endian order
' but with the first byte always zero
' strHex must only contain valid hex digits [0-9A-Fa-f]
' [2007-10-13] Changed direct >= <= comparisons with strings.
    Dim abData() As Byte
    Dim ib As Long
    Dim ic As Long
    Dim ch As String
    Dim nch As Long
    Dim nLen As Long
    Dim t As Long
    Dim v As Long
    Dim j As Integer
    
    ' Cope with odd # of digits, e.g. "fed" => "0fed"
    If Len(strHex) Mod 2 > 0 Then
        strHex = "0" & strHex
    End If
    nLen = Len(strHex) \ 2 + 1
    ReDim abData(nLen - 1)
    ib = 1
    j = 0
    For ic = 1 To Len(strHex)
        ch = mid$(strHex, ic, 1)
        nch = Asc(ch)
        ''If ch >= "0" And ch <= "9" Then
        If nch >= &H30 And nch <= &H39 Then
            ''t = Asc(ch) - Asc("0")
            t = nch - &H30
        ''ElseIf ch >= "a" And ch <= "f" Then
        ElseIf nch >= &H61 And nch <= &H66 Then
            ''t = Asc(ch) - Asc("a") + 10
            t = nch - &H61 + 10
        ''ElseIf ch >= "A" And ch <= "F" Then
        ElseIf nch >= &H41 And nch <= &H46 Then
            ''t = Asc(ch) - Asc("A") + 10
            t = nch - &H41 + 10
        Else
            ' Invalid digit
            ' Flag error?
            Debug.Print "ERROR: Invalid Hex character found!"
            Exit Function
        End If
        ' Store byte value on every alternate digit
        If j = 0 Then
            ' v = t << 8
            v = t * &H10
            j = 1
        Else
            ' b[i] = (v | t) & 0xff
            abData(ib) = CByte((v Or t) And &HFF)
            ib = ib + 1
            j = 0
        End If
    Next
        
    mpFromHex = abData
End Function

Private Sub FixArrayDim(ByRef abData() As Byte, ByVal nLen As Long)
' Redim abData to be nLen bytes long with existing contents
' aligned at the RHS of the extended array
    Dim oLen As Long
    Dim i As Long
    
    oLen = UBound(abData) + 1
    If oLen > nLen Then
        ' Truncate
        ReDim Preserve abData(nLen - 1)
    ElseIf oLen < nLen Then
        ' Shift right
        ReDim Preserve abData(nLen - 1)
        For i = oLen - 1 To 0 Step -1
            abData(i + nLen - oLen) = abData(i)
        Next
        For i = 0 To nLen - oLen - 1
            abData(i) = 0
        Next
    End If
        
End Sub

Private Sub RC4_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim i As Long
Dim j As Long
Dim Temp As Byte
Dim Offset As Long
Dim OrigLen As Long
Dim sBox(0 To 255) As Integer

    'Set the new key (optional)
    If (Len(Key) > 0) Then RC4_SetKey Key

    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMem(sBox(0), m_sBoxRC4(0), 512)

    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1

    'Encrypt the data
    For Offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        Temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = Temp
        ByteArray(Offset) = ByteArray(Offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
    Next

End Sub

Private Sub RC4_SetKey(New_Value As String)

Dim a As Long
Dim b As Long
Dim Temp As Byte
Dim Key() As Byte
Dim KeyLen As Long

    'Do nothing if the key is buffered
    If (m_KeyS = New_Value) Then Exit Sub

    'Set the new key
    m_KeyS = New_Value

    'Save the password in a byte array
    Key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)

    'Initialize s-boxes
    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    For a = 0 To 255
        b = (b + m_sBoxRC4(a) + Key(a Mod KeyLen)) Mod 256
        Temp = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(b)
        m_sBoxRC4(b) = Temp
    Next

End Sub
