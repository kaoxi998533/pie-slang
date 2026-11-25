Attribute VB_Name = "LispCore"
Option Explicit
' ===========================================================================
'  LispCore - value model of the Scheme-subset interpreter: kinds, symbols,
'  lists, numbers, equality, printing, errors and the output buffer.
'
'  Value representation (VBA Variant):
'    '()      -> Empty                 #t / #f -> Boolean
'    integer  -> Long, or Decimal beyond 32 bits (up to 28 digits)
'    real     -> Double                string  -> String
'    everything else -> LispObj (see LK_* kinds)
'  All public names start with "L" so that this interpreter can live in the
'  same workbook as the Python one.
' ===========================================================================

' ---- object kinds ----
Public Const LK_PAIR As Long = 1
Public Const LK_SYM As Long = 2
Public Const LK_LAMBDA As Long = 3
Public Const LK_PRIM As Long = 4
Public Const LK_ENV As Long = 5
Public Const LK_VECTOR As Long = 6
Public Const LK_CHAR As Long = 7
Public Const LK_HASH As Long = 8
Public Const LK_PROMISE As Long = 9
Public Const LK_MACRO As Long = 10
Public Const LK_CONT As Long = 11
Public Const LK_ERROR As Long = 12
Public Const LK_RECTYPE As Long = 13
Public Const LK_RECORD As Long = 14
Public Const LK_RECPROC As Long = 15
Public Const LK_VOID As Long = 16
Public Const LK_VALUES As Long = 17
Public Const LK_MULTI As Long = 18
Public Const LK_EOF As Long = 19

' ---- special forms (Code of the symbol) ----
Public Const LSF_QUOTE As Long = 1
Public Const LSF_QUASIQUOTE As Long = 2
Public Const LSF_UNQUOTE As Long = 3
Public Const LSF_UNQUOTESPLICING As Long = 4
Public Const LSF_IF As Long = 5
Public Const LSF_DEFINE As Long = 6
Public Const LSF_SET As Long = 7
Public Const LSF_LAMBDA As Long = 8
Public Const LSF_BEGIN As Long = 9
Public Const LSF_LET As Long = 10
Public Const LSF_LETSTAR As Long = 11
Public Const LSF_LETREC As Long = 12
Public Const LSF_COND As Long = 13
Public Const LSF_CASE As Long = 14
Public Const LSF_AND As Long = 15
Public Const LSF_OR As Long = 16
Public Const LSF_WHEN As Long = 17
Public Const LSF_UNLESS As Long = 18
Public Const LSF_DO As Long = 19
Public Const LSF_DELAY As Long = 20
Public Const LSF_CONSSTREAM As Long = 21
Public Const LSF_DEFSYNTAX As Long = 22
Public Const LSF_DEFMACRO As Long = 23
Public Const LSF_GUARD As Long = 24
Public Const LSF_DEFRECORD As Long = 25
Public Const LSF_ASSERT As Long = 26
Public Const LSF_DELAYFORCE As Long = 27
Public Const LSF_SYNTAXRULES As Long = 28
Public Const LSF_IGNORE As Long = 29
Public Const LSF_CASELAMBDA As Long = 30
Public Const LSF_LETVALUES As Long = 31
Public Const LSF_DEFVALUES As Long = 32

' error numbers used to propagate Lisp errors and escaping continuations
Public Const LERR As Long = -2147219503      ' vbObjectError + 2001
Public Const LESCAPE As Long = -2147219502   ' vbObjectError + 2002

Public Const L_INF As Double = 1.797693134862E+308

' ---- interpreter state ----
Public lgGlobal As LispObj
Public lgSymtab As LispObj
Public lgVoid As LispObj
Public lgEof As LispObj
Public lgDepth As Long
Public lgMaxDepth As Long
Public lgLine As Long
Public lgErr As Variant           ' object being raised
Public lgEscK As LispObj          ' continuation being invoked
Public lgEscVal As Variant
Public lgNextId As Long
Public lgGensym As Long
Public lgDecSep As String
' frequently used symbols
Public lsQuote As LispObj
Public lsQuasi As LispObj
Public lsUnquote As LispObj
Public lsUnquoteSplicing As LispObj
Public lsElse As LispObj
Public lsArrow As LispObj
Public lsEllipsis As LispObj
Public lsUnderscore As LispObj
Public lsLambda As LispObj
Public lsBegin As LispObj
Public lsDefine As LispObj
Public lsDot As LispObj

Private mOutParts() As String
Private mOutN As Long
Private mOutCur As String
Private mOutLast As String
Private mKeep() As LispObj
Private mKeepN As Long

' ===========================================================================
'  objects
' ===========================================================================
Public Function LNewObj(ByVal kind As Long) As LispObj
    Dim o As LispObj
    Set o = New LispObj
    o.Kind = kind
    lgNextId = lgNextId + 1
    o.Id = lgNextId
    Set LNewObj = o
End Function

Public Sub LSetV(ByRef dst As Variant, ByRef src As Variant)
    If IsObject(src) Then
        Set dst = src
    Else
        dst = src
    End If
End Sub

Public Function LCons(ByRef a As Variant, ByRef b As Variant) As LispObj
    Dim p As LispObj
    Set p = New LispObj
    p.Kind = LK_PAIR
    lgNextId = lgNextId + 1
    p.Id = lgNextId
    If IsObject(a) Then
        Set p.Car = a
    Else
        p.Car = a
    End If
    If IsObject(b) Then
        Set p.Cdr = b
    Else
        p.Cdr = b
    End If
    Set LCons = p
End Function

Public Function LNewEnv(ByVal parent As LispObj) As LispObj
    Dim e As LispObj
    Set e = LNewObj(LK_ENV)
    Set e.Env = parent
    Set LNewEnv = e
End Function

Public Function LIntern(ByVal nm As String) As LispObj
    Dim h As Long, i As Long, s As LispObj, v As Variant
    h = LHashStr(nm)
    i = lgSymtab.FindH(nm, h)
    If i >= 0 Then
        lgSymtab.ValInto i, v
        Set LIntern = v
        Exit Function
    End If
    Set s = LNewObj(LK_SYM)
    s.SName = nm
    s.H = h
    Set v = s
    lgSymtab.SetH nm, h, nm, v
    Set LIntern = s
End Function

Public Function LMakeChar(ByVal c As String) As LispObj
    Dim o As LispObj
    Set o = LNewObj(LK_CHAR)
    o.SName = c
    Set LMakeChar = o
End Function

' keep objects that may form reference cycles, so they can be freed
Public Sub LKeep(ByVal o As LispObj)
    If mKeepN = 0 Then
        ReDim mKeep(0 To 63)
    End If
    If mKeepN > UBound(mKeep) Then
        ReDim Preserve mKeep(0 To 2 * mKeepN + 1)
    End If
    Set mKeep(mKeepN) = o
    mKeepN = mKeepN + 1
End Sub

Public Sub LReleaseAll()
    Dim i As Long, ko As LispObj
    For i = 0 To mKeepN - 1
        Set ko = mKeep(i)
        If Not (ko Is Nothing) Then ko.Release
        Set mKeep(i) = Nothing
    Next
    mKeepN = 0
    If Not (lgGlobal Is Nothing) Then lgGlobal.Release
    Set lgGlobal = Nothing
    lgErr = Empty
    lgEscVal = Empty
    Set lgEscK = Nothing
End Sub

' ===========================================================================
'  output buffer
' ===========================================================================
Public Sub LOutReset()
    ReDim mOutParts(0 To 15)
    mOutN = 0
    mOutCur = vbNullString
    mOutLast = vbNullString
End Sub

Public Sub LOutWrite(ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    mOutLast = Right$(s, 1)
    mOutCur = mOutCur & s
    If Len(mOutCur) > 2048 Then
        If mOutN > UBound(mOutParts) Then
            ReDim Preserve mOutParts(0 To 2 * mOutN + 1)
        End If
        mOutParts(mOutN) = mOutCur
        mOutN = mOutN + 1
        mOutCur = vbNullString
    End If
End Sub

Public Function LOutText() As String
    If mOutN = 0 Then
        LOutText = mOutCur
    Else
        LOutText = Join(mOutParts, vbNullString) & mOutCur
    End If
End Function

Public Function LOutAtLineStart() As Boolean
    LOutAtLineStart = (Len(mOutLast) = 0 Or mOutLast = vbLf)
End Function

' ===========================================================================
'  hashing
' ===========================================================================
Public Function LHashStr(ByVal s As String) As Long
    Dim b() As Byte, i As Long, h As Long, n As Long, stp As Long
    h = 5381
    n = Len(s)
    If n = 0 Then
        LHashStr = h
        Exit Function
    End If
    b = s
    If n <= 48 Then
        For i = 0 To 2 * n - 1 Step 2
            h = ((h * 33) + b(i) + 256& * b(i + 1)) And &H3FFFFF
        Next
    Else
        For i = 0 To 31 Step 2
            h = ((h * 33) + b(i) + 256& * b(i + 1)) And &H3FFFFF
        Next
        stp = ((n - 32) \ 16 + 1) * 2
        For i = 32 To 2 * (n - 16) - 1 Step stp
            h = ((h * 33) + b(i) + 256& * b(i + 1)) And &H3FFFFF
        Next
        For i = 2 * (n - 16) To 2 * n - 1 Step 2
            h = ((h * 33) + b(i) + 256& * b(i + 1)) And &H3FFFFF
        Next
        h = ((h * 33) + (n And &HFFFF&)) And &H3FFFFF
    End If
    LHashStr = h
End Function

' key string for hash tables (equal? semantics)
Public Function LKeyStr(ByRef v As Variant) As String
    Dim o As LispObj, k As String, x As Variant, parts() As String, n As Long, i As Long
    Select Case VarType(v)
    Case vbString
        LKeyStr = "s" & v
    Case vbLong, vbInteger, vbByte, vbDecimal
        LKeyStr = "i" & CStr(v)
    Case vbDouble, vbSingle, vbCurrency
        LKeyStr = "f" & LFloatRepr(CDbl(v))
    Case vbBoolean
        If v Then
            LKeyStr = "b1"
        Else
            LKeyStr = "b0"
        End If
    Case vbEmpty
        LKeyStr = "n"
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case LK_SYM
            LKeyStr = "y" & o.SName
        Case LK_CHAR
            LKeyStr = "c" & o.SName
        Case LK_PAIR
            k = LKeyStr(o.Car)
            LKeyStr = "p" & Len(k) & ":" & k & LKeyStr(o.Cdr)
        Case LK_VECTOR
            If o.Count = 0 Then
                LKeyStr = "v0:"
                Exit Function
            End If
            ReDim parts(0 To o.Count - 1)
            For i = 0 To o.Count - 1
                o.GetInto i, x
                k = LKeyStr(x)
                parts(i) = Len(k) & ":" & k
            Next
            LKeyStr = "v" & o.Count & ":" & Join(parts, vbNullString)
        Case Else
            LKeyStr = "o" & o.Id
        End Select
    Case Else
        LKeyStr = "?" & CStr(v)
    End Select
End Function

' ===========================================================================
'  predicates and list helpers
' ===========================================================================
Public Function LIsKind(ByRef v As Variant, ByVal kind As Long) As Boolean
    Dim o As LispObj
    If IsObject(v) Then
        Set o = v
        LIsKind = (o.Kind = kind)
    End If
End Function

Public Function LIsPair(ByRef v As Variant) As Boolean
    Dim o As LispObj
    If IsObject(v) Then
        Set o = v
        LIsPair = (o.Kind = LK_PAIR)
    End If
End Function

Public Function LIsSym(ByRef v As Variant) As Boolean
    Dim o As LispObj
    If IsObject(v) Then
        Set o = v
        LIsSym = (o.Kind = LK_SYM)
    End If
End Function

Public Function LIsProc(ByRef v As Variant) As Boolean
    Dim o As LispObj
    If IsObject(v) Then
        Set o = v
        Select Case o.Kind
        Case LK_LAMBDA, LK_PRIM, LK_CONT, LK_RECPROC
            LIsProc = True
        End Select
    End If
End Function

Public Function LTruthy(ByRef v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        LTruthy = v
    Else
        LTruthy = True
    End If
End Function

Public Function LCar(ByRef v As Variant) As Variant
    Dim o As LispObj
    If Not LIsPair(v) Then LThrow2 "car: expected a pair, got", v
    Set o = v
    If IsObject(o.Car) Then
        Set LCar = o.Car
    Else
        LCar = o.Car
    End If
End Function

Public Function LCdr(ByRef v As Variant) As Variant
    Dim o As LispObj
    If Not LIsPair(v) Then LThrow2 "cdr: expected a pair, got", v
    Set o = v
    If IsObject(o.Cdr) Then
        Set LCdr = o.Cdr
    Else
        LCdr = o.Cdr
    End If
End Function

' length of a proper list, -1 if improper
Public Function LLength(ByRef v As Variant) As Long
    Dim x As Variant, o As LispObj, n As Long
    LSetV x, v
    Do While IsObject(x)
        Set o = x
        If o.Kind <> LK_PAIR Then
            LLength = -1
            Exit Function
        End If
        n = n + 1
        LSetV x, o.Cdr
    Loop
    If IsEmpty(x) Then
        LLength = n
    Else
        LLength = -1
    End If
End Function

Public Function LListFromArray(ByRef arr() As Variant, ByVal first As Long, ByVal last As Long) As Variant
    Dim r As Variant, i As Long
    r = Empty
    For i = last To first Step -1
        Set r = LCons(arr(i), r)
    Next
    If IsObject(r) Then
        Set LListFromArray = r
    Else
        LListFromArray = r
    End If
End Function

' copy the items of a proper list into arr(0..n-1)
Public Sub LListToArray(ByRef v As Variant, ByRef arr() As Variant, ByRef n As Long, ByVal who As String)
    Dim x As Variant, o As LispObj
    n = 0
    ReDim arr(0 To 7)
    LSetV x, v
    Do While IsObject(x)
        Set o = x
        If o.Kind <> LK_PAIR Then Exit Do
        If n > UBound(arr) Then
            ReDim Preserve arr(0 To 2 * n + 1)
        End If
        If IsObject(o.Car) Then
            Set arr(n) = o.Car
        Else
            arr(n) = o.Car
        End If
        n = n + 1
        LSetV x, o.Cdr
    Loop
    If Not IsEmpty(x) Then LThrow2 who & ": expected a proper list, got", v
End Sub

Public Function LList2(ByRef a As Variant, ByRef b As Variant) As LispObj
    Dim e As Variant
    Set LList2 = LCons(a, LCons(b, e))
End Function

' ===========================================================================
'  numbers
' ===========================================================================
Public Function LIsNum(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbLong, vbDouble, vbInteger, vbDecimal, vbSingle, vbCurrency, vbByte
        LIsNum = True
    End Select
End Function

Public Function LIsExact(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbLong, vbInteger, vbDecimal, vbByte
        LIsExact = True
    End Select
End Function

Public Function LIsFlo(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbDouble, vbSingle, vbCurrency
        LIsFlo = True
    End Select
End Function

Public Function LNorm(ByVal d As Variant) As Variant
    If VarType(d) = vbLong Then
        LNorm = d
    ElseIf VarType(d) = vbInteger Or VarType(d) = vbByte Then
        LNorm = CLng(d)
    ElseIf d >= -2147483648# And d <= 2147483647# Then
        LNorm = CLng(d)
    Else
        LNorm = d
    End If
End Function

Public Function LDecFromString(ByVal s As String) As Variant
    On Error GoTo Ovf
    LDecFromString = CDec(s)
    Exit Function
Ovf:
    LThrow "integer too large (more than 28 digits)"
End Function

Public Function LIntFromDouble(ByVal d As Double) As Variant
    If d >= -2147483648# And d <= 2147483647# Then
        LIntFromDouble = CLng(d)
    Else
        LIntFromDouble = LDecFromString(Format$(d, "0"))
    End If
End Function

Private Function DecArith(ByVal op As Long, ByVal a As Variant, ByVal b As Variant) As Variant
    On Error GoTo Ovf
    Select Case op
    Case 1: DecArith = CDec(a) + CDec(b)
    Case 2: DecArith = CDec(a) - CDec(b)
    Case 3: DecArith = CDec(a) * CDec(b)
    End Select
    Exit Function
Ovf:
    LThrow "integer overflow (result has more than 28 digits)"
End Function

Public Function LIntAdd(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) + CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            LIntAdd = CLng(d)
            Exit Function
        End If
    End If
    LIntAdd = LNorm(DecArith(1, a, b))
End Function

Public Function LIntSub(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) - CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            LIntSub = CLng(d)
            Exit Function
        End If
    End If
    LIntSub = LNorm(DecArith(2, a, b))
End Function

Public Function LIntMul(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) * CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            LIntMul = CLng(d)
            Exit Function
        End If
    End If
    LIntMul = LNorm(DecArith(3, a, b))
End Function

Public Function LIntNeg(ByVal a As Variant) As Variant
    If VarType(a) <> vbDecimal Then
        a = CLng(a)
        If a = -2147483647 - 1 Then
            LIntNeg = LDecFromString("2147483648")
        Else
            LIntNeg = CLng(-a)
        End If
    Else
        LIntNeg = LNorm(-a)
    End If
End Function

' floor division (q, r with the sign of b)
Public Sub LIntDivMod(ByVal a As Variant, ByVal b As Variant, ByRef q As Variant, ByRef r As Variant)
    Dim dq As Double, dr As Double, da As Double, db As Double, xq As Variant, xr As Variant
    If b = 0 Then LThrow "division by zero"
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        da = a
        db = b
        dq = Int(da / db)
        dr = da - dq * db
        If db > 0 Then
            If dr < 0 Then dq = dq - 1: dr = dr + db
            If dr >= db Then dq = dq + 1: dr = dr - db
        Else
            If dr > 0 Then dq = dq - 1: dr = dr + db
            If dr <= db Then dq = dq + 1: dr = dr - db
        End If
        q = LIntFromDouble(dq)
        r = CLng(dr)
        Exit Sub
    End If
    On Error GoTo Ovf
    xq = Int(CDec(a) / CDec(b))
    xr = CDec(a) - xq * CDec(b)
    If b > 0 Then
        Do While xr < 0
            xq = xq - 1
            xr = xr + b
        Loop
        Do While xr >= b
            xq = xq + 1
            xr = xr - b
        Loop
    Else
        Do While xr > 0
            xq = xq - 1
            xr = xr + b
        Loop
        Do While xr <= b
            xq = xq + 1
            xr = xr - b
        Loop
    End If
    q = LNorm(xq)
    r = LNorm(xr)
    Exit Sub
Ovf:
    LThrow "integer overflow in division"
End Sub

' truncating division (quotient / remainder)
Public Sub LIntQuotRem(ByVal a As Variant, ByVal b As Variant, ByRef q As Variant, ByRef r As Variant)
    LIntDivMod a, b, q, r
    If r <> 0 Then
        If (a < 0) <> (b < 0) Then
            q = LIntAdd(q, 1&)
            r = LIntSub(r, b)
        End If
    End If
End Sub

Public Function LIntPow(ByVal a As Variant, ByVal e As Variant) As Variant
    Dim r As Variant, bs As Variant, n As Long
    If VarType(e) = vbDecimal Then LThrow "expt: exponent too large"
    n = e
    r = 1&
    bs = a
    Do While n > 0
        If (n And 1) = 1 Then r = LIntMul(r, bs)
        n = n \ 2
        If n > 0 Then bs = LIntMul(bs, bs)
    Loop
    LIntPow = r
End Function

Public Function LNumCmp(ByRef a As Variant, ByRef b As Variant) As Long
    Dim dx As Double, dy As Double
    If LIsFlo(a) Or LIsFlo(b) Then
        dx = CDbl(a)
        dy = CDbl(b)
        If dx < dy Then
            LNumCmp = -1
        ElseIf dx > dy Then
            LNumCmp = 1
        End If
    Else
        If a < b Then
            LNumCmp = -1
        ElseIf a > b Then
            LNumCmp = 1
        End If
    End If
End Function

' parse an integer in base 2..36; ok=False if invalid
Public Function LParseInt(ByVal s As String, ByVal bs As Long, ByRef ok As Boolean) As Variant
    Dim i As Long, n As Long, c As String, d As Long, neg As Boolean, r As Variant, k As Long
    ok = False
    n = Len(s)
    If n = 0 Then Exit Function
    i = 1
    c = Left$(s, 1)
    If c = "-" Or c = "+" Then
        neg = (c = "-")
        i = 2
    End If
    If i > n Then Exit Function
    If bs = 10 And n - i + 1 <= 9 Then
        For k = i To n
            c = Mid$(s, k, 1)
            If c < "0" Or c > "9" Then Exit Function
        Next
        r = CLng(Mid$(s, i))
        If neg Then r = CLng(-r)
        LParseInt = r
        ok = True
        Exit Function
    End If
    r = 0&
    For k = i To n
        c = LCase$(Mid$(s, k, 1))
        If c >= "0" And c <= "9" Then
            d = Asc(c) - 48
        ElseIf c >= "a" And c <= "z" Then
            d = Asc(c) - 87
        Else
            Exit Function
        End If
        If d >= bs Then Exit Function
        r = LIntAdd(LIntMul(r, bs), d)
    Next
    If neg Then r = LIntNeg(r)
    LParseInt = r
    ok = True
End Function

' parse a Scheme number literal; ok=False if the text is not a number
Public Function LParseNumber(ByVal s As String, ByVal bs As Long, ByRef ok As Boolean) As Variant
    Dim i As Long, n As Long, c As String, digits As Long, seenDot As Boolean, seenE As Boolean, p As Long
    Dim a As Variant, b As Variant, okA As Boolean, okB As Boolean
    ok = False
    n = Len(s)
    If n = 0 Then Exit Function
    If n > 2 And Left$(s, 1) = "#" Then
        Select Case LCase$(Mid$(s, 2, 1))
        Case "x": LParseNumber = LParseNumber(Mid$(s, 3), 16, ok)
        Case "b": LParseNumber = LParseNumber(Mid$(s, 3), 2, ok)
        Case "o": LParseNumber = LParseNumber(Mid$(s, 3), 8, ok)
        Case "d", "e": LParseNumber = LParseNumber(Mid$(s, 3), bs, ok)
        Case "i"
            LParseNumber = LParseNumber(Mid$(s, 3), bs, ok)
            If ok Then LParseNumber = CDbl(LParseNumber(Mid$(s, 3), bs, ok))
        End Select
        Exit Function
    End If
    Select Case LCase$(s)
    Case "+inf.0"
        LParseNumber = L_INF
        ok = True
        Exit Function
    Case "-inf.0"
        LParseNumber = -L_INF
        ok = True
        Exit Function
    End Select
    If bs <> 10 Then
        LParseNumber = LParseInt(s, bs, ok)
        Exit Function
    End If
    p = InStr(1, s, "/")
    If p > 0 Then
        a = LParseInt(Left$(s, p - 1), 10, okA)
        b = LParseInt(Mid$(s, p + 1), 10, okB)
        If Not (okA And okB) Then Exit Function
        If Left$(Mid$(s, p + 1), 1) = "-" Or Left$(Mid$(s, p + 1), 1) = "+" Then Exit Function
        If b = 0 Then Exit Function
        LParseNumber = LExactDiv(a, b)
        ok = True
        Exit Function
    End If
    LParseNumber = LParseInt(s, 10, ok)
    If ok Then Exit Function
    i = 1
    c = Left$(s, 1)
    If c = "+" Or c = "-" Then i = 2
    For i = i To n
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits + 1
        ElseIf c = "." And Not seenDot And Not seenE Then
            seenDot = True
        ElseIf (c = "e" Or c = "E") And Not seenE And digits > 0 Then
            seenE = True
            If i < n Then
                If Mid$(s, i + 1, 1) = "+" Or Mid$(s, i + 1, 1) = "-" Then i = i + 1
            End If
            If i >= n Then Exit Function
        Else
            Exit Function
        End If
    Next
    If digits = 0 Then Exit Function
    LParseNumber = CDbl(Val(s))
    ok = True
End Function

' exact division: an integer when b divides a, otherwise a real
Public Function LExactDiv(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim q As Variant, r As Variant
    If b = 0 Then LThrow "/: division by zero"
    LIntDivMod a, b, q, r
    If r = 0 Then
        LExactDiv = q
    ElseIf VarType(a) = vbDecimal Or VarType(b) = vbDecimal Then
        LExactDiv = CDbl(CDec(a) / CDec(b))
    Else
        LExactDiv = CDbl(a) / CDbl(b)
    End If
End Function

' ===========================================================================
'  equality
' ===========================================================================
Public Function LEqv(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim oa As LispObj, ob As LispObj
    If IsObject(a) Then
        If Not IsObject(b) Then Exit Function
        If a Is b Then
            LEqv = True
            Exit Function
        End If
        Set oa = a
        Set ob = b
        If oa.Kind = LK_CHAR And ob.Kind = LK_CHAR Then LEqv = (oa.SName = ob.SName)
        Exit Function
    End If
    If IsObject(b) Then Exit Function
    If LIsNum(a) Then
        If LIsNum(b) Then
            If LIsExact(a) = LIsExact(b) Then LEqv = (LNumCmp(a, b) = 0)
        End If
        Exit Function
    End If
    If VarType(a) <> VarType(b) Then Exit Function
    Select Case VarType(a)
    Case vbEmpty
        LEqv = True
    Case vbBoolean, vbString
        LEqv = (a = b)
    End Select
End Function

Public Function LEqual(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim oa As LispObj, ob As LispObj, x As Variant, y As Variant, i As Long
    If IsObject(a) And IsObject(b) Then
        Set oa = a
        Set ob = b
        If oa Is ob Then
            LEqual = True
            Exit Function
        End If
        If oa.Kind <> ob.Kind Then Exit Function
        Select Case oa.Kind
        Case LK_PAIR
            Do
                If Not LEqual(oa.Car, ob.Car) Then Exit Function
                LSetV x, oa.Cdr
                LSetV y, ob.Cdr
                If LIsPair(x) And LIsPair(y) Then
                    Set oa = x
                    Set ob = y
                Else
                    LEqual = LEqual(x, y)
                    Exit Function
                End If
            Loop
        Case LK_VECTOR
            If oa.Count <> ob.Count Then Exit Function
            For i = 0 To oa.Count - 1
                oa.GetInto i, x
                ob.GetInto i, y
                If Not LEqual(x, y) Then Exit Function
            Next
            LEqual = True
        Case Else
            LEqual = LEqv(a, b)
        End Select
        Exit Function
    End If
    LEqual = LEqv(a, b)
End Function

' ===========================================================================
'  printing
' ===========================================================================
Public Function LFloatRepr(ByVal d As Double) As String
    Dim s As String, neg As Boolean, p As Long, ex As Long, mant As String, digits As String
    Dim pointPos As Long, e10 As Long, r As String
    If d = 0 Then
        LFloatRepr = "0.0"
        Exit Function
    End If
    If d >= L_INF Then
        LFloatRepr = "+inf.0"
        Exit Function
    ElseIf d <= -L_INF Then
        LFloatRepr = "-inf.0"
        Exit Function
    End If
    s = Trim$(Str$(d))
    If Left$(s, 1) = "-" Then
        neg = True
        s = Mid$(s, 2)
    End If
    p = InStr(1, s, "E", vbTextCompare)
    If p > 0 Then
        ex = CLng(Val(Mid$(s, p + 1)))
        mant = Left$(s, p - 1)
    Else
        mant = s
    End If
    p = InStr(1, mant, ".")
    If p = 0 Then
        digits = mant
        pointPos = Len(mant)
    Else
        digits = Left$(mant, p - 1) & Mid$(mant, p + 1)
        pointPos = p - 1
    End If
    Do While Len(digits) > 1
        If Left$(digits, 1) <> "0" Then Exit Do
        digits = Mid$(digits, 2)
        pointPos = pointPos - 1
    Loop
    Do While Len(digits) > 1
        If Right$(digits, 1) <> "0" Then Exit Do
        digits = Left$(digits, Len(digits) - 1)
    Loop
    e10 = pointPos + ex - 1
    If e10 >= 21 Or e10 < -7 Then
        r = Left$(digits, 1)
        If Len(digits) > 1 Then r = r & "." & Mid$(digits, 2)
        r = r & "e" & CStr(e10)
    ElseIf e10 >= 0 Then
        If Len(digits) <= e10 + 1 Then
            r = digits & String$(e10 + 1 - Len(digits), "0") & ".0"
        Else
            r = Left$(digits, e10 + 1) & "." & Mid$(digits, e10 + 2)
        End If
    Else
        r = "0." & String$(-e10 - 1, "0") & digits
    End If
    If neg Then r = "-" & r
    LFloatRepr = r
End Function

Public Function LUCode(ByVal c As String) As Long
    Dim k As Long
    k = AscW(c)
    If k < 0 Then k = k + 65536
    LUCode = k
End Function

Private Function WriteString(ByVal s As String) As String
    Dim i As Long, c As String, parts() As String, n As Long
    If InStr(1, s, "\") = 0 And InStr(1, s, """") = 0 And InStr(1, s, vbLf) = 0 And InStr(1, s, vbTab) = 0 And InStr(1, s, vbCr) = 0 Then
        WriteString = """" & s & """"
        Exit Function
    End If
    n = Len(s)
    ReDim parts(0 To n)
    For i = 1 To n
        c = Mid$(s, i, 1)
        Select Case c
        Case "\": c = "\\"
        Case """": c = "\"""
        Case vbLf: c = "\n"
        Case vbTab: c = "\t"
        Case vbCr: c = "\r"
        End Select
        parts(i - 1) = c
    Next
    WriteString = """" & Join(parts, vbNullString) & """"
End Function

Public Function LCharName(ByVal c As String) As String
    Select Case LUCode(c)
    Case 32: LCharName = "space"
    Case 10: LCharName = "newline"
    Case 9: LCharName = "tab"
    Case 13: LCharName = "return"
    Case 0: LCharName = "null"
    Case 127: LCharName = "delete"
    Case 27: LCharName = "escape"
    Case 7: LCharName = "alarm"
    Case 8: LCharName = "backspace"
    Case Else: LCharName = c
    End Select
End Function

Public Function LTypeName(ByRef v As Variant) As String
    Dim o As LispObj
    Select Case VarType(v)
    Case vbEmpty: LTypeName = "null"
    Case vbBoolean: LTypeName = "boolean"
    Case vbLong, vbInteger, vbByte, vbDecimal: LTypeName = "integer"
    Case vbDouble, vbSingle, vbCurrency: LTypeName = "real"
    Case vbString: LTypeName = "string"
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case LK_PAIR: LTypeName = "pair"
        Case LK_SYM: LTypeName = "symbol"
        Case LK_LAMBDA, LK_PRIM, LK_CONT, LK_RECPROC: LTypeName = "procedure"
        Case LK_VECTOR: LTypeName = "vector"
        Case LK_CHAR: LTypeName = "char"
        Case LK_HASH: LTypeName = "hash-table"
        Case LK_PROMISE: LTypeName = "promise"
        Case LK_MACRO: LTypeName = "macro"
        Case LK_ERROR: LTypeName = "error"
        Case LK_RECORD: LTypeName = o.Ref.SName
        Case LK_ENV: LTypeName = "environment"
        Case Else: LTypeName = "object"
        End Select
    Case Else: LTypeName = "object"
    End Select
End Function

' display (wr=False) or write (wr=True) representation
Public Function LToStr(ByRef v As Variant, ByVal wr As Boolean) As String
    Dim o As LispObj, parts() As String, n As Long, x As Variant, i As Long, y As Variant
    Select Case VarType(v)
    Case vbEmpty
        LToStr = "()"
    Case vbBoolean
        If v Then
            LToStr = "#t"
        Else
            LToStr = "#f"
        End If
    Case vbLong, vbInteger, vbByte, vbDecimal
        LToStr = CStr(v)
    Case vbDouble, vbSingle, vbCurrency
        LToStr = LFloatRepr(CDbl(v))
    Case vbString
        If wr Then
            LToStr = WriteString(v)
        Else
            LToStr = v
        End If
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case LK_PAIR
            ReDim parts(0 To 15)
            n = 0
            LSetV x, v
            Do While LIsPair(x)
                Set o = x
                If n > UBound(parts) Then
                    ReDim Preserve parts(0 To 2 * n + 1)
                End If
                parts(n) = LToStr(o.Car, wr)
                n = n + 1
                LSetV x, o.Cdr
                If n > 100000 Then Exit Do
            Loop
            If Not IsEmpty(x) Then
                If n > UBound(parts) - 1 Then
                    ReDim Preserve parts(0 To n + 2)
                End If
                parts(n) = "."
                parts(n + 1) = LToStr(x, wr)
                n = n + 2
            End If
            ReDim Preserve parts(0 To n - 1)
            LToStr = "(" & Join(parts, " ") & ")"
        Case LK_SYM
            LToStr = o.SName
        Case LK_CHAR
            If wr Then
                LToStr = "#\" & LCharName(o.SName)
            Else
                LToStr = o.SName
            End If
        Case LK_VECTOR
            If o.Count = 0 Then
                LToStr = "#()"
            Else
                ReDim parts(0 To o.Count - 1)
                For i = 0 To o.Count - 1
                    o.GetInto i, x
                    parts(i) = LToStr(x, wr)
                Next
                LToStr = "#(" & Join(parts, " ") & ")"
            End If
        Case LK_LAMBDA, LK_PRIM, LK_RECPROC
            If Len(o.SName) > 0 Then
                LToStr = "#<procedure " & o.SName & ">"
            Else
                LToStr = "#<procedure>"
            End If
        Case LK_CONT
            LToStr = "#<continuation>"
        Case LK_MACRO
            LToStr = "#<macro " & o.SName & ">"
        Case LK_HASH
            LToStr = "#<hash-table " & o.DCount & ">"
        Case LK_PROMISE
            LToStr = "#<promise>"
        Case LK_VOID
            LToStr = "#<void>"
        Case LK_EOF
            LToStr = "#<eof>"
        Case LK_ENV
            LToStr = "#<environment>"
        Case LK_ERROR
            ReDim parts(0 To 0)
            parts(0) = "#<error " & LToStr(o.SName, True)
            n = 1
            LSetV x, o.Car
            Do While LIsPair(x)
                Set o = x
                ReDim Preserve parts(0 To n)
                parts(n) = LToStr(o.Car, True)
                n = n + 1
                LSetV x, o.Cdr
            Loop
            LToStr = Join(parts, " ") & ">"
        Case LK_RECTYPE
            LToStr = "#<record-type " & o.SName & ">"
        Case LK_RECORD
            ReDim parts(0 To o.Count)
            parts(0) = "#<" & o.Ref.SName
            For i = 0 To o.Count - 1
                o.GetInto i, x
                o.Ref.GetInto i, y
                parts(i + 1) = LToStr(y, False) & ": " & LToStr(x, True)
            Next
            LToStr = Join(parts, " ") & ">"
        Case LK_VALUES
            If o.Count = 0 Then
                LToStr = ""
            Else
                ReDim parts(0 To o.Count - 1)
                For i = 0 To o.Count - 1
                    o.GetInto i, x
                    parts(i) = LToStr(x, wr)
                Next
                LToStr = Join(parts, " ")
            End If
        Case Else
            LToStr = "#<object>"
        End Select
    Case vbDate
        LToStr = Format$(v, "yyyy-mm-dd hh:nn:ss")
    Case Else
        LToStr = "#<?>"
    End Select
End Function

' ===========================================================================
'  errors
' ===========================================================================
Public Function LMakeError(ByVal msg As String, ByRef irritants As Variant) As LispObj
    Dim e As LispObj
    Set e = LNewObj(LK_ERROR)
    e.SName = msg
    If IsObject(irritants) Then
        Set e.Car = irritants
    Else
        e.Car = irritants
    End If
    e.I1 = lgLine
    Set LMakeError = e
End Function

Public Sub LRaise(ByRef obj As Variant)
    If IsObject(obj) Then
        Set lgErr = obj
    Else
        lgErr = obj
    End If
    Err.Raise LERR, "LispVBA", "Lisp error"
End Sub

Public Sub LThrow(ByVal msg As String)
    Dim e As Variant
    LRaise LMakeError(msg, e)
End Sub

Public Sub LThrow2(ByVal msg As String, ByRef irritant As Variant)
    Dim e As Variant
    LRaise LMakeError(msg, LCons(irritant, e))
End Sub

' the current VBA error as a Lisp condition object
Public Function LCapture() As Variant
    Dim num As Long, desc As String, e As Variant, msg As String
    num = Err.Number
    If num = LERR Then
        If IsObject(lgErr) Then Set LCapture = lgErr Else LCapture = lgErr
        lgErr = Empty
        Exit Function
    End If
    desc = Err.Description
    Select Case num
    Case 6: msg = "numeric overflow"
    Case 11: msg = "division by zero"
    Case 13: msg = "type mismatch (VBA: " & desc & ")"
    Case 9: msg = "index out of range (VBA: " & desc & ")"
    Case 28: msg = "stack overflow (recursion too deep)"
    Case 18: msg = "interrupted by user"
    Case 1004: msg = "Excel error: " & desc
    Case Else: msg = "VBA error " & num & ": " & desc
    End Select
    Set LCapture = LMakeError(msg, e)
End Function

' message of a condition, e.g. "car: expected a pair, got 5"
Public Function LCondReport(ByRef c As Variant) As String
    Dim o As LispObj, s As String, x As Variant
    If LIsKind(c, LK_ERROR) Then
        Set o = c
        s = o.SName
        LSetV x, o.Car
        Do While LIsPair(x)
            Set o = x
            s = s & " " & LToStr(o.Car, True)
            LSetV x, o.Cdr
        Loop
        LCondReport = s
    Else
        LCondReport = LToStr(c, True)
    End If
End Function
