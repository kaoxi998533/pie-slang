Attribute VB_Name = "PyRuntime"
Option Explicit
' ===========================================================================
'  PyRuntime - value model and core operations of the Python-subset
'  interpreter (numbers, strings, containers, equality, ordering, repr,
'  iteration, indexing, attributes, operators, exceptions).
'
'  Value representation (VBA Variant):
'    None  -> Empty            bool  -> Boolean
'    int   -> Long, or Decimal when it does not fit in 32 bits (<= 28 digits)
'    float -> Double           str   -> String
'    everything else -> PyObj (see OK_* kinds)
' ===========================================================================

' ---- object kinds ----
Public Const OK_LIST As Long = 1
Public Const OK_TUPLE As Long = 2
Public Const OK_DICT As Long = 3
Public Const OK_SET As Long = 4
Public Const OK_FUNC As Long = 5
Public Const OK_CLASS As Long = 6
Public Const OK_INST As Long = 7
Public Const OK_BOUND As Long = 8
Public Const OK_BUILTIN As Long = 9
Public Const OK_BMETHOD As Long = 10
Public Const OK_MODULE As Long = 11
Public Const OK_ENV As Long = 12
Public Const OK_RANGE As Long = 13
Public Const OK_ITER As Long = 14
Public Const OK_SUPER As Long = 15
Public Const OK_SENTINEL As Long = 16
Public Const OK_PROP As Long = 17

' ---- iterator kinds (PyObj.Code of OK_ITER) ----
Public Const IT_SEQ As Long = 1
Public Const IT_STR As Long = 2
Public Const IT_RANGE As Long = 3
Public Const IT_USER As Long = 4

' ---- AST node kinds ----
Public Const NK_BLOCK As Long = 1
Public Const NK_EXPR As Long = 2
Public Const NK_ASSIGN As Long = 3
Public Const NK_AUGASSIGN As Long = 4
Public Const NK_IF As Long = 5
Public Const NK_WHILE As Long = 6
Public Const NK_FOR As Long = 7
Public Const NK_BREAK As Long = 8
Public Const NK_CONTINUE As Long = 9
Public Const NK_PASS As Long = 10
Public Const NK_RETURN As Long = 11
Public Const NK_DEF As Long = 12
Public Const NK_CLASS As Long = 13
Public Const NK_GLOBAL As Long = 14
Public Const NK_NONLOCAL As Long = 15
Public Const NK_DEL As Long = 16
Public Const NK_ASSERT As Long = 17
Public Const NK_RAISE As Long = 18
Public Const NK_TRY As Long = 19
Public Const NK_HANDLER As Long = 20
Public Const NK_IMPORT As Long = 21
Public Const NK_FROMIMPORT As Long = 22
Public Const NK_YIELD As Long = 23
Public Const NK_ANNASSIGN As Long = 24
Public Const NK_CONST As Long = 30
Public Const NK_NAME As Long = 31
Public Const NK_BINOP As Long = 32
Public Const NK_UNARY As Long = 33
Public Const NK_AND As Long = 34
Public Const NK_OR As Long = 35
Public Const NK_NOT As Long = 36
Public Const NK_COMPARE As Long = 37
Public Const NK_CHAIN As Long = 38
Public Const NK_IFEXP As Long = 39
Public Const NK_CALL As Long = 40
Public Const NK_ATTR As Long = 41
Public Const NK_SUBSCR As Long = 42
Public Const NK_SLICE As Long = 43
Public Const NK_LIST As Long = 44
Public Const NK_TUPLE As Long = 45
Public Const NK_DICT As Long = 46
Public Const NK_SET As Long = 47
Public Const NK_LISTCOMP As Long = 48
Public Const NK_SETCOMP As Long = 49
Public Const NK_DICTCOMP As Long = 50
Public Const NK_GENEXP As Long = 51
Public Const NK_COMPFOR As Long = 52
Public Const NK_COMPIF As Long = 53
Public Const NK_LAMBDA As Long = 54
Public Const NK_FSTRING As Long = 55
Public Const NK_FVALUE As Long = 56
Public Const NK_STARRED As Long = 57
Public Const NK_KWARG As Long = 58
Public Const NK_DSTAR As Long = 59
Public Const NK_PARAM As Long = 60
Public Const NK_WALRUS As Long = 61

' ---- operators ----
Public Const OP_ADD As Long = 1
Public Const OP_SUB As Long = 2
Public Const OP_MUL As Long = 3
Public Const OP_DIV As Long = 4
Public Const OP_FLOORDIV As Long = 5
Public Const OP_MOD As Long = 6
Public Const OP_POW As Long = 7
Public Const OP_LSHIFT As Long = 8
Public Const OP_RSHIFT As Long = 9
Public Const OP_BITAND As Long = 10
Public Const OP_BITOR As Long = 11
Public Const OP_BITXOR As Long = 12
Public Const OP_MATMUL As Long = 13
Public Const OP_EQ As Long = 20
Public Const OP_NE As Long = 21
Public Const OP_LT As Long = 22
Public Const OP_LE As Long = 23
Public Const OP_GT As Long = 24
Public Const OP_GE As Long = 25
Public Const OP_IN As Long = 26
Public Const OP_NOTIN As Long = 27
Public Const OP_IS As Long = 28
Public Const OP_ISNOT As Long = 29
Public Const OP_NEG As Long = 30
Public Const OP_POS As Long = 31
Public Const OP_INVERT As Long = 32

' ---- statement completion status ----
Public Const ST_NORMAL As Long = 0
Public Const ST_BREAK As Long = 1
Public Const ST_CONTINUE As Long = 2
Public Const ST_RETURN As Long = 3

' error number used to propagate Python exceptions (vbObjectError + 1001)
Public Const PYERR As Long = -2147220503

' ---- interpreter state ----
Public gGlobals As PyObj
Public gBuiltins As PyObj
Public gModules As PyObj
Public gExc As PyObj            ' exception currently being raised
Public gCurExc As PyObj         ' exception currently being handled
Public gRet As Variant          ' value of a pending 'return'
Public gDepth As Long
Public gMaxDepth As Long
Public gFrameName() As String
Public gFrameLine() As Long
Public gNextId As Long
Public gNoDefault As PyObj
Public gNotImpl As PyObj
Public gObjectClass As PyObj
Public gDecSep As String

Private gOutParts() As String
Private gOutN As Long
Private gOutCur As String
Private gKeep() As PyObj
Private gKeepN As Long

' ===========================================================================
'  objects, output, cleanup
' ===========================================================================
Public Function NewObj(ByVal kind As Long) As PyObj
    Dim o As PyObj
    Set o = New PyObj
    o.Kind = kind
    gNextId = gNextId + 1
    o.Id = gNextId
    Set NewObj = o
End Function

Public Function NewList() As PyObj
    Set NewList = NewObj(OK_LIST)
End Function

Public Function NewDict() As PyObj
    Set NewDict = NewObj(OK_DICT)
End Function

Public Function NewTuple2(ByRef a As Variant, ByRef b As Variant) As PyObj
    Dim t As PyObj
    Set t = NewObj(OK_TUPLE)
    t.Push a
    t.Push b
    Set NewTuple2 = t
End Function

Public Function NewRange(ByVal st As Long, ByVal en As Long, ByVal stp As Long) As PyObj
    Dim r As PyObj
    Set r = NewObj(OK_RANGE)
    r.I1 = st
    r.I2 = en
    r.I3 = stp
    Set NewRange = r
End Function

Public Function RangeLen(ByVal r As PyObj) As Long
    Dim d As Double
    If r.I3 > 0 Then
        If r.I2 <= r.I1 Then Exit Function
        d = Int((CDbl(r.I2) - CDbl(r.I1) - 1#) / CDbl(r.I3)) + 1#
    Else
        If r.I2 >= r.I1 Then Exit Function
        d = Int((CDbl(r.I1) - CDbl(r.I2) - 1#) / CDbl(-r.I3)) + 1#
    End If
    RangeLen = CLng(d)
End Function

Public Sub SetV(ByRef dst As Variant, ByRef src As Variant)
    If IsObject(src) Then
        Set dst = src
    Else
        dst = src
    End If
End Sub

' keep objects that take part in reference cycles so they can be freed
Public Sub KeepObj(ByVal o As PyObj)
    If gKeepN = 0 Then
        ReDim gKeep(0 To 63)
    End If
    If gKeepN > UBound(gKeep) Then
        ReDim Preserve gKeep(0 To 2 * gKeepN + 1)
    End If
    Set gKeep(gKeepN) = o
    gKeepN = gKeepN + 1
End Sub

Public Sub ReleaseAll()
    Dim i As Long, ko As PyObj
    For i = 0 To gKeepN - 1
        Set ko = gKeep(i)
        If Not (ko Is Nothing) Then ko.Release
        Set gKeep(i) = Nothing
    Next
    gKeepN = 0
    If Not (gGlobals Is Nothing) Then gGlobals.Release
    If Not (gModules Is Nothing) Then gModules.Release
    Set gGlobals = Nothing
    Set gModules = Nothing
    Set gExc = Nothing
    Set gCurExc = Nothing
    gRet = Empty
End Sub

Public Sub OutReset()
    ReDim gOutParts(0 To 15)
    gOutN = 0
    gOutCur = vbNullString
End Sub

Public Sub OutWrite(ByVal s As String)
    gOutCur = gOutCur & s
    If Len(gOutCur) > 2048 Then
        If gOutN > UBound(gOutParts) Then
            ReDim Preserve gOutParts(0 To 2 * gOutN + 1)
        End If
        gOutParts(gOutN) = gOutCur
        gOutN = gOutN + 1
        gOutCur = vbNullString
    End If
End Sub

Public Function OutText() As String
    If gOutN = 0 Then
        OutText = gOutCur
    Else
        OutText = Join(gOutParts, vbNullString) & gOutCur
    End If
End Function

' ===========================================================================
'  hashing / dictionary helpers
' ===========================================================================
Public Function HashStr(ByVal s As String) As Long
    Dim b() As Byte, i As Long, h As Long, n As Long, stp As Long
    h = 5381
    n = Len(s)
    If n = 0 Then
        HashStr = h
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
    HashStr = h
End Function

' canonical key string of a hashable value (1 == 1.0 == True share a key)
Public Function KeyStr(ByRef v As Variant) As String
    Dim o As PyObj, i As Long, k As String, x As Variant, parts() As String
    Select Case VarType(v)
    Case vbString
        KeyStr = "s" & v
    Case vbLong, vbInteger, vbByte, vbDecimal
        KeyStr = "i" & CStr(v)
    Case vbBoolean
        If v Then
            KeyStr = "i1"
        Else
            KeyStr = "i0"
        End If
    Case vbDouble, vbSingle, vbCurrency
        If Abs(v) < 1E+15 Then
            If v = Int(v) Then
                KeyStr = "i" & Format$(v, "0")
                Exit Function
            End If
        End If
        KeyStr = "f" & FloatRepr(CDbl(v))
    Case vbEmpty, vbNull
        KeyStr = "N"
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case OK_TUPLE
            If o.Count = 0 Then
                KeyStr = "t0:"
                Exit Function
            End If
            ReDim parts(0 To o.Count - 1)
            For i = 0 To o.Count - 1
                o.GetInto i, x
                k = KeyStr(x)
                parts(i) = Len(k) & ":" & k
            Next
            KeyStr = "t" & o.Count & ":" & Join(parts, vbNullString)
        Case OK_LIST, OK_DICT, OK_SET
            ThrowErr "TypeError", "unhashable type: '" & TypeNameOf(v) & "'"
        Case OK_INST
            KeyStr = InstKeyStr(o)
        Case Else
            KeyStr = "o" & o.Id
        End Select
    Case Else
        KeyStr = "?" & CStr(v)
    End Select
End Function

Private Function InstKeyStr(ByVal o As PyObj) As String
    Dim f As Variant, r As Variant, v As Variant
    If ClassLookup(o.Cls, "__hash__", HashStr("__hash__"), f) Then
        If IsEmpty(f) Then ThrowErr "TypeError", "unhashable type: '" & o.Cls.PName & "'"
        If IsKind(f, OK_FUNC) Then
            Set v = o
            CallValue1 f, v, r
            InstKeyStr = "h" & PyRepr(r)
            Exit Function
        End If
    End If
    InstKeyStr = "o" & o.Id
End Function

' 31-bit string hash used by hash()
Public Function HashStr31(ByVal s As String) As Long
    Dim b() As Byte, i As Long, h As Double
    h = 5381
    If Len(s) > 0 Then
        b = s
        For i = 0 To UBound(b) Step 2
            h = h * 131# + b(i) + 256# * b(i + 1)
            h = h - Int(h / 2147483629#) * 2147483629#
        Next
    End If
    HashStr31 = CLng(h)
End Function

Public Function DictFind(ByVal d As PyObj, ByRef key As Variant) As Long
    Dim ks As String
    ks = KeyStr(key)
    DictFind = d.FindH(ks, HashStr(ks))
End Function

Public Function DictGet(ByVal d As PyObj, ByRef key As Variant, ByRef out As Variant) As Boolean
    Dim e As Long
    e = DictFind(d, key)
    If e >= 0 Then
        d.ValInto e, out
        DictGet = True
    End If
End Function

Public Sub DictSet(ByVal d As PyObj, ByRef key As Variant, ByRef v As Variant)
    Dim ks As String
    ks = KeyStr(key)
    d.SetH ks, HashStr(ks), key, v
End Sub

Public Function DictDel(ByVal d As PyObj, ByRef key As Variant) As Boolean
    Dim ks As String
    ks = KeyStr(key)
    DictDel = d.DelH(ks, HashStr(ks))
End Function

Public Sub SetAdd(ByVal st As PyObj, ByRef key As Variant)
    Dim e As Variant
    Dim ks As String
    ks = KeyStr(key)
    If st.FindH(ks, HashStr(ks)) < 0 Then st.SetH ks, HashStr(ks), key, e
End Sub

Public Sub DictSetStr(ByVal d As PyObj, ByVal k As String, ByRef v As Variant)
    Dim ks As String
    ks = "s" & k
    d.SetH ks, HashStr(ks), k, v
End Sub

Public Function DictKeysList(ByVal d As PyObj) As PyObj
    Dim r As PyObj, i As Long, x As Variant
    Set r = NewList()
    r.Reserve d.DCount
    For i = 0 To d.EntryCount - 1
        If d.LiveAt(i) Then
            d.KeyInto i, x
            r.Push x
        End If
    Next
    Set DictKeysList = r
End Function

Public Function CopyDict(ByVal d As PyObj, ByVal kind As Long) As PyObj
    Dim r As PyObj, i As Long, k As Variant, v As Variant
    Set r = NewObj(kind)
    For i = 0 To d.EntryCount - 1
        If d.LiveAt(i) Then
            d.KeyInto i, k
            d.ValInto i, v
            r.SetH d.KSAt(i), d.HAt(i), k, v
        End If
    Next
    Set CopyDict = r
End Function

' string-keyed tables (scopes, attribute dictionaries)
Public Sub NameSet(ByVal d As PyObj, ByVal nm As String, ByRef v As Variant)
    d.SetH nm, HashStr(nm), nm, v
End Sub

Public Function NameGet(ByVal d As PyObj, ByVal nm As String, ByRef out As Variant) As Boolean
    Dim i As Long
    i = d.FindH(nm, HashStr(nm))
    If i >= 0 Then
        d.ValInto i, out
        NameGet = True
    End If
End Function

' keyword argument lookup
Public Function KwGet(ByVal kw As PyObj, ByVal nm As String, ByRef out As Variant) As Boolean
    Dim ks As String, i As Long
    If kw Is Nothing Then Exit Function
    ks = "s" & nm
    i = kw.FindH(ks, HashStr(ks))
    If i >= 0 Then
        kw.ValInto i, out
        KwGet = True
    End If
End Function

' ===========================================================================
'  type predicates
' ===========================================================================
Public Function IsNum(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbLong, vbDouble, vbInteger, vbBoolean, vbDecimal, vbSingle, vbCurrency, vbByte
        IsNum = True
    End Select
End Function

Public Function IsIntType(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbLong, vbInteger, vbBoolean, vbDecimal, vbByte
        IsIntType = True
    End Select
End Function

Public Function IsFloatType(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbDouble, vbSingle, vbCurrency
        IsFloatType = True
    End Select
End Function

Public Function IsKind(ByRef v As Variant, ByVal kind As Long) As Boolean
    Dim o As PyObj
    If IsObject(v) Then
        Set o = v
        IsKind = (o.Kind = kind)
    End If
End Function

Public Function IsStrV(ByRef v As Variant) As Boolean
    IsStrV = (VarType(v) = vbString)
End Function

Public Function TypeNameOf(ByRef v As Variant) As String
    Dim o As PyObj
    Select Case VarType(v)
    Case vbEmpty, vbNull: TypeNameOf = "NoneType"
    Case vbBoolean: TypeNameOf = "bool"
    Case vbLong, vbInteger, vbByte, vbDecimal: TypeNameOf = "int"
    Case vbDouble, vbSingle, vbCurrency: TypeNameOf = "float"
    Case vbString: TypeNameOf = "str"
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case OK_LIST: TypeNameOf = "list"
        Case OK_TUPLE: TypeNameOf = "tuple"
        Case OK_DICT
            If o.Flags = 2 Then
                TypeNameOf = "Counter"
            ElseIf o.Flags = 1 Then
                TypeNameOf = "defaultdict"
            Else
                TypeNameOf = "dict"
            End If
        Case OK_SET: TypeNameOf = "set"
        Case OK_FUNC: TypeNameOf = "function"
        Case OK_CLASS: TypeNameOf = "type"
        Case OK_INST: TypeNameOf = o.Cls.PName
        Case OK_BOUND: TypeNameOf = "method"
        Case OK_BUILTIN
            If o.Flags = 1 Then
                TypeNameOf = "type"
            Else
                TypeNameOf = "builtin_function_or_method"
            End If
        Case OK_BMETHOD: TypeNameOf = "builtin_function_or_method"
        Case OK_MODULE: TypeNameOf = "module"
        Case OK_RANGE: TypeNameOf = "range"
        Case OK_ITER: TypeNameOf = "iterator"
        Case OK_SUPER: TypeNameOf = "super"
        Case OK_PROP: TypeNameOf = "property"
        Case Else: TypeNameOf = "object"
        End Select
    Case vbDate: TypeNameOf = "datetime"
    Case Else: TypeNameOf = "object"
    End Select
End Function

' ===========================================================================
'  integers (Long with automatic promotion to Decimal)
' ===========================================================================
Public Function IntOf(ByRef v As Variant) As Variant
    Select Case VarType(v)
    Case vbBoolean
        If v Then
            IntOf = 1&
        Else
            IntOf = 0&
        End If
    Case vbInteger, vbByte
        IntOf = CLng(v)
    Case Else
        IntOf = v
    End Select
End Function

Public Function NormBig(ByVal d As Variant) As Variant
    If VarType(d) = vbLong Then
        NormBig = d
    ElseIf d >= -2147483648# And d <= 2147483647# Then
        NormBig = CLng(d)
    Else
        NormBig = d
    End If
End Function

Public Function DecFromString(ByVal s As String) As Variant
    On Error GoTo Ovf
    DecFromString = CDec(s)
    Exit Function
Ovf:
    ThrowErr "OverflowError", "integer too large (more than 28 digits)"
End Function

Public Function IntFromDouble(ByVal d As Double) As Variant
    If d >= -2147483648# And d <= 2147483647# Then
        IntFromDouble = CLng(d)
    Else
        IntFromDouble = DecFromString(Format$(d, "0"))
    End If
End Function

Private Function DecArith(ByVal op As Long, ByVal a As Variant, ByVal b As Variant) As Variant
    On Error GoTo Ovf
    Select Case op
    Case OP_ADD: DecArith = CDec(a) + CDec(b)
    Case OP_SUB: DecArith = CDec(a) - CDec(b)
    Case OP_MUL: DecArith = CDec(a) * CDec(b)
    End Select
    Exit Function
Ovf:
    ThrowErr "OverflowError", "integer result too large (more than 28 digits)"
End Function

Public Function IntAdd(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    a = IntOf(a)
    b = IntOf(b)
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) + CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            IntAdd = CLng(d)
            Exit Function
        End If
    End If
    IntAdd = NormBig(DecArith(OP_ADD, a, b))
End Function

Public Function IntSub(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    a = IntOf(a)
    b = IntOf(b)
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) - CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            IntSub = CLng(d)
            Exit Function
        End If
    End If
    IntSub = NormBig(DecArith(OP_SUB, a, b))
End Function

Public Function IntMul(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim d As Double
    a = IntOf(a)
    b = IntOf(b)
    If VarType(a) <> vbDecimal And VarType(b) <> vbDecimal Then
        d = CDbl(a) * CDbl(b)
        If d >= -2147483648# And d <= 2147483647# Then
            IntMul = CLng(d)
            Exit Function
        End If
    End If
    IntMul = NormBig(DecArith(OP_MUL, a, b))
End Function

Public Function IntNeg(ByVal a As Variant) As Variant
    a = IntOf(a)
    If VarType(a) = vbLong Then
        If a = -2147483647 - 1 Then
            IntNeg = DecFromString("2147483648")
        Else
            IntNeg = CLng(-a)
        End If
    Else
        IntNeg = NormBig(-a)
    End If
End Function

' floor division and modulo with Python semantics
Public Sub IntDivMod(ByVal a As Variant, ByVal b As Variant, ByRef q As Variant, ByRef r As Variant)
    Dim dq As Double, dr As Double, da As Double, db As Double
    Dim xq As Variant, xr As Variant
    a = IntOf(a)
    b = IntOf(b)
    If b = 0 Then ThrowErr "ZeroDivisionError", "integer division or modulo by zero"
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
        q = IntFromDouble(dq)
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
    q = NormBig(xq)
    r = NormBig(xr)
    Exit Sub
Ovf:
    ThrowErr "OverflowError", "integer division overflow"
End Sub

Public Function IntPow(ByVal a As Variant, ByVal e As Variant) As Variant
    Dim r As Variant, bs As Variant, n As Long
    a = IntOf(a)
    e = IntOf(e)
    If VarType(e) = vbDecimal Then ThrowErr "OverflowError", "exponent too large"
    n = e
    r = 1&
    bs = a
    Do While n > 0
        If (n And 1) = 1 Then r = IntMul(r, bs)
        n = n \ 2
        If n > 0 Then bs = IntMul(bs, bs)
    Loop
    IntPow = r
End Function

Public Function IntToLong(ByRef v As Variant, ByVal what As String) As Long
    Dim x As Variant
    If Not IsIntType(v) Then
        If IsFloatType(v) Then ThrowErr "TypeError", "'float' object cannot be interpreted as an integer"
        ThrowErr "TypeError", what & " must be an integer, not '" & TypeNameOf(v) & "'"
    End If
    x = IntOf(v)
    If VarType(x) = vbDecimal Then ThrowErr "OverflowError", "Python int too large to convert to a 32-bit integer"
    IntToLong = x
End Function

Public Function NumToDbl(ByRef v As Variant) As Double
    If VarType(v) = vbBoolean Then
        If v Then
            NumToDbl = 1#
        Else
            NumToDbl = 0#
        End If
    Else
        NumToDbl = CDbl(v)
    End If
End Function

' parse an integer literal / int() argument; ok=False when invalid
Public Function ParseIntStr(ByVal s As String, ByVal bs As Long, ByRef ok As Boolean) As Variant
    Dim i As Long, n As Long, c As String, d As Long, neg As Boolean, r As Variant, k As Long
    ok = False
    s = Replace(Trim$(s), "_", vbNullString)
    n = Len(s)
    If n = 0 Then Exit Function
    i = 1
    c = Mid$(s, 1, 1)
    If c = "-" Or c = "+" Then
        neg = (c = "-")
        i = 2
    End If
    If bs = 16 Or bs = 8 Or bs = 2 Or bs = 0 Then
        If LCase$(Mid$(s, i, 2)) = "0x" And (bs = 16 Or bs = 0) Then
            i = i + 2: bs = 16
        ElseIf LCase$(Mid$(s, i, 2)) = "0o" And (bs = 8 Or bs = 0) Then
            i = i + 2: bs = 8
        ElseIf LCase$(Mid$(s, i, 2)) = "0b" And (bs = 2 Or bs = 0) Then
            i = i + 2: bs = 2
        End If
    End If
    If bs = 0 Then bs = 10
    If i > n Then Exit Function
    If bs = 10 And n - i + 1 <= 9 Then
        For k = i To n
            c = Mid$(s, k, 1)
            If c < "0" Or c > "9" Then Exit Function
        Next
        r = CLng(Mid$(s, i))
        If neg Then r = CLng(-r)
        ParseIntStr = r
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
        r = IntAdd(IntMul(r, bs), d)
    Next
    If neg Then r = IntNeg(r)
    ParseIntStr = r
    ok = True
End Function

' validate and parse a float literal string (Python syntax)
Public Function ParseFloatStr(ByVal s As String, ByRef ok As Boolean) As Double
    Dim i As Long, n As Long, c As String, digits As Long, seenDot As Boolean, seenE As Boolean
    ok = False
    s = Replace(Trim$(s), "_", vbNullString)
    n = Len(s)
    If n = 0 Then Exit Function
    i = 1
    c = Mid$(s, 1, 1)
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
    ParseFloatStr = Val(s)
    ok = True
End Function

' ===========================================================================
'  exceptions
' ===========================================================================
Public Function ExcClass(ByVal nm As String) As PyObj
    Dim i As Long, v As Variant
    i = gBuiltins.FindH(nm, HashStr(nm))
    If i < 0 Then i = gBuiltins.FindH("RuntimeError", HashStr("RuntimeError"))
    gBuiltins.ValInto i, v
    Set ExcClass = v
End Function

Public Function MakeExcV(ByVal clsName As String, ByRef arg As Variant) As PyObj
    Dim e As PyObj, t As PyObj
    Set e = NewObj(OK_INST)
    Set e.Cls = ExcClass(clsName)
    Set t = NewObj(OK_TUPLE)
    t.Push arg
    NameSet e, "args", t
    Set MakeExcV = e
End Function

Public Function MakeExc(ByVal clsName As String, ByVal msg As String) As PyObj
    Set MakeExc = MakeExcV(clsName, msg)
End Function

Public Sub RaiseExc(ByVal exc As PyObj)
    Set gExc = exc
    If exc.I3 = 0 Then
        exc.I1 = gFrameLine(gDepth)
        exc.I2 = gDepth
        exc.I3 = 1
    End If
    Err.Raise PYERR, "PyVBA", "Python exception"
End Sub

Public Sub ThrowErr(ByVal clsName As String, ByVal msg As String)
    RaiseExc MakeExc(clsName, msg)
End Sub

' convert the current VBA error (Err) into a Python exception object
Public Function CaptureExc() As PyObj
    Dim e As PyObj, num As Long, desc As String
    num = Err.Number
    If num = PYERR Then
        If gExc Is Nothing Then
            Set e = MakeExc("SystemError", "lost exception")
        Else
            Set e = gExc
        End If
    Else
        desc = Err.Description
        Select Case num
        Case 6: Set e = MakeExc("OverflowError", "numeric overflow")
        Case 11: Set e = MakeExc("ZeroDivisionError", "division by zero")
        Case 13: Set e = MakeExc("TypeError", "type mismatch (VBA: " & desc & ")")
        Case 9: Set e = MakeExc("IndexError", "index out of range (VBA: " & desc & ")")
        Case 28: Set e = MakeExc("RecursionError", "maximum recursion depth exceeded (out of stack space)")
        Case 7: Set e = MakeExc("MemoryError", "out of memory")
        Case 5: Set e = MakeExc("ValueError", "invalid argument (VBA: " & desc & ")")
        Case 18: Set e = MakeExc("KeyboardInterrupt", "interrupted by user")
        Case 1004: Set e = MakeExc("RuntimeError", "Excel error: " & desc)
        Case 438: Set e = MakeExc("AttributeError", "VBA: " & desc)
        Case Else: Set e = MakeExc("SystemError", "VBA error " & num & ": " & desc)
        End Select
        If e.I3 = 0 Then
            e.I1 = gFrameLine(gDepth)
            e.I2 = gDepth
            e.I3 = 1
        End If
    End If
    Set gExc = Nothing
    Set CaptureExc = e
End Function

Public Function IsSubclassOf(ByVal c As PyObj, ByVal target As PyObj) As Boolean
    Dim i As Long, b As Variant
    If c Is target Then
        IsSubclassOf = True
        Exit Function
    End If
    If target Is gObjectClass Then
        IsSubclassOf = True
        Exit Function
    End If
    For i = 0 To c.Count - 1
        c.GetInto i, b
        If IsSubclassOf(b, target) Then
            IsSubclassOf = True
            Exit Function
        End If
    Next
End Function

Public Function IsExcInstance(ByRef v As Variant) As Boolean
    Dim o As PyObj
    If IsObject(v) Then
        Set o = v
        If o.Kind = OK_INST Then IsExcInstance = IsSubclassOf(o.Cls, ExcClass("BaseException"))
    End If
End Function

Public Function ExcMatches(ByVal exc As PyObj, ByRef tv As Variant) As Boolean
    Dim o As PyObj, i As Long, x As Variant
    If Not IsObject(tv) Then ThrowErr "TypeError", "catching classes that do not inherit from BaseException is not allowed"
    Set o = tv
    If o.Kind = OK_TUPLE Then
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If ExcMatches(exc, x) Then
                ExcMatches = True
                Exit Function
            End If
        Next
        Exit Function
    End If
    If o.Kind <> OK_CLASS Then ThrowErr "TypeError", "catching classes that do not inherit from BaseException is not allowed"
    ExcMatches = IsSubclassOf(exc.Cls, o)
End Function

' str(exception) for built-in exception classes
Public Function ExcMessage(ByVal e As PyObj) As String
    Dim v As Variant, t As PyObj, x As Variant
    If Not NameGet(e, "args", v) Then Exit Function
    If Not IsKind(v, OK_TUPLE) Then
        ExcMessage = PyStr(v)
        Exit Function
    End If
    Set t = v
    If t.Count = 0 Then Exit Function
    If t.Count = 1 Then
        t.GetInto 0, x
        If IsSubclassOf(e.Cls, ExcClass("KeyError")) Then
            ExcMessage = PyRepr(x)
        Else
            ExcMessage = PyStr(x)
        End If
        Exit Function
    End If
    ExcMessage = PyRepr(t)
End Function

' "ValueError: message"
Public Function ExcLine(ByVal e As PyObj) As String
    Dim s As String, v As Variant
    Set v = e
    s = PyStr(v)
    If Len(s) = 0 Then
        ExcLine = e.Cls.PName
    Else
        ExcLine = e.Cls.PName & ": " & s
    End If
End Function

' ===========================================================================
'  truth, equality, ordering, membership, identity
' ===========================================================================
Public Function Truthy(ByRef v As Variant) As Boolean
    Select Case VarType(v)
    Case vbBoolean: Truthy = v
    Case vbLong, vbInteger, vbDouble, vbDecimal, vbSingle, vbCurrency, vbByte: Truthy = (v <> 0)
    Case vbString: Truthy = (Len(v) > 0)
    Case vbEmpty, vbNull: Truthy = False
    Case vbObject: Truthy = ObjTruthy(v)
    Case Else: Truthy = True
    End Select
End Function

Private Function ObjTruthy(ByRef v As Variant) As Boolean
    Dim o As PyObj, r As Variant
    Set o = v
    Select Case o.Kind
    Case OK_LIST, OK_TUPLE: ObjTruthy = (o.Count > 0)
    Case OK_DICT, OK_SET: ObjTruthy = (o.DCount > 0)
    Case OK_RANGE: ObjTruthy = (RangeLen(o) > 0)
    Case OK_INST
        If CallDunder0(v, "__bool__", r) Then
            ObjTruthy = Truthy(r)
        ElseIf CallDunder0(v, "__len__", r) Then
            ObjTruthy = Truthy(r)
        Else
            ObjTruthy = True
        End If
    Case Else: ObjTruthy = True
    End Select
End Function

Public Function NumCmp(ByRef a As Variant, ByRef b As Variant) As Long
    Dim x As Variant, y As Variant, dx As Double, dy As Double
    x = IntOf(a)
    y = IntOf(b)
    If IsFloatType(x) Or IsFloatType(y) Then
        dx = CDbl(x)
        dy = CDbl(y)
        If dx < dy Then
            NumCmp = -1
        ElseIf dx > dy Then
            NumCmp = 1
        End If
    Else
        If x < y Then
            NumCmp = -1
        ElseIf x > y Then
            NumCmp = 1
        End If
    End If
End Function

Public Function IsNotImpl(ByRef v As Variant) As Boolean
    If IsObject(v) Then IsNotImpl = (v Is gNotImpl)
End Function

Public Function PyEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim ta As Long, tb As Long
    ta = VarType(a)
    tb = VarType(b)
    If ta = vbString Then
        If tb = vbString Then
            PyEq = (a = b)
            Exit Function
        End If
    End If
    If ta = vbObject Or tb = vbObject Then
        PyEq = ObjEq(a, b)
        Exit Function
    End If
    If ta = vbString Or tb = vbString Then Exit Function
    If IsNum(a) Then
        If IsNum(b) Then PyEq = (NumCmp(a, b) = 0)
        Exit Function
    End If
    If ta = vbEmpty Or ta = vbNull Then PyEq = (tb = vbEmpty Or tb = vbNull)
End Function

Private Function ObjEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim oa As PyObj, ob As PyObj, r As Variant, i As Long, x As Variant, y As Variant, e As Long
    If IsObject(a) Then
        Set oa = a
        If IsObject(b) Then
            Set ob = b
            If oa Is ob Then
                ObjEq = True
                Exit Function
            End If
        End If
        If oa.Kind = OK_INST Then
            If CallDunder1(a, "__eq__", b, r) Then
                If Not IsNotImpl(r) Then
                    ObjEq = Truthy(r)
                    Exit Function
                End If
            End If
        End If
    End If
    If IsObject(b) Then
        Set ob = b
        If ob.Kind = OK_INST Then
            If CallDunder1(b, "__eq__", a, r) Then
                If Not IsNotImpl(r) Then
                    ObjEq = Truthy(r)
                    Exit Function
                End If
            End If
        End If
    End If
    If Not IsObject(a) Then Exit Function
    If Not IsObject(b) Then Exit Function
    Select Case oa.Kind
    Case OK_LIST, OK_TUPLE
        If ob.Kind <> oa.Kind Then Exit Function
        If oa.Count <> ob.Count Then Exit Function
        For i = 0 To oa.Count - 1
            oa.GetInto i, x
            ob.GetInto i, y
            If Not PyEq(x, y) Then Exit Function
        Next
        ObjEq = True
    Case OK_DICT, OK_SET
        If ob.Kind <> oa.Kind Then Exit Function
        If oa.DCount <> ob.DCount Then Exit Function
        For i = 0 To oa.EntryCount - 1
            If oa.LiveAt(i) Then
                e = ob.FindH(oa.KSAt(i), oa.HAt(i))
                If e < 0 Then Exit Function
                If oa.Kind = OK_DICT Then
                    oa.ValInto i, x
                    ob.ValInto e, y
                    If Not PyEq(x, y) Then Exit Function
                End If
            End If
        Next
        ObjEq = True
    Case OK_RANGE
        If ob.Kind <> OK_RANGE Then Exit Function
        If RangeLen(oa) <> RangeLen(ob) Then Exit Function
        If RangeLen(oa) = 0 Then
            ObjEq = True
        ElseIf oa.I1 = ob.I1 Then
            ObjEq = (RangeLen(oa) = 1 Or oa.I3 = ob.I3)
        End If
    End Select
End Function

Public Function OpSym(ByVal op As Long) As String
    Select Case op
    Case OP_ADD: OpSym = "+"
    Case OP_SUB: OpSym = "-"
    Case OP_MUL: OpSym = "*"
    Case OP_DIV: OpSym = "/"
    Case OP_FLOORDIV: OpSym = "//"
    Case OP_MOD: OpSym = "%"
    Case OP_POW: OpSym = "**"
    Case OP_LSHIFT: OpSym = "<<"
    Case OP_RSHIFT: OpSym = ">>"
    Case OP_BITAND: OpSym = "&"
    Case OP_BITOR: OpSym = "|"
    Case OP_BITXOR: OpSym = "^"
    Case OP_MATMUL: OpSym = "@"
    Case OP_LT: OpSym = "<"
    Case OP_LE: OpSym = "<="
    Case OP_GT: OpSym = ">"
    Case OP_GE: OpSym = ">="
    Case OP_EQ: OpSym = "=="
    Case OP_NE: OpSym = "!="
    Case Else: OpSym = "?"
    End Select
End Function

Private Function CmpResult(ByVal op As Long, ByVal c As Long) As Boolean
    Select Case op
    Case OP_LT: CmpResult = (c < 0)
    Case OP_LE: CmpResult = (c <= 0)
    Case OP_GT: CmpResult = (c > 0)
    Case OP_GE: CmpResult = (c >= 0)
    End Select
End Function

' <, <=, >, >=
Public Function OrderCmp(ByVal op As Long, ByRef a As Variant, ByRef b As Variant) As Boolean
    Dim oa As PyObj, ob As PyObj, i As Long, n As Long, x As Variant, y As Variant, r As Variant
    Dim nm As String, rnm As String, c As Long
    If IsNum(a) Then
        If IsNum(b) Then
            OrderCmp = CmpResult(op, NumCmp(a, b))
            Exit Function
        End If
    End If
    If VarType(a) = vbString Then
        If VarType(b) = vbString Then
            OrderCmp = CmpResult(op, StrComp(a, b, vbBinaryCompare))
            Exit Function
        End If
    End If
    Select Case op
    Case OP_LT: nm = "__lt__": rnm = "__gt__"
    Case OP_LE: nm = "__le__": rnm = "__ge__"
    Case OP_GT: nm = "__gt__": rnm = "__lt__"
    Case OP_GE: nm = "__ge__": rnm = "__le__"
    End Select
    If IsKind(a, OK_INST) Then
        If CallDunder1(a, nm, b, r) Then
            If Not IsNotImpl(r) Then
                OrderCmp = Truthy(r)
                Exit Function
            End If
        End If
    End If
    If IsKind(b, OK_INST) Then
        If CallDunder1(b, rnm, a, r) Then
            If Not IsNotImpl(r) Then
                OrderCmp = Truthy(r)
                Exit Function
            End If
        End If
    End If
    If IsObject(a) And IsObject(b) Then
        Set oa = a
        Set ob = b
        If (oa.Kind = OK_LIST Or oa.Kind = OK_TUPLE) And oa.Kind = ob.Kind Then
            n = oa.Count
            If ob.Count < n Then n = ob.Count
            For i = 0 To n - 1
                oa.GetInto i, x
                ob.GetInto i, y
                If Not PyEq(x, y) Then
                    OrderCmp = OrderCmp(op, x, y)
                    Exit Function
                End If
            Next
            OrderCmp = CmpResult(op, Sgn(oa.Count - ob.Count))
            Exit Function
        End If
        If oa.Kind = OK_SET And ob.Kind = OK_SET Then
            c = 0
            Select Case op
            Case OP_LE: OrderCmp = SetSubset(oa, ob)
            Case OP_LT: OrderCmp = SetSubset(oa, ob) And oa.DCount < ob.DCount
            Case OP_GE: OrderCmp = SetSubset(ob, oa)
            Case OP_GT: OrderCmp = SetSubset(ob, oa) And ob.DCount < oa.DCount
            End Select
            Exit Function
        End If
    End If
    ThrowErr "TypeError", "'" & OpSym(op) & "' not supported between instances of '" & TypeNameOf(a) & "' and '" & TypeNameOf(b) & "'"
End Function

Public Function SetSubset(ByVal a As PyObj, ByVal b As PyObj) As Boolean
    Dim i As Long
    For i = 0 To a.EntryCount - 1
        If a.LiveAt(i) Then
            If b.FindH(a.KSAt(i), a.HAt(i)) < 0 Then Exit Function
        End If
    Next
    SetSubset = True
End Function

Public Function PyLess(ByRef a As Variant, ByRef b As Variant) As Boolean
    If VarType(a) = vbLong Then
        If VarType(b) = vbLong Then
            PyLess = (a < b)
            Exit Function
        End If
    End If
    PyLess = OrderCmp(OP_LT, a, b)
End Function

Public Function Contains(ByRef c As Variant, ByRef x As Variant) As Boolean
    Dim o As PyObj, i As Long, y As Variant, r As Variant, it As PyObj, xv As Variant
    If VarType(c) = vbString Then
        If VarType(x) <> vbString Then ThrowErr "TypeError", "'in <string>' requires string as left operand, not " & TypeNameOf(x)
        If Len(x) = 0 Then
            Contains = True
        Else
            Contains = (InStr(1, c, x, vbBinaryCompare) > 0)
        End If
        Exit Function
    End If
    If IsObject(c) Then
        Set o = c
        Select Case o.Kind
        Case OK_LIST, OK_TUPLE
            For i = 0 To o.Count - 1
                o.GetInto i, y
                If PyEq(y, x) Then
                    Contains = True
                    Exit Function
                End If
            Next
            Exit Function
        Case OK_DICT, OK_SET
            Contains = (DictFind(o, x) >= 0)
            Exit Function
        Case OK_RANGE
            If IsNum(x) Then
                xv = IntOf(x)
                If IsFloatType(xv) Then
                    If xv <> Int(xv) Then Exit Function
                End If
                If o.I3 > 0 Then
                    If xv < o.I1 Or xv >= o.I2 Then Exit Function
                Else
                    If xv > o.I1 Or xv <= o.I2 Then Exit Function
                End If
                Contains = ((CDbl(xv) - o.I1) - Int((CDbl(xv) - o.I1) / o.I3) * o.I3 = 0)
            End If
            Exit Function
        Case OK_INST
            If CallDunder1(c, "__contains__", x, r) Then
                Contains = Truthy(r)
                Exit Function
            End If
            Set it = GetIter(c)
            Do While IterNext(it, y)
                If PyEq(y, x) Then
                    Contains = True
                    Exit Function
                End If
            Loop
            Exit Function
        Case OK_ITER
            Set it = o
            Do While IterNext(it, y)
                If PyEq(y, x) Then
                    Contains = True
                    Exit Function
                End If
            Loop
            Exit Function
        End Select
    End If
    ThrowErr "TypeError", "argument of type '" & TypeNameOf(c) & "' is not iterable"
End Function

Public Function IsSame(ByRef a As Variant, ByRef b As Variant) As Boolean
    If IsObject(a) Then
        If IsObject(b) Then IsSame = (a Is b)
        Exit Function
    End If
    If IsObject(b) Then Exit Function
    If VarType(a) <> VarType(b) Then Exit Function
    Select Case VarType(a)
    Case vbEmpty, vbNull: IsSame = True
    Case vbBoolean, vbLong, vbString, vbDouble, vbDecimal: IsSame = (a = b)
    End Select
End Function

Public Function CompareOp(ByVal op As Long, ByRef a As Variant, ByRef b As Variant) As Boolean
    If VarType(a) = vbLong Then
        If VarType(b) = vbLong Then
            Select Case op
            Case OP_LT: CompareOp = (a < b): Exit Function
            Case OP_LE: CompareOp = (a <= b): Exit Function
            Case OP_GT: CompareOp = (a > b): Exit Function
            Case OP_GE: CompareOp = (a >= b): Exit Function
            Case OP_EQ: CompareOp = (a = b): Exit Function
            Case OP_NE: CompareOp = (a <> b): Exit Function
            End Select
        End If
    End If
    Select Case op
    Case OP_EQ: CompareOp = PyEq(a, b)
    Case OP_NE
        If IsKind(a, OK_INST) Then
            Dim r As Variant
            If CallDunder1(a, "__ne__", b, r) Then
                If Not IsNotImpl(r) Then
                    CompareOp = Truthy(r)
                    Exit Function
                End If
            End If
        End If
        CompareOp = Not PyEq(a, b)
    Case OP_LT, OP_LE, OP_GT, OP_GE: CompareOp = OrderCmp(op, a, b)
    Case OP_IN: CompareOp = Contains(b, a)
    Case OP_NOTIN: CompareOp = Not Contains(b, a)
    Case OP_IS: CompareOp = IsSame(a, b)
    Case OP_ISNOT: CompareOp = Not IsSame(a, b)
    End Select
End Function

' ===========================================================================
'  repr / str
' ===========================================================================
Public Function FloatRepr(ByVal d As Double) As String
    Dim s As String, neg As Boolean, p As Long, ex As Long, mant As String, digits As String
    Dim pointPos As Long, e10 As Long, r As String
    If d = 0 Then
        FloatRepr = "0.0"
        Exit Function
    End If
    If d >= PY_INF Then
        FloatRepr = "inf"
        Exit Function
    ElseIf d <= -PY_INF Then
        FloatRepr = "-inf"
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
    If e10 >= 16 Or e10 < -4 Then
        r = Left$(digits, 1)
        If Len(digits) > 1 Then r = r & "." & Mid$(digits, 2)
        If e10 < 0 Then
            r = r & "e-"
        Else
            r = r & "e+"
        End If
        r = r & Format$(Abs(e10), "00")
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
    FloatRepr = r
End Function

Public Function UCode(ByVal c As String) As Long
    Dim k As Long
    k = AscW(c)
    If k < 0 Then k = k + 65536
    UCode = k
End Function

Public Function StrRepr(ByVal s As String) As String
    Dim q As String, i As Long, n As Long, c As String, k As Long, parts() As String, plain As Boolean
    If InStr(1, s, "'") > 0 And InStr(1, s, """") = 0 Then q = """" Else q = "'"
    n = Len(s)
    plain = (InStr(1, s, "\") = 0 And InStr(1, s, q) = 0)
    If plain Then
        For i = 1 To n
            k = AscW(Mid$(s, i, 1))
            If k >= 0 And k < 32 Then
                plain = False
                Exit For
            End If
            If k = 127 Then
                plain = False
                Exit For
            End If
        Next
    End If
    If plain Then
        StrRepr = q & s & q
        Exit Function
    End If
    ReDim parts(0 To n)
    For i = 1 To n
        c = Mid$(s, i, 1)
        k = UCode(c)
        If k < 32 Or k = 127 Then
            Select Case k
            Case 10: c = "\n"
            Case 13: c = "\r"
            Case 9: c = "\t"
            Case Else: c = "\x" & LCase$(Right$("0" & Hex$(k), 2))
            End Select
        ElseIf c = "\" Then
            c = "\\"
        ElseIf c = q Then
            c = "\" & q
        End If
        parts(i - 1) = c
    Next
    StrRepr = q & Join(parts, vbNullString) & q
End Function

Public Function HexId(ByVal id As Long) As String
    HexId = "0x" & LCase$(Right$("0000000" & Hex$(id * 16 + &H1000000), 8))
End Function

Public Function PyRepr(ByRef v As Variant) As String
    Select Case VarType(v)
    Case vbEmpty, vbNull: PyRepr = "None"
    Case vbBoolean
        If v Then
            PyRepr = "True"
        Else
            PyRepr = "False"
        End If
    Case vbLong, vbInteger, vbByte, vbDecimal: PyRepr = CStr(v)
    Case vbDouble, vbSingle, vbCurrency: PyRepr = FloatRepr(CDbl(v))
    Case vbString: PyRepr = StrRepr(v)
    Case vbObject: PyRepr = ObjRepr(v)
    Case vbDate: PyRepr = "'" & Format$(v, "yyyy-mm-dd hh:nn:ss") & "'"
    Case Else: PyRepr = "<?>"
    End Select
End Function

Private Function ObjRepr(ByRef v As Variant) As String
    Dim o As PyObj, parts() As String, i As Long, x As Variant, y As Variant, r As Variant, k As Long
    Set o = v
    Select Case o.Kind
    Case OK_LIST, OK_TUPLE
        If o.Count = 0 Then
            If o.Kind = OK_LIST Then
                ObjRepr = "[]"
            Else
                ObjRepr = "()"
            End If
            Exit Function
        End If
        ReDim parts(0 To o.Count - 1)
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If IsObject(x) Then
                If x Is o Then
                    parts(i) = "[...]"
                Else
                    parts(i) = PyRepr(x)
                End If
            Else
                parts(i) = PyRepr(x)
            End If
        Next
        If o.Kind = OK_LIST Then
            ObjRepr = "[" & Join(parts, ", ") & "]"
        ElseIf o.Count = 1 Then
            ObjRepr = "(" & parts(0) & ",)"
        Else
            ObjRepr = "(" & Join(parts, ", ") & ")"
        End If
    Case OK_DICT, OK_SET
        If o.DCount = 0 Then
            If o.Kind = OK_SET Then
                ObjRepr = "set()"
            ElseIf o.Flags = 2 Then
                ObjRepr = "Counter()"
            Else
                ObjRepr = "{}"
            End If
            Exit Function
        End If
        ReDim parts(0 To o.DCount - 1)
        k = 0
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.KeyInto i, x
                If o.Kind = OK_DICT Then
                    o.ValInto i, y
                    parts(k) = PyRepr(x) & ": " & PyRepr(y)
                Else
                    parts(k) = PyRepr(x)
                End If
                k = k + 1
            End If
        Next
        If o.Flags = 2 Then
            ObjRepr = "Counter(" & CounterRepr(o) & ")"
        Else
            ObjRepr = "{" & Join(parts, ", ") & "}"
        End If
    Case OK_FUNC
        ObjRepr = "<function " & o.PName & " at " & HexId(o.Id) & ">"
    Case OK_CLASS
        If (o.Flags And 1) <> 0 Then
            ObjRepr = "<class '" & o.PName & "'>"
        Else
            ObjRepr = "<class '__main__." & o.PName & "'>"
        End If
    Case OK_INST
        If CallDunder0(v, "__repr__", r) Then
            If VarType(r) <> vbString Then ThrowErr "TypeError", "__repr__ returned non-string (type " & TypeNameOf(r) & ")"
            ObjRepr = r
        Else
            ObjRepr = "<__main__." & o.Cls.PName & " object at " & HexId(o.Id) & ">"
        End If
    Case OK_BOUND
        SetV x, o.FnV
        ObjRepr = "<bound method " & FuncNameOf(x) & " of " & PyRepr(o.SelfV) & ">"
    Case OK_BUILTIN
        If o.Flags = 1 Then
            ObjRepr = "<class '" & o.PName & "'>"
        Else
            ObjRepr = "<built-in function " & o.PName & ">"
        End If
    Case OK_BMETHOD
        ObjRepr = "<built-in method " & o.PName & " of " & TypeNameOf(o.SelfV) & " object at " & HexId(o.Id) & ">"
    Case OK_MODULE
        ObjRepr = "<module '" & o.PName & "'>"
    Case OK_RANGE
        If o.I3 = 1 Then
            ObjRepr = "range(" & CStr(o.I1) & ", " & CStr(o.I2) & ")"
        Else
            ObjRepr = "range(" & CStr(o.I1) & ", " & CStr(o.I2) & ", " & CStr(o.I3) & ")"
        End If
    Case OK_SENTINEL
        ObjRepr = o.PName
    Case Else
        ObjRepr = "<" & TypeNameOf(v) & " object at " & HexId(o.Id) & ">"
    End Select
End Function

Public Function FuncNameOf(ByRef f As Variant) As String
    Dim o As PyObj
    If IsObject(f) Then
        Set o = f
        If o.Kind = OK_FUNC And Not (o.Cls Is Nothing) Then
            FuncNameOf = o.Cls.PName & "." & o.PName
        Else
            FuncNameOf = o.PName
        End If
    End If
End Function

Public Function PyStr(ByRef v As Variant) As String
    Dim o As PyObj, r As Variant
    If VarType(v) = vbString Then
        PyStr = v
        Exit Function
    End If
    If IsObject(v) Then
        Set o = v
        If o.Kind = OK_INST Then
            If CallDunder0(v, "__str__", r) Then
                If VarType(r) <> vbString Then ThrowErr "TypeError", "__str__ returned non-string (type " & TypeNameOf(r) & ")"
                PyStr = r
                Exit Function
            End If
        End If
    End If
    PyStr = PyRepr(v)
End Function

' ===========================================================================
'  iteration
' ===========================================================================
Public Function GetIter(ByRef v As Variant) As PyObj
    Dim it As PyObj, o As PyObj, r As Variant, ro As PyObj
    Set it = NewObj(OK_ITER)
    If VarType(v) = vbString Then
        it.Code = IT_STR
        it.SetStr v
        it.I1 = 0
        it.I2 = Len(v)
        Set GetIter = it
        Exit Function
    End If
    If IsObject(v) Then
        Set o = v
        Select Case o.Kind
        Case OK_LIST, OK_TUPLE
            it.Code = IT_SEQ
            Set it.Ref = o
        Case OK_DICT, OK_SET
            it.Code = IT_SEQ
            Set it.Ref = DictKeysList(o)
        Case OK_RANGE
            it.Code = IT_RANGE
            it.I1 = o.I1
            it.I2 = o.I2
            it.I3 = o.I3
        Case OK_ITER
            Set GetIter = o
            Exit Function
        Case OK_INST
            If CallDunder0(v, "__iter__", r) Then
                If IsKind(r, OK_INST) Then
                    Set ro = r
                    it.Code = IT_USER
                    Set it.Ref = ro
                    Set GetIter = it
                    Exit Function
                End If
                Set GetIter = GetIter(r)
                Exit Function
            End If
            GoTo NotIter
        Case Else
            GoTo NotIter
        End Select
        Set GetIter = it
        Exit Function
    End If
NotIter:
    ThrowErr "TypeError", "'" & TypeNameOf(v) & "' object is not iterable"
End Function

Public Function IterNext(ByVal it As PyObj, ByRef out As Variant) As Boolean
    If it.Code = IT_USER Then
        IterNext = UserIterNext(it.Ref, out)
    Else
        IterNext = it.IterStep(out)
    End If
End Function

Public Function NewListFrom(ByRef v As Variant) As PyObj
    Dim r As PyObj, o As PyObj, i As Long, x As Variant, it As PyObj, s As String
    Set r = NewObj(OK_LIST)
    If IsObject(v) Then
        Set o = v
        If o.Kind = OK_LIST Or o.Kind = OK_TUPLE Then
            r.Reserve o.Count
            For i = 0 To o.Count - 1
                o.GetInto i, x
                r.Push x
            Next
            Set NewListFrom = r
            Exit Function
        End If
    ElseIf VarType(v) = vbString Then
        s = v
        r.Reserve Len(s)
        For i = 1 To Len(s)
            x = Mid$(s, i, 1)
            r.Push x
        Next
        Set NewListFrom = r
        Exit Function
    End If
    Set it = GetIter(v)
    Do While IterNext(it, x)
        r.Push x
    Loop
    Set NewListFrom = r
End Function

' list/tuple itself, or a new list with the items of any iterable
Public Function SeqOf(ByRef v As Variant) As PyObj
    Dim o As PyObj
    If IsObject(v) Then
        Set o = v
        If o.Kind = OK_LIST Or o.Kind = OK_TUPLE Then
            Set SeqOf = o
            Exit Function
        End If
    End If
    Set SeqOf = NewListFrom(v)
End Function

' ===========================================================================
'  indexing and slicing
' ===========================================================================
Public Function NormIndex(ByRef idx As Variant, ByVal n As Long, ByVal what As String) As Long
    Dim i As Variant
    If Not IsIntType(idx) Then
        ThrowErr "TypeError", what & " indices must be integers or slices, not " & TypeNameOf(idx)
    End If
    i = IntOf(idx)
    If VarType(i) = vbDecimal Then ThrowErr "IndexError", what & " index out of range"
    If i < 0 Then i = i + n
    If i < 0 Or i >= n Then ThrowErr "IndexError", what & " index out of range"
    NormIndex = i
End Function

Public Sub GetItem(ByRef ov As Variant, ByRef idx As Variant, ByRef out As Variant)
    Dim o As PyObj, i As Long
    If VarType(ov) = vbString Then
        i = NormIndex(idx, Len(ov), "string")
        out = Mid$(ov, i + 1, 1)
        Exit Sub
    End If
    If IsObject(ov) Then
        Set o = ov
        Select Case o.Kind
        Case OK_LIST
            i = NormIndex(idx, o.Count, "list")
            o.GetInto i, out
            Exit Sub
        Case OK_TUPLE
            i = NormIndex(idx, o.Count, "tuple")
            o.GetInto i, out
            Exit Sub
        Case OK_DICT
            If DictGet(o, idx, out) Then Exit Sub
            DictMissing o, idx, out
            Exit Sub
        Case OK_RANGE
            i = NormIndex(idx, RangeLen(o), "range object")
            out = CLng(o.I1 + i * o.I3)
            Exit Sub
        Case OK_INST
            If CallDunder1(ov, "__getitem__", idx, out) Then Exit Sub
        Case OK_CLASS
            SetV out, ov
            Exit Sub
        Case OK_BUILTIN
            If o.Flags = 1 Then
                SetV out, ov
                Exit Sub
            End If
        End Select
    End If
    ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object is not subscriptable"
End Sub

Private Sub DictMissing(ByVal o As PyObj, ByRef key As Variant, ByRef out As Variant)
    Dim a() As Variant, v As Variant
    If o.Flags = 2 Then
        out = 0&
        Exit Sub
    End If
    If o.Flags = 1 Then
        If Not IsEmpty(o.FnV) Then
            ReDim a(0 To 0)
            SetV v, o.FnV
            CallValue v, a, 0, Nothing, out
            DictSet o, key, out
            Exit Sub
        End If
    End If
    RaiseExc MakeExcV("KeyError", key)
End Sub

Public Sub SetItem(ByRef ov As Variant, ByRef idx As Variant, ByRef v As Variant)
    Dim o As PyObj, i As Long, r As Variant
    If IsObject(ov) Then
        Set o = ov
        Select Case o.Kind
        Case OK_LIST
            i = NormIndex(idx, o.Count, "list assignment")
            o.SetI i, v
            Exit Sub
        Case OK_DICT
            DictSet o, idx, v
            Exit Sub
        Case OK_INST
            If CallDunder2(ov, "__setitem__", idx, v, r) Then Exit Sub
        End Select
    End If
    ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object does not support item assignment"
End Sub

Public Sub DelItem(ByRef ov As Variant, ByRef idx As Variant)
    Dim o As PyObj, i As Long, r As Variant
    If IsObject(ov) Then
        Set o = ov
        Select Case o.Kind
        Case OK_LIST
            i = NormIndex(idx, o.Count, "list assignment")
            o.RemoveAt i
            Exit Sub
        Case OK_DICT
            If Not DictDel(o, idx) Then RaiseExc MakeExcV("KeyError", idx)
            Exit Sub
        Case OK_INST
            If CallDunder1(ov, "__delitem__", idx, r) Then Exit Sub
        End Select
    End If
    ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object does not support item deletion"
End Sub

Private Function SliceInt(ByRef v As Variant) As Long
    Dim x As Variant
    If Not IsIntType(v) Then ThrowErr "TypeError", "slice indices must be integers or None"
    x = IntOf(v)
    If VarType(x) = vbDecimal Then
        If x > 0 Then
            SliceInt = 2147483647
        Else
            SliceInt = -2147483647
        End If
    Else
        SliceInt = x
    End If
End Function

Private Function ClampIdx(ByVal i As Long, ByVal n As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If i < 0 Then i = i + n
    If i < lo Then i = lo
    If i > hi Then i = hi
    ClampIdx = i
End Function

Public Sub SliceIndices(ByVal n As Long, ByRef vs As Variant, ByRef ve As Variant, ByRef vst As Variant, _
                        ByRef start As Long, ByRef stp As Long, ByRef cnt As Long)
    Dim st As Long, en As Long
    If IsEmpty(vst) Then
        stp = 1
    Else
        stp = SliceInt(vst)
        If stp = 0 Then ThrowErr "ValueError", "slice step cannot be zero"
    End If
    If stp > 0 Then
        If IsEmpty(vs) Then
            st = 0
        Else
            st = ClampIdx(SliceInt(vs), n, 0, n)
        End If
        If IsEmpty(ve) Then
            en = n
        Else
            en = ClampIdx(SliceInt(ve), n, 0, n)
        End If
        If en > st Then
            cnt = (en - st + stp - 1) \ stp
        Else
            cnt = 0
        End If
    Else
        If IsEmpty(vs) Then
            st = n - 1
        Else
            st = ClampIdx(SliceInt(vs), n, -1, n - 1)
        End If
        If IsEmpty(ve) Then
            en = -1
        Else
            en = ClampIdx(SliceInt(ve), n, -1, n - 1)
        End If
        If st > en Then
            cnt = (st - en - stp - 1) \ (-stp)
        Else
            cnt = 0
        End If
    End If
    start = st
End Sub

Public Sub DoSlice(ByRef ov As Variant, ByRef vs As Variant, ByRef ve As Variant, ByRef vst As Variant, ByRef out As Variant)
    Dim o As PyObj, r As PyObj, st As Long, stp As Long, cnt As Long, i As Long, x As Variant, parts() As String
    If VarType(ov) = vbString Then
        SliceIndices Len(ov), vs, ve, vst, st, stp, cnt
        If cnt <= 0 Then
            out = vbNullString
        ElseIf stp = 1 Then
            out = Mid$(ov, st + 1, cnt)
        Else
            ReDim parts(0 To cnt - 1)
            For i = 0 To cnt - 1
                parts(i) = Mid$(ov, st + 1 + i * stp, 1)
            Next
            out = Join(parts, vbNullString)
        End If
        Exit Sub
    End If
    If IsObject(ov) Then
        Set o = ov
        If o.Kind = OK_LIST Or o.Kind = OK_TUPLE Then
            SliceIndices o.Count, vs, ve, vst, st, stp, cnt
            Set r = NewObj(o.Kind)
            If cnt > 0 Then r.Reserve cnt
            For i = 0 To cnt - 1
                o.GetInto st + i * stp, x
                r.Push x
            Next
            Set out = r
            Exit Sub
        ElseIf o.Kind = OK_RANGE Then
            SliceIndices RangeLen(o), vs, ve, vst, st, stp, cnt
            Set out = NewRange(o.I1 + st * o.I3, o.I1 + (st + cnt * stp) * o.I3, o.I3 * stp)
            Exit Sub
        End If
    End If
    ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object is not subscriptable (slice)"
End Sub

Public Sub SetSlice(ByRef ov As Variant, ByRef vs As Variant, ByRef ve As Variant, ByRef vst As Variant, ByRef v As Variant)
    Dim o As PyObj, items As PyObj, st As Long, stp As Long, cnt As Long, i As Long, x As Variant
    Dim arr() As Variant, n As Long, k As Long, total As Long
    If Not IsKind(ov, OK_LIST) Then ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object does not support slice assignment"
    Set o = ov
    Set items = NewListFrom(v)
    n = o.Count
    SliceIndices n, vs, ve, vst, st, stp, cnt
    If stp = 1 Then
        If cnt < 0 Then cnt = 0
        total = n - cnt + items.Count
        ReDim arr(0 To IIf(total > 0, total - 1, 0))
        k = 0
        For i = 0 To st - 1
            o.GetInto i, x
            SetV arr(k), x
            k = k + 1
        Next
        For i = 0 To items.Count - 1
            items.GetInto i, x
            SetV arr(k), x
            k = k + 1
        Next
        For i = st + cnt To n - 1
            o.GetInto i, x
            SetV arr(k), x
            k = k + 1
        Next
        o.SetArray arr, total
    Else
        If items.Count <> cnt Then
            ThrowErr "ValueError", "attempt to assign sequence of size " & items.Count & " to extended slice of size " & cnt
        End If
        For i = 0 To cnt - 1
            items.GetInto i, x
            o.SetI st + i * stp, x
        Next
    End If
End Sub

Public Sub DelSlice(ByRef ov As Variant, ByRef vs As Variant, ByRef ve As Variant, ByRef vst As Variant)
    Dim o As PyObj, st As Long, stp As Long, cnt As Long, i As Long, x As Variant
    Dim drop() As Boolean, arr() As Variant, k As Long, n As Long
    If Not IsKind(ov, OK_LIST) Then ThrowErr "TypeError", "'" & TypeNameOf(ov) & "' object does not support item deletion"
    Set o = ov
    n = o.Count
    If n = 0 Then Exit Sub
    SliceIndices n, vs, ve, vst, st, stp, cnt
    If cnt <= 0 Then Exit Sub
    ReDim drop(0 To n - 1)
    For i = 0 To cnt - 1
        drop(st + i * stp) = True
    Next
    ReDim arr(0 To n - 1)
    For i = 0 To n - 1
        If Not drop(i) Then
            o.GetInto i, x
            SetV arr(k), x
            k = k + 1
        End If
    Next
    o.SetArray arr, k
End Sub

' ===========================================================================
'  attributes
' ===========================================================================
Public Function ClassLookup(ByVal cls As PyObj, ByVal nm As String, ByVal h As Long, ByRef out As Variant) As Boolean
    Dim i As Long, bv As Variant
    i = cls.FindH(nm, h)
    If i >= 0 Then
        cls.ValInto i, out
        ClassLookup = True
        Exit Function
    End If
    For i = 0 To cls.Count - 1
        cls.GetInto i, bv
        If ClassLookup(bv, nm, h, out) Then
            ClassLookup = True
            Exit Function
        End If
    Next
End Function

' turn a class attribute into what attribute access on selfV returns
Public Sub BindAttr(ByRef f As Variant, ByRef selfV As Variant, ByVal cls As PyObj, ByRef out As Variant)
    Dim fo As PyObj, b As PyObj
    If IsObject(f) Then
        Set fo = f
        Select Case fo.Kind
        Case OK_FUNC, OK_BUILTIN
            If fo.Kind = OK_BUILTIN And fo.Flags <> 2 Then
                Set out = fo
                Exit Sub
            End If
            If fo.Kind = OK_FUNC And fo.Flags = 1 Then
                Set out = fo
                Exit Sub
            End If
            Set b = NewObj(OK_BOUND)
            Set b.FnV = fo
            If fo.Kind = OK_FUNC And fo.Flags = 2 Then
                Set b.SelfV = cls
            ElseIf IsObject(selfV) Then
                Set b.SelfV = selfV
            Else
                b.SelfV = selfV
            End If
            Set out = b
            Exit Sub
        Case OK_PROP
            CallValue1 fo.FnV, selfV, out
            Exit Sub
        End Select
    End If
    SetV out, f
End Sub

Public Function TryGetAttr(ByRef ov As Variant, ByVal nm As String, ByRef out As Variant) As Boolean
    TryGetAttr = TryGetAttrH(ov, nm, HashStr(nm), out)
End Function

Public Function TryGetAttrH(ByRef ov As Variant, ByVal nm As String, ByVal h As Long, ByRef out As Variant) As Boolean
    Dim o As PyObj, i As Long, f As Variant, b As PyObj, c As PyObj, sv As Variant, cv As Variant
    TryGetAttrH = True
    If IsObject(ov) Then
        Set o = ov
        Select Case o.Kind
        Case OK_INST
            i = o.FindH(nm, h)
            If i >= 0 Then
                o.ValInto i, out
                Exit Function
            End If
            If ClassLookup(o.Cls, nm, h, f) Then
                BindAttr f, ov, o.Cls, out
                Exit Function
            End If
            If nm = "__class__" Then
                Set out = o.Cls
                Exit Function
            End If
            If nm = "__dict__" Then
                Set out = AttrDict(o)
                Exit Function
            End If
            If ClassLookup(o.Cls, "__getattr__", HashStr("__getattr__"), f) Then
                CallValue2 f, ov, nm, out
                Exit Function
            End If
        Case OK_CLASS
            If ClassLookup(o, nm, h, f) Then
                If IsKind(f, OK_FUNC) Then
                    Set b = f
                    If b.Flags = 2 Then
                        Set cv = o
                        BindAttr f, cv, o, out
                        Exit Function
                    End If
                End If
                SetV out, f
                Exit Function
            End If
            If nm = "__name__" Then
                out = o.PName
                Exit Function
            End If
            If nm = "__bases__" Then
                Set b = NewObj(OK_TUPLE)
                For i = 0 To o.Count - 1
                    o.GetInto i, f
                    b.Push f
                Next
                Set out = b
                Exit Function
            End If
            If nm = "__dict__" Then
                Set out = AttrDict(o)
                Exit Function
            End If
        Case OK_MODULE
            i = o.FindH(nm, h)
            If i >= 0 Then
                o.ValInto i, out
                Exit Function
            End If
            If nm = "__name__" Then
                out = o.PName
                Exit Function
            End If
        Case OK_FUNC
            If nm = "__name__" Then
                out = o.PName
                Exit Function
            End If
            i = o.FindH(nm, h)
            If i >= 0 Then
                o.ValInto i, out
                Exit Function
            End If
        Case OK_BOUND
            If nm = "__self__" Then
                SetV out, o.SelfV
                Exit Function
            End If
            If nm = "__func__" Then
                SetV out, o.FnV
                Exit Function
            End If
            If nm = "__name__" Then
                SetV f, o.FnV
                Set b = f
                out = b.PName
                Exit Function
            End If
        Case OK_SUPER
            Set c = o.Cls
            SetV sv, o.SelfV
            For i = 0 To c.Count - 1
                c.GetInto i, cv
                If ClassLookup(cv, nm, h, f) Then
                    BindAttr f, sv, c, out
                    Exit Function
                End If
            Next
            If nm = "__init__" Then
                Set out = BuiltinObj("object.__init__")
                Exit Function
            End If
        Case OK_BUILTIN
            If nm = "__name__" Then
                out = o.PName
                Exit Function
            End If
            If o.Flags = 1 Then
                If TypeHasMethod(o.Code, nm) Then
                    Set b = NewObj(OK_BMETHOD)
                    b.PName = nm
                    b.Flags = 1
                    Set out = b
                    Exit Function
                End If
            End If
        Case OK_PROP
            If nm = "setter" Or nm = "getter" Or nm = "deleter" Then
                Set out = MakeBMethod(ov, nm)
                Exit Function
            End If
        End Select
    End If
    If HasBuiltinMethod(ov, nm) Then
        Set out = MakeBMethod(ov, nm)
        Exit Function
    End If
    TryGetAttrH = False
End Function

Public Function AttrDict(ByVal o As PyObj) As PyObj
    Dim d As PyObj, i As Long, k As Variant, v As Variant
    Set d = NewDict()
    For i = 0 To o.EntryCount - 1
        If o.LiveAt(i) Then
            o.KeyInto i, k
            o.ValInto i, v
            DictSet d, k, v
        End If
    Next
    Set AttrDict = d
End Function

Public Function MakeBMethod(ByRef selfV As Variant, ByVal nm As String) As PyObj
    Dim b As PyObj
    Set b = NewObj(OK_BMETHOD)
    b.PName = nm
    If IsObject(selfV) Then
        Set b.SelfV = selfV
    Else
        b.SelfV = selfV
    End If
    Set MakeBMethod = b
End Function

Public Function BuiltinObj(ByVal nm As String) As PyObj
    Dim b As PyObj
    Set b = NewObj(OK_BUILTIN)
    b.PName = nm
    b.Code = BuiltinCode(nm)
    Set BuiltinObj = b
End Function

Public Sub PyGetAttr(ByRef ov As Variant, ByVal nm As String, ByRef out As Variant)
    PyGetAttrH ov, nm, HashStr(nm), out
End Sub

Public Sub PyGetAttrH(ByRef ov As Variant, ByVal nm As String, ByVal h As Long, ByRef out As Variant)
    Dim o As PyObj
    If TryGetAttrH(ov, nm, h, out) Then Exit Sub
    If IsObject(ov) Then
        Set o = ov
        If o.Kind = OK_MODULE Then ThrowErr "AttributeError", "module '" & o.PName & "' has no attribute '" & nm & "'"
        If o.Kind = OK_CLASS Then ThrowErr "AttributeError", "type object '" & o.PName & "' has no attribute '" & nm & "'"
    End If
    ThrowErr "AttributeError", "'" & TypeNameOf(ov) & "' object has no attribute '" & nm & "'"
End Sub

Public Sub PySetAttr(ByRef ov As Variant, ByVal nm As String, ByRef v As Variant)
    PySetAttrH ov, nm, HashStr(nm), v
End Sub

Public Sub PySetAttrH(ByRef ov As Variant, ByVal nm As String, ByVal h As Long, ByRef v As Variant)
    Dim o As PyObj, f As Variant, p As PyObj, sv As Variant, r As Variant
    If IsObject(ov) Then
        Set o = ov
        Select Case o.Kind
        Case OK_INST
            If (o.Cls.Flags And 4) <> 0 Then
                If ClassLookup(o.Cls, nm, h, f) Then
                    If IsKind(f, OK_PROP) Then
                        Set p = f
                        If IsEmpty(p.SelfV) Then ThrowErr "AttributeError", "can't set attribute '" & nm & "'"
                        SetV sv, p.SelfV
                        CallValue2 sv, ov, v, r
                        Exit Sub
                    End If
                End If
            End If
            o.SetH nm, h, nm, v
            Exit Sub
        Case OK_CLASS, OK_MODULE, OK_FUNC
            o.SetH nm, h, nm, v
            Exit Sub
        End Select
    End If
    ThrowErr "AttributeError", "'" & TypeNameOf(ov) & "' object has no attribute '" & nm & "'"
End Sub

Public Sub PyDelAttr(ByRef ov As Variant, ByVal nm As String)
    Dim o As PyObj
    If IsObject(ov) Then
        Set o = ov
        If o.Kind = OK_INST Or o.Kind = OK_CLASS Or o.Kind = OK_MODULE Then
            If o.DelH(nm, HashStr(nm)) Then Exit Sub
        End If
    End If
    ThrowErr "AttributeError", "'" & TypeNameOf(ov) & "' object has no attribute '" & nm & "'"
End Sub

' ===========================================================================
'  operators
' ===========================================================================
Private Function RepeatStr(ByVal s As String, ByVal n As Variant) As String
    If VarType(n) = vbDecimal Then ThrowErr "MemoryError", "repeated string is too long"
    If n <= 0 Or Len(s) = 0 Then Exit Function
    If CDbl(Len(s)) * CDbl(n) > 100000000# Then ThrowErr "MemoryError", "repeated string is too long"
    RepeatStr = Replace(Space$(n), " ", s)
End Function

Private Function RepeatSeq(ByVal o As PyObj, ByVal n As Variant) As PyObj
    Dim r As PyObj, i As Long, j As Long, x As Variant
    Set r = NewObj(o.Kind)
    If VarType(n) = vbDecimal Then ThrowErr "MemoryError", "repeated sequence is too long"
    If n > 0 Then
        If CDbl(o.Count) * CDbl(n) > 50000000# Then ThrowErr "MemoryError", "repeated sequence is too long"
        r.Reserve o.Count * n
        For j = 1 To n
            For i = 0 To o.Count - 1
                o.GetInto i, x
                r.Push x
            Next
        Next
    End If
    Set RepeatSeq = r
End Function

Private Sub NumBinOp(ByVal op As Long, ByVal a As Variant, ByVal b As Variant, ByRef out As Variant)
    Dim fa As Boolean, q As Variant, r As Variant, da As Double, db As Double
    a = IntOf(a)
    b = IntOf(b)
    fa = IsFloatType(a) Or IsFloatType(b)
    Select Case op
    Case OP_ADD
        If fa Then
            out = CDbl(a) + CDbl(b)
        Else
            out = IntAdd(a, b)
        End If
    Case OP_SUB
        If fa Then
            out = CDbl(a) - CDbl(b)
        Else
            out = IntSub(a, b)
        End If
    Case OP_MUL
        If fa Then
            out = CDbl(a) * CDbl(b)
        Else
            out = IntMul(a, b)
        End If
    Case OP_DIV
        db = CDbl(b)
        If db = 0 Then ThrowErr "ZeroDivisionError", "division by zero"
        If VarType(a) = vbDecimal Or VarType(b) = vbDecimal Then
            out = CDbl(CDec(a) / CDec(b))
        Else
            out = CDbl(a) / db
        End If
    Case OP_FLOORDIV
        If fa Then
            db = CDbl(b)
            If db = 0 Then ThrowErr "ZeroDivisionError", "float floor division by zero"
            out = Int(CDbl(a) / db)
        Else
            IntDivMod a, b, q, r
            out = q
        End If
    Case OP_MOD
        If fa Then
            da = CDbl(a)
            db = CDbl(b)
            If db = 0 Then ThrowErr "ZeroDivisionError", "float modulo"
            out = da - db * Int(da / db)
        Else
            IntDivMod a, b, q, r
            out = r
        End If
    Case OP_POW
        If Not fa Then
            If b >= 0 Then
                out = IntPow(a, b)
                Exit Sub
            End If
        End If
        da = CDbl(a)
        db = CDbl(b)
        If da = 0 And db < 0 Then ThrowErr "ZeroDivisionError", "0.0 cannot be raised to a negative power"
        If da < 0 And db <> Int(db) Then ThrowErr "ValueError", "negative number cannot be raised to a fractional power"
        out = da ^ db
    Case OP_BITAND, OP_BITOR, OP_BITXOR, OP_LSHIFT, OP_RSHIFT
        If fa Then ThrowErr "TypeError", "unsupported operand type(s) for " & OpSym(op) & ": '" & TypeNameOf(a) & "' and '" & TypeNameOf(b) & "'"
        Select Case op
        Case OP_LSHIFT
            If b < 0 Then ThrowErr "ValueError", "negative shift count"
            out = IntMul(a, IntPow(2&, b))
        Case OP_RSHIFT
            If b < 0 Then ThrowErr "ValueError", "negative shift count"
            If b > 95 Then
                If a < 0 Then
                    out = -1&
                Else
                    out = 0&
                End If
            Else
                IntDivMod a, IntPow(2&, b), q, r
                out = q
            End If
        Case Else
            If VarType(a) = vbDecimal Or VarType(b) = vbDecimal Then ThrowErr "OverflowError", "bitwise operations are limited to 32-bit integers"
            If op = OP_BITAND Then
                out = CLng(CLng(a) And CLng(b))
            ElseIf op = OP_BITOR Then
                out = CLng(CLng(a) Or CLng(b))
            Else
                out = CLng(CLng(a) Xor CLng(b))
            End If
        End Select
    Case Else
        ThrowErr "TypeError", "unsupported operand type(s) for " & OpSym(op) & ": '" & TypeNameOf(a) & "' and '" & TypeNameOf(b) & "'"
    End Select
End Sub

Private Function DunderName(ByVal op As Long) As String
    Select Case op
    Case OP_ADD: DunderName = "add"
    Case OP_SUB: DunderName = "sub"
    Case OP_MUL: DunderName = "mul"
    Case OP_DIV: DunderName = "truediv"
    Case OP_FLOORDIV: DunderName = "floordiv"
    Case OP_MOD: DunderName = "mod"
    Case OP_POW: DunderName = "pow"
    Case OP_LSHIFT: DunderName = "lshift"
    Case OP_RSHIFT: DunderName = "rshift"
    Case OP_BITAND: DunderName = "and"
    Case OP_BITOR: DunderName = "or"
    Case OP_BITXOR: DunderName = "xor"
    Case OP_MATMUL: DunderName = "matmul"
    End Select
End Function

Private Function TryBinDunder(ByVal op As Long, ByRef a As Variant, ByRef b As Variant, ByRef out As Variant) As Boolean
    Dim nm As String
    nm = DunderName(op)
    If IsKind(a, OK_INST) Then
        If CallDunder1(a, "__" & nm & "__", b, out) Then
            If Not IsNotImpl(out) Then
                TryBinDunder = True
                Exit Function
            End If
        End If
    End If
    If IsKind(b, OK_INST) Then
        If CallDunder1(b, "__r" & nm & "__", a, out) Then
            If Not IsNotImpl(out) Then
                TryBinDunder = True
                Exit Function
            End If
        End If
    End If
End Function

Private Function SetCombine(ByVal op As Long, ByVal a As PyObj, ByVal b As PyObj) As PyObj
    Dim r As PyObj, i As Long, k As Variant, e As Variant
    Set r = NewObj(OK_SET)
    Select Case op
    Case OP_BITOR
        Set r = CopyDict(a, OK_SET)
        For i = 0 To b.EntryCount - 1
            If b.LiveAt(i) Then
                b.KeyInto i, k
                If r.FindH(b.KSAt(i), b.HAt(i)) < 0 Then r.SetH b.KSAt(i), b.HAt(i), k, e
            End If
        Next
    Case OP_BITAND
        For i = 0 To a.EntryCount - 1
            If a.LiveAt(i) Then
                If b.FindH(a.KSAt(i), a.HAt(i)) >= 0 Then
                    a.KeyInto i, k
                    r.SetH a.KSAt(i), a.HAt(i), k, e
                End If
            End If
        Next
    Case OP_SUB
        For i = 0 To a.EntryCount - 1
            If a.LiveAt(i) Then
                If b.FindH(a.KSAt(i), a.HAt(i)) < 0 Then
                    a.KeyInto i, k
                    r.SetH a.KSAt(i), a.HAt(i), k, e
                End If
            End If
        Next
    Case OP_BITXOR
        For i = 0 To a.EntryCount - 1
            If a.LiveAt(i) Then
                If b.FindH(a.KSAt(i), a.HAt(i)) < 0 Then
                    a.KeyInto i, k
                    r.SetH a.KSAt(i), a.HAt(i), k, e
                End If
            End If
        Next
        For i = 0 To b.EntryCount - 1
            If b.LiveAt(i) Then
                If a.FindH(b.KSAt(i), b.HAt(i)) < 0 Then
                    b.KeyInto i, k
                    r.SetH b.KSAt(i), b.HAt(i), k, e
                End If
            End If
        Next
    End Select
    Set SetCombine = r
End Function

Public Sub BinOp(ByVal op As Long, ByRef a As Variant, ByRef b As Variant, ByRef out As Variant)
    Dim ta As Long, tb As Long, d As Double, oa As PyObj, ob As PyObj, r As PyObj, i As Long, x As Variant
    Dim la As Long, lb As Long, lr As Long, lq As Long
    ta = VarType(a)
    tb = VarType(b)
    If ta = vbLong And tb = vbLong Then
        Select Case op
        Case OP_ADD
            d = CDbl(a) + CDbl(b)
            If d >= -2147483648# And d <= 2147483647# Then
                out = CLng(d)
                Exit Sub
            End If
        Case OP_SUB
            d = CDbl(a) - CDbl(b)
            If d >= -2147483648# And d <= 2147483647# Then
                out = CLng(d)
                Exit Sub
            End If
        Case OP_MUL
            d = CDbl(a) * CDbl(b)
            If d >= -2147483648# And d <= 2147483647# Then
                out = CLng(d)
                Exit Sub
            End If
        Case OP_MOD, OP_FLOORDIV
            la = a
            lb = b
            If lb <> 0 And lb <> -1 Then
                lr = la Mod lb
                If op = OP_MOD Then
                    If lr <> 0 Then
                        If (lr < 0) <> (lb < 0) Then lr = lr + lb
                    End If
                    out = lr
                Else
                    lq = la \ lb
                    If lr <> 0 Then
                        If (la < 0) <> (lb < 0) Then lq = lq - 1
                    End If
                    out = lq
                End If
                Exit Sub
            End If
        Case OP_DIV
            If b <> 0 Then
                out = CDbl(a) / CDbl(b)
                Exit Sub
            End If
        End Select
    End If
    If IsNum(a) Then
        If IsNum(b) Then
            NumBinOp op, a, b, out
            Exit Sub
        End If
    End If
    Select Case op
    Case OP_ADD
        If ta = vbString And tb = vbString Then
            out = a & b
            Exit Sub
        End If
        If ta = vbObject And tb = vbObject Then
            Set oa = a
            Set ob = b
            If (oa.Kind = OK_LIST Or oa.Kind = OK_TUPLE) And oa.Kind = ob.Kind Then
                Set r = NewObj(oa.Kind)
                r.Reserve oa.Count + ob.Count
                For i = 0 To oa.Count - 1
                    oa.GetInto i, x
                    r.Push x
                Next
                For i = 0 To ob.Count - 1
                    ob.GetInto i, x
                    r.Push x
                Next
                Set out = r
                Exit Sub
            End If
        End If
    Case OP_MUL
        If ta = vbString And IsIntType(b) Then
            out = RepeatStr(a, IntOf(b))
            Exit Sub
        End If
        If tb = vbString And IsIntType(a) Then
            out = RepeatStr(b, IntOf(a))
            Exit Sub
        End If
        If IsIntType(b) Then
            If IsKind(a, OK_LIST) Or IsKind(a, OK_TUPLE) Then
                Set out = RepeatSeq(a, IntOf(b))
                Exit Sub
            End If
        End If
        If IsIntType(a) Then
            If IsKind(b, OK_LIST) Or IsKind(b, OK_TUPLE) Then
                Set out = RepeatSeq(b, IntOf(a))
                Exit Sub
            End If
        End If
    Case OP_MOD
        If ta = vbString Then
            out = PercentFormat(CStr(a), b)
            Exit Sub
        End If
    Case OP_SUB, OP_BITAND, OP_BITOR, OP_BITXOR
        If IsKind(a, OK_SET) And IsKind(b, OK_SET) Then
            Set out = SetCombine(op, a, b)
            Exit Sub
        End If
        If op = OP_BITOR Then
            If IsKind(a, OK_DICT) And IsKind(b, OK_DICT) Then
                Set oa = a
                Set ob = b
                Set r = CopyDict(oa, OK_DICT)
                DictUpdate r, ob
                Set out = r
                Exit Sub
            End If
        End If
    End Select
    If TryBinDunder(op, a, b, out) Then Exit Sub
    ThrowErr "TypeError", "unsupported operand type(s) for " & OpSym(op) & ": '" & TypeNameOf(a) & "' and '" & TypeNameOf(b) & "'"
End Sub

Public Sub DictUpdate(ByVal d As PyObj, ByVal src As PyObj)
    Dim i As Long, k As Variant, v As Variant
    For i = 0 To src.EntryCount - 1
        If src.LiveAt(i) Then
            src.KeyInto i, k
            src.ValInto i, v
            d.SetH src.KSAt(i), src.HAt(i), k, v
        End If
    Next
End Sub

Public Sub UnaryOp(ByVal op As Long, ByRef a As Variant, ByRef out As Variant)
    Dim nm As String
    If IsNum(a) Then
        Select Case op
        Case OP_NEG
            If IsFloatType(a) Then
                out = -CDbl(a)
            Else
                out = IntNeg(a)
            End If
            Exit Sub
        Case OP_POS
            If IsFloatType(a) Then
                out = CDbl(a)
            Else
                out = IntOf(a)
            End If
            Exit Sub
        Case OP_INVERT
            If IsIntType(a) Then
                out = IntSub(IntNeg(a), 1&)
                Exit Sub
            End If
        End Select
    End If
    Select Case op
    Case OP_NEG: nm = "__neg__"
    Case OP_POS: nm = "__pos__"
    Case OP_INVERT: nm = "__invert__"
    End Select
    If CallDunder0(a, nm, out) Then Exit Sub
    Select Case op
    Case OP_NEG: nm = "-"
    Case OP_POS: nm = "+"
    Case OP_INVERT: nm = "~"
    End Select
    ThrowErr "TypeError", "bad operand type for unary " & nm & ": '" & TypeNameOf(a) & "'"
End Sub

Public Function PyLen(ByRef v As Variant) As Long
    Dim o As PyObj, r As Variant
    If VarType(v) = vbString Then
        PyLen = Len(v)
        Exit Function
    End If
    If IsObject(v) Then
        Set o = v
        Select Case o.Kind
        Case OK_LIST, OK_TUPLE: PyLen = o.Count: Exit Function
        Case OK_DICT, OK_SET: PyLen = o.DCount: Exit Function
        Case OK_RANGE: PyLen = RangeLen(o): Exit Function
        Case OK_INST
            If CallDunder0(v, "__len__", r) Then
                PyLen = IntToLong(r, "__len__() result")
                Exit Function
            End If
        End Select
    End If
    ThrowErr "TypeError", "object of type '" & TypeNameOf(v) & "' has no len()"
End Function
