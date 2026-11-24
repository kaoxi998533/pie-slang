Attribute VB_Name = "PyBuiltins"
Option Explicit
' ===========================================================================
'  PyBuiltins - built-in functions, methods of str / list / dict / set /
'  tuple, string formatting, sorting and the small standard modules
'  (math, random, time, string, sys, collections, functools, itertools).
' ===========================================================================

Public Const PY_INF As Double = 1.797693134862E+308

' ---- built-in function codes ----
Public Const BI_PRINT As Long = 1
Public Const BI_LEN As Long = 2
Public Const BI_RANGE As Long = 3
Public Const BI_INT As Long = 4
Public Const BI_FLOAT As Long = 5
Public Const BI_STR As Long = 6
Public Const BI_BOOL As Long = 7
Public Const BI_LIST As Long = 8
Public Const BI_DICT As Long = 9
Public Const BI_TUPLE As Long = 10
Public Const BI_SET As Long = 11
Public Const BI_ABS As Long = 12
Public Const BI_MIN As Long = 13
Public Const BI_MAX As Long = 14
Public Const BI_SUM As Long = 15
Public Const BI_SORTED As Long = 16
Public Const BI_REVERSED As Long = 17
Public Const BI_ENUMERATE As Long = 18
Public Const BI_ZIP As Long = 19
Public Const BI_MAP As Long = 20
Public Const BI_FILTER As Long = 21
Public Const BI_ANY As Long = 22
Public Const BI_ALL As Long = 23
Public Const BI_ROUND As Long = 24
Public Const BI_POW As Long = 25
Public Const BI_DIVMOD As Long = 26
Public Const BI_ISINSTANCE As Long = 27
Public Const BI_ISSUBCLASS As Long = 28
Public Const BI_TYPE As Long = 29
Public Const BI_REPR As Long = 30
Public Const BI_CHR As Long = 31
Public Const BI_ORD As Long = 32
Public Const BI_HEX As Long = 33
Public Const BI_BIN As Long = 34
Public Const BI_OCT As Long = 35
Public Const BI_INPUT As Long = 36
Public Const BI_ID As Long = 37
Public Const BI_HASH As Long = 38
Public Const BI_CALLABLE As Long = 39
Public Const BI_GETATTR As Long = 40
Public Const BI_SETATTR As Long = 41
Public Const BI_HASATTR As Long = 42
Public Const BI_DELATTR As Long = 43
Public Const BI_FORMAT As Long = 44
Public Const BI_ITER As Long = 45
Public Const BI_NEXT As Long = 46
Public Const BI_STATICMETHOD As Long = 48
Public Const BI_CLASSMETHOD As Long = 49
Public Const BI_PROPERTY As Long = 50
Public Const BI_EXIT As Long = 51
Public Const BI_SUPER As Long = 52
Public Const BI_GLOBALS As Long = 53
Public Const BI_DIR As Long = 55
Public Const BI_VARS As Long = 56
Public Const BI_FROZENSET As Long = 57
Public Const BI_RANGETYPE As Long = 58
Public Const BI_TYPETYPE As Long = 59
Public Const BI_EXC_INIT As Long = 70
Public Const BI_EXC_STR As Long = 71
Public Const BI_EXC_REPR As Long = 72
Public Const BI_OBJ_INIT As Long = 73
' math
Public Const BI_M_SQRT As Long = 100
Public Const BI_M_FLOOR As Long = 101
Public Const BI_M_CEIL As Long = 102
Public Const BI_M_FABS As Long = 103
Public Const BI_M_FACTORIAL As Long = 104
Public Const BI_M_GCD As Long = 105
Public Const BI_M_POW As Long = 106
Public Const BI_M_EXP As Long = 107
Public Const BI_M_LOG As Long = 108
Public Const BI_M_LOG10 As Long = 109
Public Const BI_M_LOG2 As Long = 110
Public Const BI_M_SIN As Long = 111
Public Const BI_M_COS As Long = 112
Public Const BI_M_TAN As Long = 113
Public Const BI_M_ASIN As Long = 114
Public Const BI_M_ACOS As Long = 115
Public Const BI_M_ATAN As Long = 116
Public Const BI_M_ATAN2 As Long = 117
Public Const BI_M_DEGREES As Long = 118
Public Const BI_M_RADIANS As Long = 119
Public Const BI_M_HYPOT As Long = 120
Public Const BI_M_ISQRT As Long = 121
Public Const BI_M_TRUNC As Long = 122
Public Const BI_M_ISCLOSE As Long = 123
Public Const BI_M_COMB As Long = 124
Public Const BI_M_PERM As Long = 125
Public Const BI_M_PROD As Long = 126
Public Const BI_M_FSUM As Long = 127
Public Const BI_M_LCM As Long = 128
Public Const BI_M_COPYSIGN As Long = 129
Public Const BI_M_FMOD As Long = 130
Public Const BI_M_ISFINITE As Long = 131
Public Const BI_M_ISINF As Long = 132
Public Const BI_M_ISNAN As Long = 133
' random
Public Const BI_R_RANDOM As Long = 150
Public Const BI_R_RANDINT As Long = 151
Public Const BI_R_CHOICE As Long = 152
Public Const BI_R_SHUFFLE As Long = 153
Public Const BI_R_UNIFORM As Long = 154
Public Const BI_R_SEED As Long = 155
Public Const BI_R_RANDRANGE As Long = 156
Public Const BI_R_SAMPLE As Long = 157
' time / sys
Public Const BI_T_TIME As Long = 170
Public Const BI_T_SLEEP As Long = 171
Public Const BI_T_PERF As Long = 172
Public Const BI_S_SETRECLIMIT As Long = 180
Public Const BI_S_GETRECLIMIT As Long = 181
' collections / functools / itertools
Public Const BI_C_DEFAULTDICT As Long = 190
Public Const BI_C_COUNTER As Long = 191
Public Const BI_C_ORDEREDDICT As Long = 192
Public Const BI_C_DEQUE As Long = 193
Public Const BI_F_REDUCE As Long = 200
Public Const BI_F_LRU As Long = 201
Public Const BI_F_CACHED As Long = 202
Public Const BI_F_LRU_DECO As Long = 203
Public Const BI_I_PERMUTATIONS As Long = 210
Public Const BI_I_COMBINATIONS As Long = 211
Public Const BI_I_PRODUCT As Long = 212
Public Const BI_I_CHAIN As Long = 213
Public Const BI_I_ACCUMULATE As Long = 214
Public Const BI_I_COMBWR As Long = 215
' dataclasses / typing
Public Const BI_DC_DATACLASS As Long = 220
Public Const BI_DC_DECO As Long = 221
Public Const BI_DC_FIELD As Long = 222
Public Const BI_DC_INIT As Long = 223
Public Const BI_DC_REPR As Long = 224
Public Const BI_DC_EQ As Long = 225
Public Const BI_DC_LT As Long = 226
Public Const BI_DC_LE As Long = 227
Public Const BI_DC_GT As Long = 228
Public Const BI_DC_GE As Long = 229
Public Const BI_DC_ASDICT As Long = 230
Public Const BI_TY_TYPEVAR As Long = 235
' host (Excel) functions, implemented in PyExcel
Public Const BI_HOST_FIRST As Long = 300
Public Const BI_X_CELL As Long = 300
Public Const BI_X_SETCELL As Long = 301
Public Const BI_X_GETRANGE As Long = 302
Public Const BI_X_SETRANGE As Long = 303
Public Const BI_X_MSGBOX As Long = 304
Public Const BI_X_EVAL As Long = 305
Public Const BI_X_SHEETNAMES As Long = 306
Public Const BI_X_CLEAR As Long = 307
Public Const BI_HOST_LAST As Long = 399

Private Const STR_METHODS As String = " upper lower strip lstrip rstrip split rsplit join replace find rfind index rindex count startswith endswith isdigit isnumeric isdecimal isalpha isalnum isspace isupper islower istitle title capitalize swapcase center ljust rjust zfill format splitlines partition rpartition casefold removeprefix removesuffix expandtabs encode "
Private Const LIST_METHODS As String = " append extend insert pop remove index count sort reverse clear copy appendleft popleft extendleft rotate "
Private Const DICT_METHODS As String = " keys values items get pop popitem setdefault update clear copy most_common elements total "
Private Const SET_METHODS As String = " add remove discard pop clear copy union intersection difference symmetric_difference issubset issuperset isdisjoint update intersection_update difference_update symmetric_difference_update "

Private mRndInit As Boolean

' ---------------------------------------------------------------------------
'  registration
' ---------------------------------------------------------------------------
Private Function RegB(ByVal target As PyObj, ByVal nm As String, ByVal code As Long, Optional ByVal isType As Boolean = False) As PyObj
    Dim o As PyObj
    Set o = NewObj(OK_BUILTIN)
    o.PName = nm
    o.Code = code
    If isType Then o.Flags = 1
    NameSet target, nm, o
    Set RegB = o
End Function

Private Function MakeExcClass(ByVal nm As String, ByVal bs As PyObj) As PyObj
    Dim c As PyObj
    Set c = NewObj(OK_CLASS)
    c.PName = nm
    c.Flags = 1
    If Not (bs Is Nothing) Then c.Push bs
    NameSet gBuiltins, nm, c
    Set MakeExcClass = c
End Function

Public Sub InitBuiltins()
    Dim b As PyObj, ex As PyObj, ar As PyObj, lk As PyObj, rt As PyObj, se As PyObj, o As PyObj, im As PyObj
    Set gBuiltins = NewObj(OK_ENV)
    Set gNoDefault = NewObj(OK_SENTINEL)
    gNoDefault.PName = "<no default>"
    Set gNotImpl = NewObj(OK_SENTINEL)
    gNotImpl.PName = "NotImplemented"
    NameSet gBuiltins, "NotImplemented", gNotImpl
    Set gObjectClass = NewObj(OK_CLASS)
    gObjectClass.PName = "object"
    gObjectClass.Flags = 1
    NameSet gBuiltins, "object", gObjectClass

    RegB gBuiltins, "print", BI_PRINT
    RegB gBuiltins, "len", BI_LEN
    RegB gBuiltins, "range", BI_RANGE, True
    RegB gBuiltins, "int", BI_INT, True
    RegB gBuiltins, "float", BI_FLOAT, True
    RegB gBuiltins, "str", BI_STR, True
    RegB gBuiltins, "bool", BI_BOOL, True
    RegB gBuiltins, "list", BI_LIST, True
    RegB gBuiltins, "dict", BI_DICT, True
    RegB gBuiltins, "tuple", BI_TUPLE, True
    RegB gBuiltins, "set", BI_SET, True
    RegB gBuiltins, "frozenset", BI_FROZENSET, True
    RegB gBuiltins, "type", BI_TYPE, True
    RegB gBuiltins, "abs", BI_ABS
    RegB gBuiltins, "min", BI_MIN
    RegB gBuiltins, "max", BI_MAX
    RegB gBuiltins, "sum", BI_SUM
    RegB gBuiltins, "sorted", BI_SORTED
    RegB gBuiltins, "reversed", BI_REVERSED
    RegB gBuiltins, "enumerate", BI_ENUMERATE
    RegB gBuiltins, "zip", BI_ZIP
    RegB gBuiltins, "map", BI_MAP
    RegB gBuiltins, "filter", BI_FILTER
    RegB gBuiltins, "any", BI_ANY
    RegB gBuiltins, "all", BI_ALL
    RegB gBuiltins, "round", BI_ROUND
    RegB gBuiltins, "pow", BI_POW
    RegB gBuiltins, "divmod", BI_DIVMOD
    RegB gBuiltins, "isinstance", BI_ISINSTANCE
    RegB gBuiltins, "issubclass", BI_ISSUBCLASS
    RegB gBuiltins, "repr", BI_REPR
    RegB gBuiltins, "ascii", BI_REPR
    RegB gBuiltins, "chr", BI_CHR
    RegB gBuiltins, "ord", BI_ORD
    RegB gBuiltins, "hex", BI_HEX
    RegB gBuiltins, "bin", BI_BIN
    RegB gBuiltins, "oct", BI_OCT
    RegB gBuiltins, "input", BI_INPUT
    RegB gBuiltins, "id", BI_ID
    RegB gBuiltins, "hash", BI_HASH
    RegB gBuiltins, "callable", BI_CALLABLE
    RegB gBuiltins, "getattr", BI_GETATTR
    RegB gBuiltins, "setattr", BI_SETATTR
    RegB gBuiltins, "hasattr", BI_HASATTR
    RegB gBuiltins, "delattr", BI_DELATTR
    RegB gBuiltins, "format", BI_FORMAT
    RegB gBuiltins, "iter", BI_ITER
    RegB gBuiltins, "next", BI_NEXT
    RegB gBuiltins, "staticmethod", BI_STATICMETHOD
    RegB gBuiltins, "classmethod", BI_CLASSMETHOD
    RegB gBuiltins, "property", BI_PROPERTY
    RegB gBuiltins, "exit", BI_EXIT
    RegB gBuiltins, "quit", BI_EXIT
    RegB gBuiltins, "super", BI_SUPER
    RegB gBuiltins, "globals", BI_GLOBALS
    RegB gBuiltins, "dir", BI_DIR
    RegB gBuiltins, "vars", BI_VARS
    ' Excel helpers
    RegB gBuiltins, "cell", BI_X_CELL
    RegB gBuiltins, "set_cell", BI_X_SETCELL
    RegB gBuiltins, "get_range", BI_X_GETRANGE
    RegB gBuiltins, "set_range", BI_X_SETRANGE
    RegB gBuiltins, "msgbox", BI_X_MSGBOX
    RegB gBuiltins, "excel_eval", BI_X_EVAL
    RegB gBuiltins, "sheet_names", BI_X_SHEETNAMES
    RegB gBuiltins, "clear_range", BI_X_CLEAR

    ' exception hierarchy
    Set b = MakeExcClass("BaseException", Nothing)
    Set o = RegB(b, "__init__", BI_EXC_INIT)
    o.Flags = 2
    Set o = RegB(b, "__str__", BI_EXC_STR)
    o.Flags = 2
    Set o = RegB(b, "__repr__", BI_EXC_REPR)
    o.Flags = 2
    MakeExcClass "SystemExit", b
    MakeExcClass "KeyboardInterrupt", b
    Set ex = MakeExcClass("Exception", b)
    Set ar = MakeExcClass("ArithmeticError", ex)
    MakeExcClass "ZeroDivisionError", ar
    MakeExcClass "OverflowError", ar
    Set lk = MakeExcClass("LookupError", ex)
    MakeExcClass "IndexError", lk
    MakeExcClass "KeyError", lk
    MakeExcClass "ValueError", ex
    MakeExcClass "TypeError", ex
    Set o = MakeExcClass("NameError", ex)
    MakeExcClass "UnboundLocalError", o
    MakeExcClass "AttributeError", ex
    Set rt = MakeExcClass("RuntimeError", ex)
    MakeExcClass "RecursionError", rt
    MakeExcClass "NotImplementedError", rt
    MakeExcClass "StopIteration", ex
    MakeExcClass "AssertionError", ex
    Set im = MakeExcClass("ImportError", ex)
    MakeExcClass "ModuleNotFoundError", im
    Set se = MakeExcClass("SyntaxError", ex)
    MakeExcClass "IndentationError", se
    MakeExcClass "MemoryError", ex
    MakeExcClass "SystemError", ex
    Set o = MakeExcClass("OSError", ex)
    MakeExcClass "FileNotFoundError", o
    MakeExcClass "EOFError", ex
End Sub

Public Function BuiltinCode(ByVal nm As String) As Long
    Select Case nm
    Case "object.__init__": BuiltinCode = BI_OBJ_INIT
    End Select
End Function

' ---------------------------------------------------------------------------
'  helpers
' ---------------------------------------------------------------------------
Private Sub NeedArgs(ByVal fname As String, ByVal na As Long, ByVal mn As Long, ByVal mx As Long)
    If na >= mn And na <= mx Then Exit Sub
    If mn = mx Then
        ThrowErr "TypeError", fname & "() takes exactly " & mn & " argument" & IIf(mn = 1, "", "s") & " (" & na & " given)"
    ElseIf na < mn Then
        ThrowErr "TypeError", fname & "() takes at least " & mn & " argument" & IIf(mn = 1, "", "s") & " (" & na & " given)"
    Else
        ThrowErr "TypeError", fname & "() takes at most " & mx & " argument" & IIf(mx = 1, "", "s") & " (" & na & " given)"
    End If
End Sub

Private Function ArgStr(ByRef v As Variant, ByVal what As String) As String
    If VarType(v) <> vbString Then ThrowErr "TypeError", what & " must be str, not " & TypeNameOf(v)
    ArgStr = v
End Function

Private Function ArgOr(ByRef args() As Variant, ByVal na As Long, ByVal i As Long, ByVal kw As PyObj, ByVal kwName As String, ByRef out As Variant) As Boolean
    ' positional argument i, or keyword kwName
    If i < na Then
        SetV out, args(i)
        ArgOr = True
    ElseIf KwGet(kw, kwName, out) Then
        ArgOr = True
    End If
End Function

Private Function NumArg(ByRef v As Variant, ByVal fname As String) As Double
    If Not IsNum(v) Then ThrowErr "TypeError", "must be real number, not " & TypeNameOf(v)
    NumArg = NumToDbl(v)
End Function

Private Function PyListOf(ByRef v As Variant) As PyObj
    Set PyListOf = NewListFrom(v)
End Function

Private Function ToVar(ByVal o As PyObj) As Variant
    Set ToVar = o
End Function

Private Function IsWs(ByVal c As String) As Boolean
    Dim k As Long
    k = UCode(c)
    Select Case k
    Case 9 To 13, 28 To 32, 133, 160, &H1680&, &H2000& To &H200A&, &H2028&, &H2029&, &H202F&, &H205F&, &H3000&
        IsWs = True
    End Select
End Function

Private Function IsAlphaC(ByVal c As String) As Boolean
    Dim k As Long
    k = UCode(c)
    If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Then
        IsAlphaC = True
    ElseIf k >= 128 Then
        If UCase$(c) <> LCase$(c) Then
            IsAlphaC = True
        ElseIf (k >= &H4E00& And k <= &H9FFF&) Or (k >= &H3400& And k <= &H4DBF&) Or (k >= &H3040& And k <= &H30FF&) Or (k >= &HAC00& And k <= &HD7AF&) Or (k >= &HF900& And k <= &HFAFF&) Then
            IsAlphaC = True
        End If
    End If
End Function

Private Function IsDigitCh(ByVal c As String) As Boolean
    IsDigitCh = (c >= "0" And c <= "9")
End Function

Public Function IntToBase(ByVal x As Variant, ByVal bs As Long) As String
    Dim q As Variant, r As Variant, s As String
    x = IntOf(x)
    If x = 0 Then
        IntToBase = "0"
        Exit Function
    End If
    Do While x > 0
        IntDivMod x, bs, q, r
        s = Mid$("0123456789abcdefghijklmnopqrstuvwxyz", r + 1, 1) & s
        x = q
    Loop
    IntToBase = s
End Function

Private Function MakeTuple(ByRef args() As Variant, ByVal fromI As Long, ByVal na As Long) As PyObj
    Dim t As PyObj, i As Long
    Set t = NewObj(OK_TUPLE)
    For i = fromI To na - 1
        t.Push args(i)
    Next
    Set MakeTuple = t
End Function

' ---------------------------------------------------------------------------
'  sorting (stable merge sort)
' ---------------------------------------------------------------------------
Private Function KeyBefore(ByRef a As Variant, ByRef b As Variant, ByVal rev As Boolean) As Boolean
    If rev Then
        KeyBefore = PyLess(b, a)
    Else
        KeyBefore = PyLess(a, b)
    End If
End Function

Private Sub MergeSortIdx(ByRef idx() As Long, ByRef tmp() As Long, ByRef keys() As Variant, ByVal lo As Long, ByVal hi As Long, ByVal rev As Boolean)
    Dim md As Long, i As Long, j As Long, k As Long, t As Long
    If hi - lo < 12 Then
        For i = lo + 1 To hi
            t = idx(i)
            j = i - 1
            Do While j >= lo
                If Not KeyBefore(keys(t), keys(idx(j)), rev) Then Exit Do
                idx(j + 1) = idx(j)
                j = j - 1
            Loop
            idx(j + 1) = t
        Next
        Exit Sub
    End If
    md = (lo + hi) \ 2
    MergeSortIdx idx, tmp, keys, lo, md, rev
    MergeSortIdx idx, tmp, keys, md + 1, hi, rev
    If Not KeyBefore(keys(idx(md + 1)), keys(idx(md)), rev) Then Exit Sub
    i = lo
    j = md + 1
    k = lo
    Do While i <= md And j <= hi
        If KeyBefore(keys(idx(j)), keys(idx(i)), rev) Then
            tmp(k) = idx(j)
            j = j + 1
        Else
            tmp(k) = idx(i)
            i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= md
        tmp(k) = idx(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = idx(j)
        j = j + 1
        k = k + 1
    Loop
    For k = lo To hi
        idx(k) = tmp(k)
    Next
End Sub

Public Sub SortList(ByVal lst As PyObj, ByRef keyFn As Variant, ByVal rev As Boolean)
    Dim n As Long, items() As Variant, keys() As Variant, idx() As Long, tmp() As Long, i As Long, res() As Variant, kv As Variant
    n = lst.Count
    If n < 2 Then Exit Sub
    lst.GetArray items
    If IsEmpty(keyFn) Then
        lst.GetArray keys
    Else
        ReDim keys(0 To n - 1)
        For i = 0 To n - 1
            CallValue1 keyFn, items(i), kv
            SetV keys(i), kv
        Next
    End If
    ReDim idx(0 To n - 1)
    ReDim tmp(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = i
    Next
    MergeSortIdx idx, tmp, keys, 0, n - 1, rev
    ReDim res(0 To n - 1)
    For i = 0 To n - 1
        SetV res(i), items(idx(i))
    Next
    lst.SetArray res, n
End Sub

Private Sub SortArgs(ByVal kw As PyObj, ByRef keyFn As Variant, ByRef rev As Boolean)
    Dim v As Variant
    keyFn = Empty
    rev = False
    If KwGet(kw, "key", v) Then SetV keyFn, v
    If KwGet(kw, "reverse", v) Then rev = Truthy(v)
End Sub

' ---------------------------------------------------------------------------
'  string formatting
' ---------------------------------------------------------------------------
Private Function RoundHalfEven(ByVal d As Double, ByVal prec As Long) As Double
    If prec <= 15 And Abs(d) < 1E+15 Then
        RoundHalfEven = Round(d, prec)
    Else
        RoundHalfEven = d
    End If
End Function

Public Function FixedStr(ByVal d As Double, ByVal prec As Long) As String
    Dim s As String
    If prec > 30 Then prec = 30
    If prec = 0 Then
        s = Format$(RoundHalfEven(d, 0), "0")
    Else
        s = Format$(RoundHalfEven(d, prec), "0." & String$(prec, "0"))
    End If
    If gDecSep <> "." Then s = Replace(s, gDecSep, ".")
    FixedStr = s
End Function

Public Function SciStr(ByVal d As Double, ByVal prec As Long, ByVal upper As Boolean) As String
    Dim s As String, p As Long, mant As String, ex As Long
    If prec > 30 Then prec = 30
    If prec = 0 Then
        s = Format$(d, "0E+00")
    Else
        s = Format$(d, "0." & String$(prec, "0") & "E+00")
    End If
    If gDecSep <> "." Then s = Replace(s, gDecSep, ".")
    p = InStr(1, s, "E", vbTextCompare)
    mant = Left$(s, p - 1)
    ex = CLng(Val(Mid$(s, p + 1)))
    If upper Then
        s = mant & "E"
    Else
        s = mant & "e"
    End If
    If ex < 0 Then
        s = s & "-"
    Else
        s = s & "+"
    End If
    SciStr = s & Format$(Abs(ex), "00")
End Function

Private Function StripZeros(ByVal s As String) As String
    Dim p As Long, mant As String, rest As String
    p = InStr(1, s, "e", vbTextCompare)
    If p > 0 Then
        mant = Left$(s, p - 1)
        rest = Mid$(s, p)
    Else
        mant = s
    End If
    If InStr(1, mant, ".") > 0 Then
        Do While Right$(mant, 1) = "0"
            mant = Left$(mant, Len(mant) - 1)
        Loop
        If Right$(mant, 1) = "." Then mant = Left$(mant, Len(mant) - 1)
    End If
    StripZeros = mant & rest
End Function

Private Function GenStr(ByVal d As Double, ByVal prec As Long, ByVal alt As Boolean) As String
    Dim s As String, ex As Long, p As Long
    If prec <= 0 Then prec = 1
    If d = 0 Then
        ex = 0
    Else
        s = SciStr(d, prec - 1, False)
        p = InStr(1, s, "e")
        ex = CLng(Val(Mid$(s, p + 1)))
    End If
    If ex < -4 Or ex >= prec Then
        s = SciStr(d, prec - 1, False)
    Else
        s = FixedStr(d, prec - 1 - ex)
    End If
    If Not alt Then s = StripZeros(s)
    GenStr = s
End Function

Private Function GroupDigits(ByVal s As String, ByVal sep As String) As String
    Dim intPart As String, rest As String, p As Long, r As String, n As Long
    p = InStr(1, s, ".")
    If p > 0 Then
        intPart = Left$(s, p - 1)
        rest = Mid$(s, p)
    Else
        intPart = s
    End If
    n = Len(intPart)
    Do While n > 3
        r = sep & Right$(intPart, 3) & r
        intPart = Left$(intPart, n - 3)
        n = n - 3
    Loop
    GroupDigits = intPart & r & rest
End Function

Private Function PadStr(ByVal body As String, ByVal prefix As String, ByVal fill As String, ByVal align As String, ByVal wid As Long) As String
    Dim total As Long, padn As Long, lft As Long
    total = Len(prefix) + Len(body)
    If wid <= total Then
        PadStr = prefix & body
        Exit Function
    End If
    padn = wid - total
    Select Case align
    Case "<": PadStr = prefix & body & String$(padn, fill)
    Case "^"
        lft = padn \ 2
        PadStr = String$(lft, fill) & prefix & body & String$(padn - lft, fill)
    Case "=": PadStr = prefix & String$(padn, fill) & body
    Case Else: PadStr = String$(padn, fill) & prefix & body
    End Select
End Function

' format(value, spec) - the format specification mini-language
Public Function FormatValue(ByRef v As Variant, ByVal spec As String) As String
    Dim fill As String, align As String, sgn As String, alt As Boolean, zero As Boolean
    Dim wid As Long, grp As String, prec As Long, typ As String, i As Long, n As Long, c As String
    Dim body As String, neg As Boolean, d As Double, pre As String, r As Variant, x As Variant, sg As String
    Dim sv As Variant
    If IsKind(v, OK_INST) Then
        sv = spec
        If CallDunder1(v, "__format__", sv, r) Then
            FormatValue = PyStr(r)
            Exit Function
        End If
        sv = PyStr(v)
        FormatValue = FormatValue(sv, spec)
        Exit Function
    End If
    fill = " "
    prec = -1
    sgn = "-"
    n = Len(spec)
    i = 1
    If n >= 2 Then
        If InStr(1, "<>^=", Mid$(spec, 2, 1)) > 0 Then
            fill = Mid$(spec, 1, 1)
            align = Mid$(spec, 2, 1)
            i = 3
        End If
    End If
    If Len(align) = 0 And n >= 1 Then
        If InStr(1, "<>^=", Mid$(spec, 1, 1)) > 0 Then
            align = Mid$(spec, 1, 1)
            i = 2
        End If
    End If
    c = Mid$(spec, i, 1)
    If c = "+" Or c = "-" Or c = " " Then
        sgn = c
        i = i + 1
    End If
    If Mid$(spec, i, 1) = "#" Then
        alt = True
        i = i + 1
    End If
    If Mid$(spec, i, 1) = "0" Then
        zero = True
        i = i + 1
    End If
    Do While i <= n
        c = Mid$(spec, i, 1)
        If c < "0" Or c > "9" Then Exit Do
        wid = wid * 10 + (Asc(c) - 48)
        i = i + 1
    Loop
    c = Mid$(spec, i, 1)
    If c = "," Or c = "_" Then
        grp = c
        i = i + 1
    End If
    If Mid$(spec, i, 1) = "." Then
        i = i + 1
        prec = 0
        Do While i <= n
            c = Mid$(spec, i, 1)
            If c < "0" Or c > "9" Then Exit Do
            prec = prec * 10 + (Asc(c) - 48)
            i = i + 1
        Loop
    End If
    If i <= n Then
        typ = Mid$(spec, i, 1)
        i = i + 1
    End If
    If i <= n Then ThrowErr "ValueError", "Invalid format specifier '" & spec & "'"
    If zero And Len(align) = 0 Then
        fill = "0"
        align = "="
    End If
    If Len(typ) = 0 Then
        If VarType(v) = vbString Or VarType(v) = vbBoolean Or Not IsNum(v) Then
            typ = "s"
        ElseIf IsIntType(v) Then
            typ = "d"
        ElseIf prec >= 0 Then
            typ = "g"
        Else
            typ = "r"
        End If
    End If
    Select Case typ
    Case "s"
        body = PyStr(v)
        If prec >= 0 Then body = Left$(body, prec)
        If Len(align) = 0 Then align = "<"
        FormatValue = PadStr(body, "", fill, align, wid)
        Exit Function
    Case "d", "n"
        If Not IsIntType(v) Then ThrowErr "ValueError", "Unknown format code 'd' for object of type '" & TypeNameOf(v) & "'"
        x = IntOf(v)
        neg = (x < 0)
        If neg Then x = IntNeg(x)
        body = CStr(x)
        If Len(grp) > 0 Then body = GroupDigits(body, grp)
    Case "f", "F", "%"
        If Not IsNum(v) Then ThrowErr "ValueError", "Unknown format code '" & typ & "' for object of type '" & TypeNameOf(v) & "'"
        d = NumToDbl(v)
        If typ = "%" Then d = d * 100
        neg = (d < 0)
        d = Abs(d)
        If prec < 0 Then prec = 6
        If d >= PY_INF Then
            body = "inf"
        Else
            body = FixedStr(d, prec)
        End If
        If Len(grp) > 0 Then body = GroupDigits(body, grp)
        If typ = "%" Then body = body & "%"
    Case "e", "E"
        If Not IsNum(v) Then ThrowErr "ValueError", "Unknown format code '" & typ & "' for object of type '" & TypeNameOf(v) & "'"
        d = NumToDbl(v)
        neg = (d < 0)
        d = Abs(d)
        If prec < 0 Then prec = 6
        body = SciStr(d, prec, typ = "E")
    Case "g", "G"
        If Not IsNum(v) Then ThrowErr "ValueError", "Unknown format code '" & typ & "' for object of type '" & TypeNameOf(v) & "'"
        d = NumToDbl(v)
        neg = (d < 0)
        d = Abs(d)
        If prec < 0 Then prec = 6
        body = GenStr(d, prec, alt)
        If typ = "G" Then body = UCase$(body)
        If Len(grp) > 0 Then body = GroupDigits(body, grp)
    Case "r"
        d = NumToDbl(v)
        neg = (d < 0)
        body = FloatRepr(Abs(d))
        If Len(grp) > 0 Then body = GroupDigits(body, grp)
    Case "x", "X", "o", "b"
        If Not IsIntType(v) Then ThrowErr "ValueError", "Unknown format code '" & typ & "' for object of type '" & TypeNameOf(v) & "'"
        x = IntOf(v)
        neg = (x < 0)
        If neg Then x = IntNeg(x)
        Select Case typ
        Case "x", "X"
            body = IntToBase(x, 16)
            If alt Then pre = "0x"
        Case "o"
            body = IntToBase(x, 8)
            If alt Then pre = "0o"
        Case "b"
            body = IntToBase(x, 2)
            If alt Then pre = "0b"
        End Select
        If typ = "X" Then
            body = UCase$(body)
            pre = UCase$(pre)
        End If
    Case "c"
        body = ChrW(IntToLong(v, "%c"))
    Case Else
        ThrowErr "ValueError", "Unknown format code '" & typ & "' for object of type '" & TypeNameOf(v) & "'"
    End Select
    If neg Then
        sg = "-"
    ElseIf sgn = "+" Then
        sg = "+"
    ElseIf sgn = " " Then
        sg = " "
    End If
    If Len(align) = 0 Then align = ">"
    FormatValue = PadStr(body, sg & pre, fill, align, wid)
End Function

' "..." % args
Public Function PercentFormat(ByVal fmt As String, ByRef argsV As Variant) As String
    Dim items As PyObj, mapping As PyObj, ai As Long, i As Long, n As Long, c As String, p As Long, j As Long
    Dim parts() As String, np As Long, flags As String, wid As String, prec As String, conv As String
    Dim spec As String, v As Variant, key As String, kv As Variant, s As String, x As Variant
    If IsKind(argsV, OK_TUPLE) Then
        Set items = argsV
    Else
        Set items = NewObj(OK_TUPLE)
        items.Push argsV
        If IsKind(argsV, OK_DICT) Then Set mapping = argsV
    End If
    n = Len(fmt)
    ReDim parts(0 To 15)
    i = 1
    Do While i <= n
        p = InStr(i, fmt, "%")
        If p = 0 Then
            AddStrPart parts, np, Mid$(fmt, i)
            Exit Do
        End If
        AddStrPart parts, np, Mid$(fmt, i, p - i)
        i = p + 1
        If Mid$(fmt, i, 1) = "%" Then
            AddStrPart parts, np, "%"
            i = i + 1
        Else
            key = vbNullString
            If Mid$(fmt, i, 1) = "(" Then
                j = InStr(i, fmt, ")")
                If j = 0 Then ThrowErr "ValueError", "incomplete format key"
                key = Mid$(fmt, i + 1, j - i - 1)
                i = j + 1
            End If
            flags = vbNullString
            Do While i <= n
                c = Mid$(fmt, i, 1)
                If InStr(1, "-+ 0#", c) = 0 Then Exit Do
                flags = flags & c
                i = i + 1
            Loop
            wid = vbNullString
            Do While i <= n
                c = Mid$(fmt, i, 1)
                If c < "0" Or c > "9" Then Exit Do
                wid = wid & c
                i = i + 1
            Loop
            prec = vbNullString
            If Mid$(fmt, i, 1) = "." Then
                i = i + 1
                prec = "."
                Do While i <= n
                    c = Mid$(fmt, i, 1)
                    If c < "0" Or c > "9" Then Exit Do
                    prec = prec & c
                    i = i + 1
                Loop
                If prec = "." Then prec = ".0"
            End If
            conv = Mid$(fmt, i, 1)
            i = i + 1
            If Len(conv) = 0 Then ThrowErr "ValueError", "incomplete format"
            If Len(key) > 0 Then
                If mapping Is Nothing Then ThrowErr "TypeError", "format requires a mapping"
                kv = key
                GetItem ToVar(mapping), kv, v
            Else
                If ai >= items.Count Then ThrowErr "TypeError", "not enough arguments for format string"
                items.GetInto ai, v
                ai = ai + 1
            End If
            ' build a format() spec: [align][sign][#][0][width][.prec][type]
            spec = vbNullString
            If InStr(1, flags, "-") > 0 Then
                spec = "<"
            ElseIf InStr(1, flags, "0") = 0 Or conv = "s" Or conv = "r" Or conv = "a" Or conv = "c" Then
                spec = ">"
            End If
            If InStr(1, flags, "+") > 0 Then
                spec = spec & "+"
            ElseIf InStr(1, flags, " ") > 0 Then
                spec = spec & " "
            End If
            If InStr(1, flags, "#") > 0 Then spec = spec & "#"
            If Len(spec) = 0 Or spec = "+" Or spec = " " Or spec = "#" Or spec = "+#" Or spec = " #" Then
                If InStr(1, flags, "0") > 0 And InStr(1, flags, "-") = 0 Then spec = spec & "0"
            End If
            Select Case conv
            Case "s"
                x = PyStr(v)
                s = FormatValue(x, spec & wid & prec & "s")
            Case "r", "a"
                x = PyRepr(v)
                s = FormatValue(x, spec & wid & prec & "s")
            Case "d", "i", "u"
                If IsFloatType(v) Then
                    x = IntFromDouble(Fix(CDbl(v)))
                Else
                    x = v
                End If
                If Not IsIntType(x) Then ThrowErr "TypeError", "%d format: a real number is required, not " & TypeNameOf(v)
                s = FormatValue(x, spec & wid & "d")
            Case "f", "F", "e", "E", "g", "G"
                If Not IsNum(v) Then ThrowErr "TypeError", "must be real number, not " & TypeNameOf(v)
                If Len(prec) = 0 Then prec = ".6"
                s = FormatValue(v, spec & wid & prec & conv)
            Case "x", "X", "o"
                s = FormatValue(v, spec & wid & conv)
            Case "c"
                If VarType(v) = vbString Then
                    x = v
                Else
                    x = ChrW(IntToLong(v, "%c"))
                End If
                s = FormatValue(x, spec & wid & "s")
            Case Else
                ThrowErr "ValueError", "unsupported format character '" & conv & "'"
            End Select
            AddStrPart parts, np, s
        End If
    Loop
    If mapping Is Nothing And ai < items.Count Then ThrowErr "TypeError", "not all arguments converted during string formatting"
    If np = 0 Then Exit Function
    ReDim Preserve parts(0 To np - 1)
    PercentFormat = Join(parts, vbNullString)
End Function

Private Sub AddStrPart(ByRef parts() As String, ByRef np As Long, ByVal s As String)
    If np > UBound(parts) Then
        ReDim Preserve parts(0 To 2 * np + 1)
    End If
    parts(np) = s
    np = np + 1
End Sub

' "...".format(*args, **kw)
Public Function StrFormatMethod(ByVal fmt As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj) As String
    Dim i As Long, n As Long, c As String, parts() As String, np As Long, auto As Long, j As Long, depth As Long
    Dim field As String, fname As String, conv As String, spec As String, v As Variant, p As Long, k As Long
    Dim bname As String, rest As String, idx As Long, acc As String, kv As Variant, r As Variant, ok As Boolean
    n = Len(fmt)
    ReDim parts(0 To 15)
    i = 1
    Do While i <= n
        c = Mid$(fmt, i, 1)
        If c = "{" Then
            If Mid$(fmt, i + 1, 1) = "{" Then
                AddStrPart parts, np, "{"
                i = i + 2
            Else
                j = i + 1
                depth = 1
                Do While j <= n
                    c = Mid$(fmt, j, 1)
                    If c = "{" Then
                        depth = depth + 1
                    ElseIf c = "}" Then
                        depth = depth - 1
                        If depth = 0 Then Exit Do
                    End If
                    j = j + 1
                Loop
                If j > n Then ThrowErr "ValueError", "Single '{' encountered in format string"
                field = Mid$(fmt, i + 1, j - i - 1)
                spec = vbNullString
                conv = vbNullString
                p = InStr(1, field, ":")
                If p > 0 Then
                    spec = Mid$(field, p + 1)
                    field = Left$(field, p - 1)
                End If
                p = InStr(1, field, "!")
                If p > 0 Then
                    conv = Mid$(field, p + 1)
                    field = Left$(field, p - 1)
                End If
                If InStr(1, spec, "{") > 0 Then spec = StrFormatMethod(spec, args, na, kw)
                ' split the field name into bname and accessors
                k = 1
                Do While k <= Len(field)
                    c = Mid$(field, k, 1)
                    If c = "." Or c = "[" Then Exit Do
                    k = k + 1
                Loop
                bname = Left$(field, k - 1)
                rest = Mid$(field, k)
                If Len(bname) = 0 Then
                    If auto >= na Then ThrowErr "IndexError", "Replacement index " & auto & " out of range for positional args tuple"
                    SetV v, args(auto)
                    auto = auto + 1
                ElseIf IsNumeric(bname) And InStr(1, bname, ".") = 0 Then
                    idx = CLng(bname)
                    If idx >= na Then ThrowErr "IndexError", "Replacement index " & idx & " out of range for positional args tuple"
                    SetV v, args(idx)
                Else
                    If Not KwGet(kw, bname, v) Then RaiseExc MakeExcV("KeyError", bname)
                End If
                Do While Len(rest) > 0
                    If Left$(rest, 1) = "." Then
                        k = 2
                        Do While k <= Len(rest)
                            c = Mid$(rest, k, 1)
                            If c = "." Or c = "[" Then Exit Do
                            k = k + 1
                        Loop
                        acc = Mid$(rest, 2, k - 2)
                        rest = Mid$(rest, k)
                        PyGetAttr v, acc, r
                        SetV v, r
                    Else
                        k = InStr(1, rest, "]")
                        If k = 0 Then ThrowErr "ValueError", "Missing ']' in format string"
                        acc = Mid$(rest, 2, k - 2)
                        rest = Mid$(rest, k + 1)
                        kv = ParseIntStr(acc, 10, ok)
                        If Not ok Then kv = acc
                        GetItem v, kv, r
                        SetV v, r
                    End If
                Loop
                If conv = "r" Or conv = "a" Then
                    v = PyRepr(v)
                ElseIf conv = "s" Then
                    v = PyStr(v)
                End If
                If Len(spec) = 0 Then
                    AddStrPart parts, np, PyStr(v)
                Else
                    AddStrPart parts, np, FormatValue(v, spec)
                End If
                i = j + 1
            End If
        ElseIf c = "}" Then
            If Mid$(fmt, i + 1, 1) = "}" Then
                AddStrPart parts, np, "}"
                i = i + 2
            Else
                ThrowErr "ValueError", "Single '}' encountered in format string"
            End If
        Else
            j = i
            Do While j <= n
                c = Mid$(fmt, j, 1)
                If c = "{" Or c = "}" Then Exit Do
                j = j + 1
            Loop
            AddStrPart parts, np, Mid$(fmt, i, j - i)
            i = j
        End If
    Loop
    If np = 0 Then Exit Function
    ReDim Preserve parts(0 To np - 1)
    StrFormatMethod = Join(parts, vbNullString)
End Function

' ---------------------------------------------------------------------------
'  built-in functions
' ---------------------------------------------------------------------------
Public Sub CallBuiltin(ByVal fo As PyObj, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim c As Long
    c = fo.Code
    If c < BI_ISINSTANCE Then
        CallBuiltinA fo, c, args, na, kw, out
    ElseIf c < 60 Then
        CallBuiltinA2 fo, c, args, na, kw, out
    ElseIf c < 100 Then
        CallBuiltinB fo, c, args, na, kw, out
    ElseIf c < 150 Then
        CallMath c, args, na, kw, out
    ElseIf c >= BI_DC_DATACLASS And c < 240 Then
        CallDataclass fo, c, args, na, kw, out
    ElseIf c < BI_HOST_FIRST Then
        CallLib fo, c, args, na, kw, out
    ElseIf c <= BI_HOST_LAST Then
        HostCall c, args, na, kw, out
    Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End If
End Sub

Private Sub CallBuiltinA(ByVal fo As PyObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim i As Long, s As String, v As Variant, r As PyObj, x As Variant, y As Variant, it As PyObj, ok As Boolean
    Dim parts() As String, sep As String, endS As String, lst As PyObj, d As Double, o As PyObj, keyFn As Variant
    Dim rev As Boolean, n As Long, st As Long, en As Long, stp As Long, j As Long, its() As PyObj, t As PyObj
    Dim q As Variant, rm As Variant, bs As Long
    Select Case c
    Case BI_PRINT
        sep = " "
        endS = vbLf
        If KwGet(kw, "sep", v) Then
            If Not IsEmpty(v) Then sep = PyStr(v)
        End If
        If KwGet(kw, "end", v) Then
            If Not IsEmpty(v) Then endS = PyStr(v)
        End If
        If na = 0 Then
            OutWrite endS
        Else
            ReDim parts(0 To na - 1)
            For i = 0 To na - 1
                parts(i) = PyStr(args(i))
            Next
            OutWrite Join(parts, sep) & endS
        End If
        out = Empty
    Case BI_LEN
        NeedArgs "len", na, 1, 1
        out = PyLen(args(0))
    Case BI_RANGE
        NeedArgs "range", na, 1, 3
        If na = 1 Then
            st = 0
            en = IntToLong(args(0), "range() argument")
            stp = 1
        Else
            st = IntToLong(args(0), "range() argument")
            en = IntToLong(args(1), "range() argument")
            stp = 1
            If na = 3 Then stp = IntToLong(args(2), "range() argument")
            If stp = 0 Then ThrowErr "ValueError", "range() arg 3 must not be zero"
        End If
        Set out = NewRange(st, en, stp)
    Case BI_INT
        If na = 0 Then
            out = 0&
            Exit Sub
        End If
        bs = 10
        If ArgOr(args, na, 1, kw, "base", v) Then bs = IntToLong(v, "base")
        SetV x, args(0)
        If VarType(x) = vbString Then
            v = ParseIntStr(x, bs, ok)
            If Not ok Then ThrowErr "ValueError", "invalid literal for int() with base " & bs & ": " & StrRepr(CStr(x))
            out = v
        ElseIf IsIntType(x) Then
            out = IntOf(x)
        ElseIf IsFloatType(x) Then
            d = CDbl(x)
            If d >= PY_INF Or d <= -PY_INF Then ThrowErr "OverflowError", "cannot convert float infinity to integer"
            out = IntFromDouble(Fix(d))
        ElseIf CallDunder0(x, "__int__", v) Then
            out = v
        Else
            ThrowErr "TypeError", "int() argument must be a string or a number, not '" & TypeNameOf(x) & "'"
        End If
    Case BI_FLOAT
        If na = 0 Then
            out = 0#
            Exit Sub
        End If
        SetV x, args(0)
        If VarType(x) = vbString Then
            s = LCase$(Trim$(x))
            If s = "inf" Or s = "+inf" Or s = "infinity" Then
                out = PY_INF
            ElseIf s = "-inf" Or s = "-infinity" Then
                out = -PY_INF
            ElseIf s = "nan" Then
                ThrowErr "ValueError", "NaN is not supported"
            Else
                d = ParseFloatStr(x, ok)
                If Not ok Then ThrowErr "ValueError", "could not convert string to float: " & StrRepr(CStr(x))
                out = d
            End If
        ElseIf IsNum(x) Then
            If VarType(x) = vbDecimal Then
                out = CDbl(x)
            Else
                out = NumToDbl(x)
            End If
        ElseIf CallDunder0(x, "__float__", v) Then
            out = v
        Else
            ThrowErr "TypeError", "float() argument must be a string or a number, not '" & TypeNameOf(x) & "'"
        End If
    Case BI_STR
        If na = 0 Then
            out = vbNullString
        Else
            out = PyStr(args(0))
        End If
    Case BI_BOOL
        If na = 0 Then
            out = False
        Else
            out = Truthy(args(0))
        End If
    Case BI_LIST
        If na = 0 Then
            Set out = NewList()
        Else
            Set out = NewListFrom(args(0))
        End If
    Case BI_TUPLE
        If na = 0 Then
            Set out = NewObj(OK_TUPLE)
        ElseIf IsKind(args(0), OK_TUPLE) Then
            SetV out, args(0)
        Else
            Set r = NewListFrom(args(0))
            r.Kind = OK_TUPLE
            Set out = r
        End If
    Case BI_SET, BI_FROZENSET
        Set r = NewObj(OK_SET)
        If na > 0 Then
            Set it = GetIter(args(0))
            Do While IterNext(it, x)
                SetAdd r, x
            Loop
        End If
        Set out = r
    Case BI_DICT
        Set r = NewDict()
        If na > 0 Then DictUpdateFrom r, args(0)
        If Not (kw Is Nothing) Then DictUpdate r, kw
        Set out = r
    Case BI_ABS
        NeedArgs "abs", na, 1, 1
        SetV x, args(0)
        If IsFloatType(x) Then
            out = Abs(CDbl(x))
        ElseIf IsIntType(x) Then
            x = IntOf(x)
            If x < 0 Then
                out = IntNeg(x)
            Else
                out = x
            End If
        ElseIf CallDunder0(x, "__abs__", v) Then
            SetV out, v
        Else
            ThrowErr "TypeError", "bad operand type for abs(): '" & TypeNameOf(x) & "'"
        End If
    Case BI_MIN
        MinMax "min", args, na, kw, False, out
    Case BI_MAX
        MinMax "max", args, na, kw, True, out
    Case BI_SUM
        NeedArgs "sum", na, 1, 2
        If Not ArgOr(args, na, 1, kw, "start", x) Then x = 0&
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            BinOp OP_ADD, x, v, y
            SetV x, y
        Loop
        SetV out, x
    Case BI_SORTED
        NeedArgs "sorted", na, 1, 1
        Set r = NewListFrom(args(0))
        SortArgs kw, keyFn, rev
        SortList r, keyFn, rev
        Set out = r
    Case BI_REVERSED
        NeedArgs "reversed", na, 1, 1
        If IsKind(args(0), OK_INST) Then
            If CallDunder0(args(0), "__reversed__", v) Then
                SetV out, v
                Exit Sub
            End If
        End If
        Set r = NewListFrom(args(0))
        r.ReverseItems
        Set out = GetIter(ToVar(r))
    Case BI_ENUMERATE
        NeedArgs "enumerate", na, 1, 2
        If ArgOr(args, na, 1, kw, "start", x) Then x = IntOf(x) Else x = 0&
        Set r = NewList()
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            r.Push ToVar(NewTuple2(x, v))
            x = IntAdd(x, 1&)
        Loop
        Set out = r
    Case BI_ZIP
        Set r = NewList()
        If na > 0 Then
            ReDim its(0 To na - 1)
            For i = 0 To na - 1
                Set its(i) = GetIter(args(i))
            Next
            Do
                Set t = NewObj(OK_TUPLE)
                For i = 0 To na - 1
                    If Not IterNext(its(i), v) Then GoTo ZipDone
                    t.Push v
                Next
                r.Push ToVar(t)
            Loop
        End If
ZipDone:
        Set out = r
    Case BI_MAP
        If na < 2 Then ThrowErr "TypeError", "map() must have at least two arguments."
        Set r = NewList()
        ReDim its(0 To na - 2)
        For i = 1 To na - 1
            Set its(i - 1) = GetIter(args(i))
        Next
        Dim margs() As Variant
        ReDim margs(0 To na - 2)
        Do
            For i = 0 To na - 2
                If Not IterNext(its(i), v) Then GoTo MapDone
                SetV margs(i), v
            Next
            CallValue args(0), margs, na - 1, Nothing, y
            r.Push y
        Loop
MapDone:
        Set out = r
    Case BI_FILTER
        NeedArgs "filter", na, 2, 2
        Set r = NewList()
        Set it = GetIter(args(1))
        Do While IterNext(it, v)
            If IsEmpty(args(0)) Then
                If Truthy(v) Then r.Push v
            Else
                CallValue1 args(0), v, y
                If Truthy(y) Then r.Push v
            End If
        Loop
        Set out = r
    Case BI_ANY
        NeedArgs "any", na, 1, 1
        out = False
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            If Truthy(v) Then
                out = True
                Exit Do
            End If
        Loop
    Case BI_ALL
        NeedArgs "all", na, 1, 1
        out = True
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            If Not Truthy(v) Then
                out = False
                Exit Do
            End If
        Loop
    Case BI_ROUND
        NeedArgs "round", na, 1, 2
        SetV x, args(0)
        If Not ArgOr(args, na, 1, kw, "ndigits", y) Then y = Empty
        If IsEmpty(y) Then
            If IsIntType(x) Then
                out = IntOf(x)
            ElseIf IsFloatType(x) Then
                out = IntFromDouble(Round(CDbl(x)))
            ElseIf CallDunder0(x, "__round__", v) Then
                SetV out, v
            Else
                ThrowErr "TypeError", "type " & TypeNameOf(x) & " doesn't define __round__ method"
            End If
        Else
            n = IntToLong(y, "ndigits")
            If IsIntType(x) Then
                If n >= 0 Then
                    out = IntOf(x)
                Else
                    v = IntPow(10&, -n)
                    IntDivMod x, v, q, rm
                    If IntMul(rm, 2&) > v Or (IntMul(rm, 2&) = v And (q Mod 2) <> 0) Then q = IntAdd(q, 1&)
                    out = IntMul(q, v)
                End If
            Else
                d = NumArg(x, "round")
                If n >= 0 Then
                    If n > 15 Then
                        out = d
                    Else
                        out = Round(d, n)
                    End If
                Else
                    out = Round(d / (10 ^ -n)) * (10 ^ -n)
                End If
            End If
        End If
    Case BI_POW
        NeedArgs "pow", na, 2, 3
        If na = 3 Then
            If Not IsEmpty(args(2)) Then
                out = PowMod(args(0), args(1), args(2))
                Exit Sub
            End If
        End If
        BinOp OP_POW, args(0), args(1), out
    Case BI_DIVMOD
        NeedArgs "divmod", na, 2, 2
        BinOp OP_FLOORDIV, args(0), args(1), x
        BinOp OP_MOD, args(0), args(1), y
        Set out = NewTuple2(x, y)
    Case Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End Select
End Sub

Private Sub CallBuiltinA2(ByVal fo As PyObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim i As Long, s As String, v As Variant, r As PyObj, x As Variant, y As Variant, it As PyObj, ok As Boolean
    Dim parts() As String, sep As String, endS As String, lst As PyObj, d As Double, o As PyObj, keyFn As Variant
    Dim rev As Boolean, n As Long, st As Long, en As Long, stp As Long, j As Long, its() As PyObj, t As PyObj
    Dim q As Variant, rm As Variant, bs As Long
    Select Case c
    Case BI_ISINSTANCE
        NeedArgs "isinstance", na, 2, 2
        out = IsInstanceOf(args(0), args(1))
    Case BI_ISSUBCLASS
        NeedArgs "issubclass", na, 2, 2
        out = IsSubclassV(args(0), args(1))
    Case BI_TYPE
        NeedArgs "type", na, 1, 1
        TypeOfValue args(0), out
    Case BI_REPR
        NeedArgs "repr", na, 1, 1
        out = PyRepr(args(0))
    Case BI_CHR
        NeedArgs "chr", na, 1, 1
        n = IntToLong(args(0), "chr() argument")
        If n < 0 Or n > &H10FFFF Then ThrowErr "ValueError", "chr() arg not in range(0x110000)"
        If n > 65535 Then
            n = n - 65536
            out = ChrW(&HD800& + (n \ 1024)) & ChrW(&HDC00& + (n And 1023))
        Else
            out = ChrW(n)
        End If
    Case BI_ORD
        NeedArgs "ord", na, 1, 1
        s = ArgStr(args(0), "ord() argument")
        If Len(s) = 2 Then
            n = UCode(Left$(s, 1))
            If n >= &HD800& And n <= &HDBFF& Then
                out = CLng((n - &HD800&) * 1024 + (UCode(Right$(s, 1)) - &HDC00&) + 65536)
                Exit Sub
            End If
        End If
        If Len(s) <> 1 Then ThrowErr "TypeError", "ord() expected a character, but string of length " & Len(s) & " found"
        out = UCode(s)
    Case BI_HEX, BI_BIN, BI_OCT
        NeedArgs fo.PName, na, 1, 1
        If Not IsIntType(args(0)) Then ThrowErr "TypeError", "'" & TypeNameOf(args(0)) & "' object cannot be interpreted as an integer"
        x = IntOf(args(0))
        s = vbNullString
        If x < 0 Then
            s = "-"
            x = IntNeg(x)
        End If
        Select Case c
        Case BI_HEX: out = s & "0x" & IntToBase(x, 16)
        Case BI_BIN: out = s & "0b" & IntToBase(x, 2)
        Case Else: out = s & "0o" & IntToBase(x, 8)
        End Select
    Case BI_INPUT
        If na > 0 Then
            s = PyStr(args(0))
        Else
            s = vbNullString
        End If
        out = HostInput(s)
        OutWrite s & out & vbLf
    Case BI_ID
        NeedArgs "id", na, 1, 1
        If IsObject(args(0)) Then
            Set o = args(0)
            out = o.Id
        Else
            out = HashStr(KeyStr(args(0)))
        End If
    Case BI_HASH
        NeedArgs "hash", na, 1, 1
        s = KeyStr(args(0))
        If VarType(args(0)) = vbLong Then
            out = args(0)
        ElseIf Left$(s, 1) = "i" And Len(s) < 11 Then
            out = CLng(Mid$(s, 2))
        ElseIf Left$(s, 1) = "h" Then
            out = HashStr31(s)
            If IsKind(args(0), OK_INST) Then
                If CallDunder0(args(0), "__hash__", v) Then SetV out, v
            End If
        Else
            out = HashStr31(s)
        End If
    Case BI_CALLABLE
        NeedArgs "callable", na, 1, 1
        out = False
        If IsObject(args(0)) Then
            Set o = args(0)
            Select Case o.Kind
            Case OK_FUNC, OK_BUILTIN, OK_BOUND, OK_BMETHOD, OK_CLASS
                out = True
            Case OK_INST
                out = ClassLookup(o.Cls, "__call__", HashStr("__call__"), v)
            End Select
        End If
    Case BI_GETATTR
        NeedArgs "getattr", na, 2, 3
        s = ArgStr(args(1), "attribute name")
        If TryGetAttr(args(0), s, v) Then
            SetV out, v
        ElseIf na = 3 Then
            SetV out, args(2)
        Else
            PyGetAttr args(0), s, out
        End If
    Case BI_SETATTR
        NeedArgs "setattr", na, 3, 3
        PySetAttr args(0), ArgStr(args(1), "attribute name"), args(2)
        out = Empty
    Case BI_HASATTR
        NeedArgs "hasattr", na, 2, 2
        out = TryGetAttr(args(0), ArgStr(args(1), "attribute name"), v)
    Case BI_DELATTR
        NeedArgs "delattr", na, 2, 2
        PyDelAttr args(0), ArgStr(args(1), "attribute name")
        out = Empty
    Case BI_FORMAT
        NeedArgs "format", na, 1, 2
        If na = 2 Then
            s = ArgStr(args(1), "format spec")
        Else
            s = vbNullString
        End If
        If Len(s) = 0 Then
            out = PyStr(args(0))
        Else
            out = FormatValue(args(0), s)
        End If
    Case BI_ITER
        NeedArgs "iter", na, 1, 1
        Set out = GetIter(args(0))
    Case BI_NEXT
        NeedArgs "next", na, 1, 2
        ok = False
        If IsKind(args(0), OK_ITER) Then
            Set o = args(0)
            ok = IterNext(o, v)
        ElseIf IsKind(args(0), OK_INST) Then
            Set o = args(0)
            ok = UserIterNext(o, v)
        Else
            ThrowErr "TypeError", "'" & TypeNameOf(args(0)) & "' object is not an iterator"
        End If
        If ok Then
            SetV out, v
        ElseIf na = 2 Then
            SetV out, args(1)
        Else
            RaiseExc NewExcInst(ExcClass("StopIteration"), NewObj(OK_TUPLE))
        End If
    Case BI_STATICMETHOD, BI_CLASSMETHOD
        NeedArgs fo.PName, na, 1, 1
        If Not IsKind(args(0), OK_FUNC) Then ThrowErr "TypeError", fo.PName & "() expects a function"
        Set o = args(0)
        If c = BI_STATICMETHOD Then
            Set out = CopyFunc(o, 1)
        Else
            Set out = CopyFunc(o, 2)
        End If
    Case BI_PROPERTY
        Set o = NewObj(OK_PROP)
        If ArgOr(args, na, 0, kw, "fget", v) Then SetV o.FnV, v
        If ArgOr(args, na, 1, kw, "fset", v) Then
            If IsObject(v) Then
                Set o.SelfV = v
            Else
                o.SelfV = v
            End If
        End If
        Set out = o
    Case BI_EXIT
        If na > 0 Then
            RaiseExc MakeExcV("SystemExit", args(0))
        Else
            RaiseExc NewExcInst(ExcClass("SystemExit"), NewObj(OK_TUPLE))
        End If
    Case BI_SUPER
        NeedArgs "super", na, 2, 2
        If Not IsKind(args(0), OK_CLASS) Then ThrowErr "TypeError", "super() argument 1 must be a type"
        Set o = NewObj(OK_SUPER)
        Set o.Cls = args(0)
        If IsObject(args(1)) Then
            Set o.SelfV = args(1)
        Else
            o.SelfV = args(1)
        End If
        Set out = o
    Case BI_GLOBALS
        Set out = AttrDict(gGlobals)
    Case BI_DIR
        Set r = NewList()
        If na = 0 Then
            Set o = gGlobals
            AddNames r, o
        ElseIf IsObject(args(0)) Then
            Set o = args(0)
            AddNames r, o
            If o.Kind = OK_INST Then AddClassNames r, o.Cls
            If o.Kind = OK_CLASS Then AddClassNames r, o
        End If
        If Not IsObject(args(0)) Or na = 0 Then
        End If
        If na > 0 Then AddMethodNames r, args(0)
        x = Empty
        SortList r, x, False
        Set out = r
    Case BI_VARS
        If na = 0 Then
            Set out = AttrDict(gGlobals)
        Else
            If Not IsObject(args(0)) Then ThrowErr "TypeError", "vars() argument must have __dict__ attribute"
            Set out = AttrDict(args(0))
        End If
    Case Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End Select
End Sub

Private Sub AddNames(ByVal r As PyObj, ByVal o As PyObj)
    Dim i As Long, k As Variant
    For i = 0 To o.EntryCount - 1
        If o.LiveAt(i) Then
            o.KeyInto i, k
            If VarType(k) = vbString Then
                If Not Contains(ToVar(r), k) Then r.Push k
            End If
        End If
    Next
End Sub

Private Sub AddClassNames(ByVal r As PyObj, ByVal c As PyObj)
    Dim i As Long, b As Variant
    AddNames r, c
    For i = 0 To c.Count - 1
        c.GetInto i, b
        AddClassNames r, b
    Next
End Sub

Private Sub AddMethodNames(ByVal r As PyObj, ByRef v As Variant)
    Dim lst As String, arr() As String, i As Long, x As Variant
    If VarType(v) = vbString Then
        lst = STR_METHODS
    ElseIf IsKind(v, OK_LIST) Then
        lst = LIST_METHODS
    ElseIf IsKind(v, OK_DICT) Then
        lst = DICT_METHODS
    ElseIf IsKind(v, OK_SET) Then
        lst = SET_METHODS
    End If
    If Len(lst) = 0 Then Exit Sub
    arr = Split(Trim$(lst), " ")
    For i = 0 To UBound(arr)
        x = arr(i)
        If Not Contains(ToVar(r), x) Then r.Push x
    Next
End Sub

Private Function CopyFunc(ByVal f As PyObj, ByVal flags As Long) As PyObj
    Dim g As PyObj
    Set g = NewObj(OK_FUNC)
    g.PName = f.PName
    Set g.Node = f.Node
    Set g.Env = f.Env
    Set g.Cls = f.Cls
    Set g.Defaults = f.Defaults
    g.Flags = flags
    Set CopyFunc = g
End Function

Private Function PowMod(ByRef a As Variant, ByRef e As Variant, ByRef m As Variant) As Variant
    Dim r As Variant, bs As Variant, n As Variant, q As Variant, rm As Variant, mm As Variant
    If Not (IsIntType(a) And IsIntType(e) And IsIntType(m)) Then ThrowErr "TypeError", "pow() 3rd argument not allowed unless all arguments are integers"
    mm = IntOf(m)
    If mm = 0 Then ThrowErr "ValueError", "pow() 3rd argument cannot be 0"
    n = IntOf(e)
    If n < 0 Then ThrowErr "ValueError", "negative exponents are not supported with a modulus"
    IntDivMod a, mm, q, bs
    r = 1&
    IntDivMod r, mm, q, r
    Do While n > 0
        IntDivMod n, 2&, q, rm
        If rm = 1 Then
            IntDivMod IntMul(r, bs), mm, q, r
        End If
        n = q
        If n > 0 Then IntDivMod IntMul(bs, bs), mm, q, bs
    Loop
    PowMod = r
End Function

Private Sub MinMax(ByVal fname As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByVal isMax As Boolean, ByRef out As Variant)
    Dim keyFn As Variant, hasKey As Boolean, dflt As Variant, hasDflt As Boolean, it As PyObj, seq As PyObj
    Dim i As Long, best As Variant, bestK As Variant, x As Variant, k As Variant, first As Boolean
    hasKey = KwGet(kw, "key", keyFn)
    If hasKey Then
        If IsEmpty(keyFn) Then hasKey = False
    End If
    hasDflt = KwGet(kw, "default", dflt)
    If na = 0 Then ThrowErr "TypeError", fname & " expected at least 1 argument, got 0"
    If na = 1 Then
        Set it = GetIter(args(0))
    Else
        Set seq = NewList()
        For i = 0 To na - 1
            seq.Push args(i)
        Next
        Set it = GetIter(ToVar(seq))
    End If
    first = True
    Do While IterNext(it, x)
        If hasKey Then
            CallValue1 keyFn, x, k
        Else
            SetV k, x
        End If
        If first Then
            SetV best, x
            SetV bestK, k
            first = False
        ElseIf isMax Then
            If PyLess(bestK, k) Then
                SetV best, x
                SetV bestK, k
            End If
        Else
            If PyLess(k, bestK) Then
                SetV best, x
                SetV bestK, k
            End If
        End If
    Loop
    If first Then
        If hasDflt Then
            SetV out, dflt
            Exit Sub
        End If
        ThrowErr "ValueError", fname & "() arg is an empty sequence"
    End If
    SetV out, best
End Sub

Private Function IsInstanceOf(ByRef x As Variant, ByRef t As Variant) As Boolean
    Dim o As PyObj, i As Long, y As Variant, xo As PyObj
    If Not IsObject(t) Then ThrowErr "TypeError", "isinstance() arg 2 must be a type or tuple of types"
    Set o = t
    Select Case o.Kind
    Case OK_TUPLE
        For i = 0 To o.Count - 1
            o.GetInto i, y
            If IsInstanceOf(x, y) Then
                IsInstanceOf = True
                Exit Function
            End If
        Next
    Case OK_CLASS
        If o Is gObjectClass Then
            IsInstanceOf = True
        ElseIf IsKind(x, OK_INST) Then
            Set xo = x
            IsInstanceOf = IsSubclassOf(xo.Cls, o)
        End If
    Case OK_BUILTIN
        Select Case o.Code
        Case BI_INT: IsInstanceOf = IsIntType(x)
        Case BI_FLOAT: IsInstanceOf = IsFloatType(x)
        Case BI_STR: IsInstanceOf = (VarType(x) = vbString)
        Case BI_BOOL: IsInstanceOf = (VarType(x) = vbBoolean)
        Case BI_LIST: IsInstanceOf = IsKind(x, OK_LIST)
        Case BI_DICT: IsInstanceOf = IsKind(x, OK_DICT)
        Case BI_TUPLE: IsInstanceOf = IsKind(x, OK_TUPLE)
        Case BI_SET, BI_FROZENSET: IsInstanceOf = IsKind(x, OK_SET)
        Case BI_RANGE: IsInstanceOf = IsKind(x, OK_RANGE)
        Case BI_TYPE
            If IsKind(x, OK_CLASS) Then
                IsInstanceOf = True
            ElseIf IsKind(x, OK_BUILTIN) Then
                Set xo = x
                IsInstanceOf = (xo.Flags = 1)
            End If
        End Select
    Case Else
        ThrowErr "TypeError", "isinstance() arg 2 must be a type or tuple of types"
    End Select
End Function

Private Function IsSubclassV(ByRef a As Variant, ByRef t As Variant) As Boolean
    Dim o As PyObj, i As Long, y As Variant, ao As PyObj
    If Not IsKind(a, OK_CLASS) Then
        If IsKind(a, OK_BUILTIN) And IsKind(t, OK_BUILTIN) Then
            Set ao = a
            Set o = t
            IsSubclassV = (ao.Code = o.Code) Or (ao.Code = BI_BOOL And o.Code = BI_INT)
            Exit Function
        End If
        ThrowErr "TypeError", "issubclass() arg 1 must be a class"
    End If
    Set ao = a
    If IsKind(t, OK_TUPLE) Then
        Set o = t
        For i = 0 To o.Count - 1
            o.GetInto i, y
            If IsSubclassV(a, y) Then
                IsSubclassV = True
                Exit Function
            End If
        Next
        Exit Function
    End If
    If Not IsKind(t, OK_CLASS) Then Exit Function
    Set o = t
    IsSubclassV = IsSubclassOf(ao, o)
End Function

Private Sub TypeOfValue(ByRef x As Variant, ByRef out As Variant)
    Dim o As PyObj, nm As String, t As PyObj
    If IsObject(x) Then
        Set o = x
        If o.Kind = OK_INST Then
            Set out = o.Cls
            Exit Sub
        End If
    End If
    nm = TypeNameOf(x)
    If NameGet(gBuiltins, nm, out) Then
        If IsKind(out, OK_BUILTIN) Then Exit Sub
    End If
    Set t = NewObj(OK_BUILTIN)
    t.PName = nm
    t.Flags = 1
    t.Code = -1
    Set out = t
End Sub

Private Sub DictUpdateFrom(ByVal d As PyObj, ByRef src As Variant)
    Dim it As PyObj, x As Variant, pr As PyObj, k As Variant, v As Variant
    If IsKind(src, OK_DICT) Then
        DictUpdate d, src
        Exit Sub
    End If
    Set it = GetIter(src)
    Do While IterNext(it, x)
        Set pr = SeqOf(x)
        If pr.Count <> 2 Then ThrowErr "ValueError", "dictionary update sequence element has length " & pr.Count & "; 2 is required"
        pr.GetInto 0, k
        pr.GetInto 1, v
        DictSet d, k, v
    Loop
End Sub

Private Sub CallBuiltinB(ByVal fo As PyObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim e As PyObj, t As PyObj, v As Variant, x As Variant, parts() As String, i As Long, cls As PyObj
    Select Case c
    Case BI_EXC_INIT
        If na < 1 Then ThrowErr "TypeError", "__init__ needs self"
        Set e = args(0)
        Set v = MakeTuple(args, 1, na)
        NameSet e, "args", v
        out = Empty
    Case BI_EXC_STR
        Set e = args(0)
        out = ExcMessage(e)
    Case BI_EXC_REPR
        Set e = args(0)
        If NameGet(e, "args", v) Then
            If IsKind(v, OK_TUPLE) Then
                Set t = v
                If t.Count = 0 Then
                    out = e.Cls.PName & "()"
                Else
                    ReDim parts(0 To t.Count - 1)
                    For i = 0 To t.Count - 1
                        t.GetInto i, x
                        parts(i) = PyRepr(x)
                    Next
                    out = e.Cls.PName & "(" & Join(parts, ", ") & ")"
                End If
                Exit Sub
            End If
        End If
        out = e.Cls.PName & "()"
    Case BI_OBJ_INIT
        out = Empty
    Case Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End Select
End Sub

' ---------------------------------------------------------------------------
'  math module
' ---------------------------------------------------------------------------
Private Function Floor1(ByVal d As Double) As Variant
    Floor1 = IntFromDouble(Int(d))
End Function

Private Sub CallMath(ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim d As Double, d2 As Double, x As Variant, y As Variant, i As Long, n As Variant, k As Variant, r As Variant
    Dim q As Variant, rm As Variant, it As PyObj, v As Variant
    Select Case c
    Case BI_M_SQRT
        NeedArgs "sqrt", na, 1, 1
        d = NumArg(args(0), "sqrt")
        If d < 0 Then ThrowErr "ValueError", "math domain error"
        out = Sqr(d)
    Case BI_M_FLOOR
        NeedArgs "floor", na, 1, 1
        If IsIntType(args(0)) Then
            out = IntOf(args(0))
        Else
            out = Floor1(NumArg(args(0), "floor"))
        End If
    Case BI_M_CEIL
        NeedArgs "ceil", na, 1, 1
        If IsIntType(args(0)) Then
            out = IntOf(args(0))
        Else
            out = IntFromDouble(-Int(-NumArg(args(0), "ceil")))
        End If
    Case BI_M_TRUNC
        NeedArgs "trunc", na, 1, 1
        If IsIntType(args(0)) Then
            out = IntOf(args(0))
        Else
            out = IntFromDouble(Fix(NumArg(args(0), "trunc")))
        End If
    Case BI_M_FABS
        NeedArgs "fabs", na, 1, 1
        out = Abs(NumArg(args(0), "fabs"))
    Case BI_M_FACTORIAL
        NeedArgs "factorial", na, 1, 1
        If Not IsIntType(args(0)) Then ThrowErr "TypeError", "factorial() only accepts integral values"
        n = IntOf(args(0))
        If n < 0 Then ThrowErr "ValueError", "factorial() not defined for negative values"
        r = 1&
        For i = 2 To n
            r = IntMul(r, i)
        Next
        out = r
    Case BI_M_GCD, BI_M_LCM
        If c = BI_M_GCD Then
            r = 0&
        Else
            r = 1&
        End If
        For i = 0 To na - 1
            If Not IsIntType(args(i)) Then ThrowErr "TypeError", "'" & TypeNameOf(args(i)) & "' object cannot be interpreted as an integer"
            x = IntOf(args(i))
            If x < 0 Then x = IntNeg(x)
            If c = BI_M_GCD Then
                r = GcdV(r, x)
            Else
                If x = 0 Or r = 0 Then
                    r = 0&
                Else
                    IntDivMod IntMul(r, x), GcdV(r, x), q, rm
                    r = q
                End If
            End If
        Next
        out = r
    Case BI_M_POW
        NeedArgs "pow", na, 2, 2
        d = NumArg(args(0), "pow")
        d2 = NumArg(args(1), "pow")
        If d < 0 And d2 <> Int(d2) Then ThrowErr "ValueError", "math domain error"
        If d = 0 And d2 < 0 Then ThrowErr "ValueError", "math domain error"
        out = d ^ d2
    Case BI_M_EXP
        NeedArgs "exp", na, 1, 1
        out = Exp(NumArg(args(0), "exp"))
    Case BI_M_LOG
        NeedArgs "log", na, 1, 2
        d = NumArg(args(0), "log")
        If d <= 0 Then ThrowErr "ValueError", "math domain error"
        If na = 2 Then
            d2 = NumArg(args(1), "log")
            If d2 <= 0 Or d2 = 1 Then ThrowErr "ValueError", "math domain error"
            out = Log(d) / Log(d2)
        Else
            out = Log(d)
        End If
    Case BI_M_LOG10
        NeedArgs "log10", na, 1, 1
        d = NumArg(args(0), "log10")
        If d <= 0 Then ThrowErr "ValueError", "math domain error"
        out = Log(d) / Log(10#)
        If Abs(out - Round(out)) < 0.000000000001 Then
            If 10 ^ Round(out) = d Then out = CDbl(Round(out))
        End If
    Case BI_M_LOG2
        NeedArgs "log2", na, 1, 1
        d = NumArg(args(0), "log2")
        If d <= 0 Then ThrowErr "ValueError", "math domain error"
        out = Log(d) / Log(2#)
        If Abs(out - Round(out)) < 0.000000000001 Then
            If 2 ^ Round(out) = d Then out = CDbl(Round(out))
        End If
    Case BI_M_SIN
        out = Sin(NumArg(args(0), "sin"))
    Case BI_M_COS
        out = Cos(NumArg(args(0), "cos"))
    Case BI_M_TAN
        out = Tan(NumArg(args(0), "tan"))
    Case BI_M_ASIN, BI_M_ACOS
        d = NumArg(args(0), "asin")
        If d < -1 Or d > 1 Then ThrowErr "ValueError", "math domain error"
        If Abs(d) = 1 Then
            d2 = Sgn(d) * 2 * Atn(1#)
        Else
            d2 = Atn(d / Sqr(1 - d * d))
        End If
        If c = BI_M_ASIN Then
            out = d2
        Else
            out = 2 * Atn(1#) - d2
        End If
    Case BI_M_ATAN
        out = Atn(NumArg(args(0), "atan"))
    Case BI_M_ATAN2
        NeedArgs "atan2", na, 2, 2
        d = NumArg(args(0), "atan2")
        d2 = NumArg(args(1), "atan2")
        out = Atan2(d, d2)
    Case BI_M_DEGREES
        out = NumArg(args(0), "degrees") * 45# / Atn(1#)
    Case BI_M_RADIANS
        out = NumArg(args(0), "radians") * Atn(1#) / 45#
    Case BI_M_HYPOT
        d = 0
        For i = 0 To na - 1
            d2 = NumArg(args(i), "hypot")
            d = d + d2 * d2
        Next
        out = Sqr(d)
    Case BI_M_ISQRT
        NeedArgs "isqrt", na, 1, 1
        If Not IsIntType(args(0)) Then ThrowErr "TypeError", "isqrt() argument must be an integer"
        n = IntOf(args(0))
        If n < 0 Then ThrowErr "ValueError", "isqrt() argument must be nonnegative"
        r = IntFromDouble(Int(Sqr(CDbl(n))))
        Do While IntMul(r, r) > n
            r = IntSub(r, 1&)
        Loop
        Do While IntMul(IntAdd(r, 1&), IntAdd(r, 1&)) <= n
            r = IntAdd(r, 1&)
        Loop
        out = r
    Case BI_M_ISCLOSE
        NeedArgs "isclose", na, 2, 2
        d = NumArg(args(0), "isclose")
        d2 = NumArg(args(1), "isclose")
        If KwGet(kw, "rel_tol", v) Then x = NumToDbl(v) Else x = 0.000000001
        If KwGet(kw, "abs_tol", v) Then y = NumToDbl(v) Else y = 0#
        If Abs(d - d2) <= Abs(x * d) Or Abs(d - d2) <= Abs(x * d2) Or Abs(d - d2) <= y Then
            out = True
        Else
            out = False
        End If
    Case BI_M_COMB, BI_M_PERM
        If na = 1 And c = BI_M_PERM Then
            n = IntOf(args(0))
            k = n
        Else
            NeedArgs "comb", na, 2, 2
            n = IntOf(args(0))
            k = IntOf(args(1))
        End If
        If n < 0 Or k < 0 Then ThrowErr "ValueError", "n and k must be non-negative integers"
        If k > n Then
            out = 0&
            Exit Sub
        End If
        r = 1&
        If c = BI_M_COMB Then
            If k > n - k Then k = n - k
            For i = 1 To k
                IntDivMod IntMul(r, IntSub(n, k - i)), i, q, rm
                r = q
            Next
        Else
            For i = 0 To k - 1
                r = IntMul(r, IntSub(n, i))
            Next
        End If
        out = r
    Case BI_M_PROD
        NeedArgs "prod", na, 1, 1
        If Not KwGet(kw, "start", x) Then x = 1&
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            BinOp OP_MUL, x, v, y
            SetV x, y
        Loop
        SetV out, x
    Case BI_M_FSUM
        NeedArgs "fsum", na, 1, 1
        d = 0
        d2 = 0
        Set it = GetIter(args(0))
        Do While IterNext(it, v)
            ' Kahan summation
            x = NumToDbl(v) - d2
            y = d + x
            d2 = (y - d) - x
            d = y
        Loop
        out = d
    Case BI_M_COPYSIGN
        d = Abs(NumArg(args(0), "copysign"))
        If NumArg(args(1), "copysign") < 0 Then out = -d Else out = d
    Case BI_M_FMOD
        d = NumArg(args(0), "fmod")
        d2 = NumArg(args(1), "fmod")
        If d2 = 0 Then ThrowErr "ValueError", "math domain error"
        out = d - d2 * Fix(d / d2)
    Case BI_M_ISFINITE
        out = (Abs(NumArg(args(0), "isfinite")) < PY_INF)
    Case BI_M_ISINF
        out = (Abs(NumArg(args(0), "isinf")) >= PY_INF)
    Case BI_M_ISNAN
        out = False
    Case Else
        ThrowErr "SystemError", "unknown math function"
    End Select
End Sub

Private Function GcdV(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim q As Variant, r As Variant
    Do While b <> 0
        IntDivMod a, b, q, r
        a = b
        b = r
    Loop
    GcdV = a
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    Dim pi As Double
    pi = 4 * Atn(1#)
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + pi
        Else
            Atan2 = Atn(y / x) - pi
        End If
    Else
        If y > 0 Then
            Atan2 = pi / 2
        ElseIf y < 0 Then
            Atan2 = -pi / 2
        Else
            Atan2 = 0
        End If
    End If
End Function

' ---------------------------------------------------------------------------
'  random, time, sys, collections, functools, itertools
' ---------------------------------------------------------------------------
Private Function RandDbl() As Double
    If Not mRndInit Then
        Randomize
        mRndInit = True
    End If
    RandDbl = (Int(Rnd() * 65536#) * 65536# + Int(Rnd() * 65536#)) / 4294967296#
End Function

Private Function RandBelow(ByVal n As Double) As Double
    RandBelow = Int(RandDbl() * n)
End Function

Private Sub CallLib(ByVal fo As PyObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim lst As PyObj, i As Long, j As Long, n As Long, x As Variant, y As Variant, v As Variant, r As PyObj
    Dim d As Double, st As Long, en As Long, stp As Long, it As PyObj, t0 As Double, k As Long, o As PyObj
    Dim ks As String, e As Long, a2() As Variant, kd As PyObj, idx() As Long
    Select Case c
    Case BI_R_RANDOM
        out = RandDbl()
    Case BI_R_UNIFORM
        NeedArgs "uniform", na, 2, 2
        d = NumArg(args(0), "uniform")
        out = d + (NumArg(args(1), "uniform") - d) * RandDbl()
    Case BI_R_RANDINT
        NeedArgs "randint", na, 2, 2
        st = IntToLong(args(0), "randint")
        en = IntToLong(args(1), "randint")
        If en < st Then ThrowErr "ValueError", "empty range for randint()"
        out = CLng(st + RandBelow(CDbl(en) - st + 1))
    Case BI_R_RANDRANGE
        NeedArgs "randrange", na, 1, 3
        If na = 1 Then
            st = 0
            en = IntToLong(args(0), "randrange")
        Else
            st = IntToLong(args(0), "randrange")
            en = IntToLong(args(1), "randrange")
        End If
        stp = 1
        If na = 3 Then stp = IntToLong(args(2), "randrange")
        If stp = 0 Then ThrowErr "ValueError", "zero step for randrange()"
        n = RangeLen(NewRange(st, en, stp))
        If n <= 0 Then ThrowErr "ValueError", "empty range for randrange()"
        out = CLng(st + stp * RandBelow(n))
    Case BI_R_CHOICE
        NeedArgs "choice", na, 1, 1
        Set lst = SeqOf(args(0))
        If lst.Count = 0 Then ThrowErr "IndexError", "Cannot choose from an empty sequence"
        lst.GetInto CLng(RandBelow(lst.Count)), out
    Case BI_R_SHUFFLE
        NeedArgs "shuffle", na, 1, 1
        If Not IsKind(args(0), OK_LIST) Then ThrowErr "TypeError", "shuffle() argument must be a list"
        Set lst = args(0)
        For i = lst.Count - 1 To 1 Step -1
            j = CLng(RandBelow(i + 1))
            lst.GetInto i, x
            lst.GetInto j, y
            lst.SetI i, y
            lst.SetI j, x
        Next
        out = Empty
    Case BI_R_SAMPLE
        NeedArgs "sample", na, 2, 2
        Set lst = NewListFrom(args(0))
        k = IntToLong(args(1), "sample")
        If k < 0 Or k > lst.Count Then ThrowErr "ValueError", "Sample larger than population or is negative"
        Set r = NewList()
        For i = 0 To k - 1
            j = i + CLng(RandBelow(lst.Count - i))
            lst.GetInto i, x
            lst.GetInto j, y
            lst.SetI i, y
            lst.SetI j, x
            r.Push y
        Next
        Set out = r
    Case BI_R_SEED
        Rnd -1
        If na > 0 Then
            If IsNum(args(0)) Then
                Randomize NumToDbl(args(0))
            Else
                Randomize HashStr(PyStr(args(0)))
            End If
        Else
            Randomize
        End If
        mRndInit = True
        out = Empty
    Case BI_T_TIME
        out = (CDbl(Date) - 25569#) * 86400# + CDbl(Timer)
    Case BI_T_PERF
        out = CDbl(Timer)
    Case BI_T_SLEEP
        NeedArgs "sleep", na, 1, 1
        d = NumArg(args(0), "sleep")
        t0 = Timer
        Do While Timer - t0 < d And Timer >= t0
            DoEvents
        Loop
        out = Empty
    Case BI_S_SETRECLIMIT
        NeedArgs "setrecursionlimit", na, 1, 1
        n = IntToLong(args(0), "limit")
        If n < 10 Then ThrowErr "ValueError", "recursion limit must be at least 10"
        If n > 5000 Then n = 5000
        If n > gMaxDepth Then
            ReDim Preserve gFrameName(0 To n + 2)
            ReDim Preserve gFrameLine(0 To n + 2)
        End If
        gMaxDepth = n
        out = Empty
    Case BI_S_GETRECLIMIT
        out = gMaxDepth
    Case BI_C_DEFAULTDICT
        Set r = NewDict()
        r.Flags = 1
        If na > 0 Then SetV r.FnV, args(0)
        If na > 1 Then DictUpdateFrom r, args(1)
        Set out = r
    Case BI_C_COUNTER
        Set r = NewDict()
        r.Flags = 2
        If na > 0 Then CounterAdd r, args(0), 1
        If Not (kw Is Nothing) Then CounterAdd r, ToVar(kw), 1
        Set out = r
    Case BI_C_ORDEREDDICT
        Set r = NewDict()
        If na > 0 Then DictUpdateFrom r, args(0)
        Set out = r
    Case BI_C_DEQUE
        If na > 0 Then
            Set r = NewListFrom(args(0))
        Else
            Set r = NewList()
        End If
        Set out = r
    Case BI_F_REDUCE
        NeedArgs "reduce", na, 2, 3
        Set it = GetIter(args(1))
        If na = 3 Then
            SetV x, args(2)
        ElseIf Not IterNext(it, x) Then
            ThrowErr "TypeError", "reduce() of empty iterable with no initial value"
        End If
        Do While IterNext(it, v)
            CallValue2 args(0), x, v, y
            SetV x, y
        Loop
        SetV out, x
    Case BI_F_LRU
        ' lru_cache(fn) / lru_cache(maxsize=...) / cache(fn)
        If na = 1 Then
            If IsKind(args(0), OK_FUNC) Or IsKind(args(0), OK_BOUND) Then
                Set out = MakeCached(args(0))
                Exit Sub
            End If
        End If
        Set o = NewObj(OK_BUILTIN)
        o.PName = "lru_cache_decorator"
        o.Code = BI_F_LRU_DECO
        Set out = o
    Case BI_F_LRU_DECO
        NeedArgs "lru_cache", na, 1, 1
        Set out = MakeCached(args(0))
    Case BI_F_CACHED
        ' memoized call
        Set kd = fo.Ref
        Set r = MakeTuple(args, 0, na)
        If Not (kw Is Nothing) Then r.Push ToVar(CopyDict(kw, OK_TUPLE))
        ks = KeyStr(ToVar(r))
        e = kd.FindH(ks, HashStr(ks))
        If e >= 0 Then
            kd.ValInto e, out
        Else
            SetV y, fo.FnV
            CallValue y, args, na, kw, out
            kd.SetH ks, HashStr(ks), ks, out
        End If
    Case BI_I_PERMUTATIONS, BI_I_COMBINATIONS, BI_I_COMBWR
        Set lst = NewListFrom(args(0))
        n = lst.Count
        If na > 1 Then
            If IsEmpty(args(1)) Then
                k = n
            Else
                k = IntToLong(args(1), "r")
            End If
        ElseIf KwGet(kw, "r", v) Then
            k = IntToLong(v, "r")
        ElseIf c = BI_I_PERMUTATIONS Then
            k = n
        Else
            ThrowErr "TypeError", "missing required argument 'r'"
        End If
        Set r = NewList()
        If k >= 0 Then
            If k <= n Or c = BI_I_COMBWR Then
                ReDim idx(0 To IIf(k > 0, k - 1, 0))
                CombRec c, lst, k, 0, 0, idx, r
            End If
        End If
        Set out = r
    Case BI_I_PRODUCT
        n = 1
        If KwGet(kw, "repeat", v) Then n = IntToLong(v, "repeat")
        Set r = NewList()
        Set o = NewList()
        For j = 0 To n - 1
            For i = 0 To na - 1
                o.Push ToVar(NewListFrom(args(i)))
            Next
        Next
        ProductRec o, 0, NewObj(OK_TUPLE), r
        Set out = r
    Case BI_I_CHAIN
        Set r = NewList()
        For i = 0 To na - 1
            Set it = GetIter(args(i))
            Do While IterNext(it, v)
                r.Push v
            Loop
        Next
        Set out = r
    Case BI_I_ACCUMULATE
        Set r = NewList()
        Set it = GetIter(args(0))
        If KwGet(kw, "initial", x) Then
            r.Push x
        ElseIf Not IterNext(it, x) Then
            Set out = r
            Exit Sub
        Else
            r.Push x
        End If
        Do While IterNext(it, v)
            If na > 1 Then
                CallValue2 args(1), x, v, y
            Else
                BinOp OP_ADD, x, v, y
            End If
            SetV x, y
            r.Push x
        Loop
        Set out = r
    Case Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End Select
End Sub

Private Function MakeCached(ByRef f As Variant) As PyObj
    Dim o As PyObj
    Set o = NewObj(OK_BUILTIN)
    o.PName = FuncNameOf(f)
    o.Code = BI_F_CACHED
    SetV o.FnV, f
    Set o.Ref = NewObj(OK_ENV)
    Set MakeCached = o
End Function

Private Sub CombRec(ByVal c As Long, ByVal pool As PyObj, ByVal k As Long, ByVal depth As Long, ByVal startAt As Long, ByRef idx() As Long, ByVal r As PyObj)
    Dim i As Long, j As Long, t As PyObj, x As Variant, used As Boolean
    If depth = k Then
        Set t = NewObj(OK_TUPLE)
        For j = 0 To k - 1
            pool.GetInto idx(j), x
            t.Push x
        Next
        r.Push ToVar(t)
        Exit Sub
    End If
    For i = IIf(c = BI_I_PERMUTATIONS, 0, startAt) To pool.Count - 1
        If c = BI_I_PERMUTATIONS Then
            used = False
            For j = 0 To depth - 1
                If idx(j) = i Then
                    used = True
                    Exit For
                End If
            Next
            If Not used Then
                idx(depth) = i
                CombRec c, pool, k, depth + 1, 0, idx, r
            End If
        Else
            idx(depth) = i
            If c = BI_I_COMBWR Then
                CombRec c, pool, k, depth + 1, i, idx, r
            Else
                CombRec c, pool, k, depth + 1, i + 1, idx, r
            End If
        End If
    Next
End Sub

Private Sub ProductRec(ByVal pools As PyObj, ByVal depth As Long, ByVal cur As PyObj, ByVal r As PyObj)
    Dim i As Long, pool As PyObj, x As Variant, t As PyObj, pv As Variant
    If depth >= pools.Count Then
        Set t = NewListFrom(ToVar(cur))
        t.Kind = OK_TUPLE
        r.Push ToVar(t)
        Exit Sub
    End If
    pools.GetInto depth, pv
    Set pool = pv
    For i = 0 To pool.Count - 1
        pool.GetInto i, x
        cur.Push x
        ProductRec pools, depth + 1, cur, r
        cur.RemoveAt cur.Count - 1
    Next
End Sub

Private Sub CounterAdd(ByVal d As PyObj, ByRef src As Variant, ByVal sgnv As Long)
    Dim it As PyObj, x As Variant, cur As Variant, s As PyObj, i As Long, k As Variant, v As Variant
    If IsKind(src, OK_DICT) Then
        Set s = src
        For i = 0 To s.EntryCount - 1
            If s.LiveAt(i) Then
                s.KeyInto i, k
                s.ValInto i, v
                If Not DictGet(d, k, cur) Then cur = 0&
                If sgnv < 0 Then v = IntNeg(v)
                BinOp OP_ADD, cur, v, x
                DictSet d, k, x
            End If
        Next
        Exit Sub
    End If
    Set it = GetIter(src)
    Do While IterNext(it, x)
        If Not DictGet(d, x, cur) Then cur = 0&
        cur = IntAdd(cur, sgnv)
        DictSet d, x, cur
    Loop
End Sub

' ---------------------------------------------------------------------------
'  modules
' ---------------------------------------------------------------------------
Private Sub TypingAlias(ByVal m As PyObj, ByVal aliasName As String, ByVal builtinName As String)
    Dim v As Variant
    If NameGet(gBuiltins, builtinName, v) Then NameSet m, aliasName, v
End Sub

' ---------------------------------------------------------------------------
'  dataclasses
' ---------------------------------------------------------------------------
Private Function DcFields(ByVal cls As PyObj) As PyObj
    Dim v As Variant
    If ClassLookup(cls, "__dataclass_fields__", HashStr("__dataclass_fields__"), v) Then
        Set DcFields = v
    Else
        Set DcFields = NewObj(OK_TUPLE)
    End If
End Function

Private Sub DcInstall(ByVal cls As PyObj, ByVal nm As String, ByVal code As Long)
    Dim o As PyObj
    If cls.FindH(nm, HashStr(nm)) >= 0 Then Exit Sub
    Set o = RegB(cls, nm, code)
    o.Flags = 2
End Sub

Private Sub MakeDataclass(ByVal cls As PyObj, ByVal opts As PyObj)
    Dim flds As PyObj, i As Long, j As Long, b As Variant, bf As PyObj, x As Variant, ann As Variant, a As PyObj
    Dim k As Variant, v As Variant, seen As PyObj
    Set flds = NewObj(OK_TUPLE)
    Set seen = NewObj(OK_SET)
    For i = cls.Count - 1 To 0 Step -1
        cls.GetInto i, b
        Set bf = DcFields(b)
        For j = 0 To bf.Count - 1
            bf.GetInto j, x
            If DictFind(seen, x) < 0 Then
                flds.Push x
                SetAdd seen, x
            End If
        Next
    Next
    If NameGet(cls, "__annotations__", ann) Then
        Set a = ann
        For i = 0 To a.EntryCount - 1
            If a.LiveAt(i) Then
                a.KeyInto i, k
                If DictFind(seen, k) < 0 Then
                    flds.Push k
                    SetAdd seen, k
                End If
            End If
        Next
    End If
    Set v = flds
    NameSet cls, "__dataclass_fields__", v
    DcInstall cls, "__init__", BI_DC_INIT
    DcInstall cls, "__repr__", BI_DC_REPR
    DcInstall cls, "__eq__", BI_DC_EQ
    If KwGet(opts, "order", v) Then
        If Truthy(v) Then
            DcInstall cls, "__lt__", BI_DC_LT
            DcInstall cls, "__le__", BI_DC_LE
            DcInstall cls, "__gt__", BI_DC_GT
            DcInstall cls, "__ge__", BI_DC_GE
        End If
    End If
End Sub

Private Function DcTuple(ByVal inst As PyObj) As PyObj
    Dim flds As PyObj, t As PyObj, i As Long, nm As Variant, v As Variant
    Set flds = DcFields(inst.Cls)
    Set t = NewObj(OK_TUPLE)
    For i = 0 To flds.Count - 1
        flds.GetInto i, nm
        PyGetAttr ToVar(inst), CStr(nm), v
        t.Push v
    Next
    Set DcTuple = t
End Function

Private Sub CallDataclass(ByVal fo As PyObj, ByVal c As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim o As PyObj, inst As PyObj, other As PyObj, flds As PyObj, i As Long, nm As Variant, v As Variant, d As Variant
    Dim fld As PyObj, used As Long, parts() As String, f As Variant, r As Variant
    Select Case c
    Case BI_DC_DATACLASS
        If na = 1 Then
            If IsKind(args(0), OK_CLASS) Then
                MakeDataclass args(0), kw
                SetV out, args(0)
                Exit Sub
            End If
        End If
        Set o = NewObj(OK_BUILTIN)
        o.PName = "dataclass"
        o.Code = BI_DC_DECO
        If kw Is Nothing Then
            Set o.Ref = NewDict()
        Else
            Set o.Ref = kw
        End If
        Set out = o
    Case BI_DC_DECO
        NeedArgs "dataclass", na, 1, 1
        If Not IsKind(args(0), OK_CLASS) Then ThrowErr "TypeError", "dataclass() should be applied to a class"
        MakeDataclass args(0), fo.Ref
        SetV out, args(0)
    Case BI_DC_FIELD
        Set o = NewObj(OK_SENTINEL)
        o.PName = "<dataclass field>"
        o.Flags = 7
        If KwGet(kw, "default", v) Then
            If IsObject(v) Then Set o.SelfV = v Else o.SelfV = v
        Else
            Set o.SelfV = gNoDefault
        End If
        If KwGet(kw, "default_factory", v) Then
            If IsObject(v) Then Set o.FnV = v Else o.FnV = v
        End If
        Set out = o
    Case BI_DC_INIT
        Set inst = args(0)
        Set flds = DcFields(inst.Cls)
        If na - 1 > flds.Count Then ThrowErr "TypeError", "__init__() takes " & (flds.Count + 1) & " positional arguments but " & na & " were given"
        For i = 0 To flds.Count - 1
            flds.GetInto i, nm
            If i + 1 < na Then
                SetV v, args(i + 1)
            ElseIf KwGet(kw, CStr(nm), v) Then
                used = used + 1
            ElseIf ClassLookup(inst.Cls, CStr(nm), HashStr(CStr(nm)), d) Then
                SetV v, d
                If IsKind(d, OK_SENTINEL) Then
                    Set fld = d
                    If fld.Flags = 7 Then
                        If Not IsEmpty(fld.FnV) Then
                            SetV f, fld.FnV
                            CallValue0 f, v
                        Else
                            SetV v, fld.SelfV
                            If IsObject(v) Then
                                If v Is gNoDefault Then ThrowErr "TypeError", "__init__() missing 1 required positional argument: '" & nm & "'"
                            End If
                        End If
                    End If
                End If
            Else
                ThrowErr "TypeError", "__init__() missing 1 required positional argument: '" & nm & "'"
            End If
            inst.SetH CStr(nm), HashStr(CStr(nm)), nm, v
        Next
        If Not (kw Is Nothing) Then
            If kw.DCount > used Then
                For i = 0 To kw.EntryCount - 1
                    If kw.LiveAt(i) Then
                        kw.KeyInto i, nm
                        If Not Contains(ToVar(flds), nm) Then ThrowErr "TypeError", "__init__() got an unexpected keyword argument '" & nm & "'"
                    End If
                Next
            End If
        End If
        If ClassLookup(inst.Cls, "__post_init__", HashStr("__post_init__"), f) Then CallValue1 f, ToVar(inst), r
        out = Empty
    Case BI_DC_REPR
        Set inst = args(0)
        Set flds = DcFields(inst.Cls)
        If flds.Count = 0 Then
            out = inst.Cls.PName & "()"
        Else
            ReDim parts(0 To flds.Count - 1)
            For i = 0 To flds.Count - 1
                flds.GetInto i, nm
                PyGetAttr ToVar(inst), CStr(nm), v
                parts(i) = nm & "=" & PyRepr(v)
            Next
            out = inst.Cls.PName & "(" & Join(parts, ", ") & ")"
        End If
    Case BI_DC_EQ, BI_DC_LT, BI_DC_LE, BI_DC_GT, BI_DC_GE
        NeedArgs fo.PName, na, 2, 2
        Set inst = args(0)
        If Not IsKind(args(1), OK_INST) Then
            Set out = gNotImpl
            Exit Sub
        End If
        Set other = args(1)
        If Not (other.Cls Is inst.Cls) Then
            Set out = gNotImpl
            Exit Sub
        End If
        Select Case c
        Case BI_DC_EQ: out = PyEq(ToVar(DcTuple(inst)), ToVar(DcTuple(other)))
        Case BI_DC_LT: out = OrderCmp(OP_LT, ToVar(DcTuple(inst)), ToVar(DcTuple(other)))
        Case BI_DC_LE: out = OrderCmp(OP_LE, ToVar(DcTuple(inst)), ToVar(DcTuple(other)))
        Case BI_DC_GT: out = OrderCmp(OP_GT, ToVar(DcTuple(inst)), ToVar(DcTuple(other)))
        Case Else: out = OrderCmp(OP_GE, ToVar(DcTuple(inst)), ToVar(DcTuple(other)))
        End Select
    Case BI_DC_ASDICT
        NeedArgs "asdict", na, 1, 1
        If Not IsKind(args(0), OK_INST) Then ThrowErr "TypeError", "asdict() should be called on dataclass instances"
        Set inst = args(0)
        Set flds = DcFields(inst.Cls)
        Set o = NewDict()
        For i = 0 To flds.Count - 1
            flds.GetInto i, nm
            PyGetAttr ToVar(inst), CStr(nm), v
            DictSet o, nm, v
        Next
        Set out = o
    Case BI_TY_TYPEVAR
        If na > 1 And fo.PName = "cast" Then
            SetV out, args(1)
        ElseIf na > 0 And fo.PName = "NewType" Then
            SetV out, args(na - 1)
        Else
            Set out = gObjectClass
        End If
    Case Else
        ThrowErr "SystemError", "unknown builtin " & fo.PName
    End Select
End Sub

Private Function NewModule(ByVal nm As String) As PyObj
    Dim m As PyObj
    Set m = NewObj(OK_MODULE)
    m.PName = nm
    NameSet m, "__name__", nm
    Set NewModule = m
End Function

Public Function BuildModule(ByVal nm As String) As PyObj
    Dim m As PyObj, v As Variant, names() As String, ti As Long
    Select Case nm
    Case "math"
        Set m = NewModule(nm)
        RegB m, "sqrt", BI_M_SQRT
        RegB m, "floor", BI_M_FLOOR
        RegB m, "ceil", BI_M_CEIL
        RegB m, "trunc", BI_M_TRUNC
        RegB m, "fabs", BI_M_FABS
        RegB m, "factorial", BI_M_FACTORIAL
        RegB m, "gcd", BI_M_GCD
        RegB m, "lcm", BI_M_LCM
        RegB m, "pow", BI_M_POW
        RegB m, "exp", BI_M_EXP
        RegB m, "log", BI_M_LOG
        RegB m, "log10", BI_M_LOG10
        RegB m, "log2", BI_M_LOG2
        RegB m, "sin", BI_M_SIN
        RegB m, "cos", BI_M_COS
        RegB m, "tan", BI_M_TAN
        RegB m, "asin", BI_M_ASIN
        RegB m, "acos", BI_M_ACOS
        RegB m, "atan", BI_M_ATAN
        RegB m, "atan2", BI_M_ATAN2
        RegB m, "degrees", BI_M_DEGREES
        RegB m, "radians", BI_M_RADIANS
        RegB m, "hypot", BI_M_HYPOT
        RegB m, "isqrt", BI_M_ISQRT
        RegB m, "isclose", BI_M_ISCLOSE
        RegB m, "comb", BI_M_COMB
        RegB m, "perm", BI_M_PERM
        RegB m, "prod", BI_M_PROD
        RegB m, "fsum", BI_M_FSUM
        RegB m, "copysign", BI_M_COPYSIGN
        RegB m, "fmod", BI_M_FMOD
        RegB m, "isfinite", BI_M_ISFINITE
        RegB m, "isinf", BI_M_ISINF
        RegB m, "isnan", BI_M_ISNAN
        v = 4# * Atn(1#)
        NameSet m, "pi", v
        v = 8# * Atn(1#)
        NameSet m, "tau", v
        v = Exp(1#)
        NameSet m, "e", v
        v = PY_INF
        NameSet m, "inf", v
    Case "random"
        Set m = NewModule(nm)
        RegB m, "random", BI_R_RANDOM
        RegB m, "randint", BI_R_RANDINT
        RegB m, "choice", BI_R_CHOICE
        RegB m, "shuffle", BI_R_SHUFFLE
        RegB m, "uniform", BI_R_UNIFORM
        RegB m, "seed", BI_R_SEED
        RegB m, "randrange", BI_R_RANDRANGE
        RegB m, "sample", BI_R_SAMPLE
    Case "time"
        Set m = NewModule(nm)
        RegB m, "time", BI_T_TIME
        RegB m, "sleep", BI_T_SLEEP
        RegB m, "perf_counter", BI_T_PERF
    Case "sys"
        Set m = NewModule(nm)
        RegB m, "setrecursionlimit", BI_S_SETRECLIMIT
        RegB m, "getrecursionlimit", BI_S_GETRECLIMIT
        RegB m, "exit", BI_EXIT
        v = 2147483647
        NameSet m, "maxsize", v
        v = "3.12 (Python subset interpreter for VBA)"
        NameSet m, "version", v
        v = "vba"
        NameSet m, "platform", v
        Set v = NewList()
        NameSet m, "argv", v
    Case "string"
        Set m = NewModule(nm)
        v = "abcdefghijklmnopqrstuvwxyz"
        NameSet m, "ascii_lowercase", v
        v = UCase$(v)
        NameSet m, "ascii_uppercase", v
        v = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        NameSet m, "ascii_letters", v
        v = "0123456789"
        NameSet m, "digits", v
        v = "0123456789abcdefABCDEF"
        NameSet m, "hexdigits", v
        v = "01234567"
        NameSet m, "octdigits", v
        v = "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
        NameSet m, "punctuation", v
        v = " " & vbTab & vbLf & vbCr & ChrW(11) & ChrW(12)
        NameSet m, "whitespace", v
    Case "collections"
        Set m = NewModule(nm)
        RegB m, "defaultdict", BI_C_DEFAULTDICT
        RegB m, "Counter", BI_C_COUNTER
        RegB m, "OrderedDict", BI_C_ORDEREDDICT
        RegB m, "deque", BI_C_DEQUE
    Case "functools"
        Set m = NewModule(nm)
        RegB m, "reduce", BI_F_REDUCE
        RegB m, "lru_cache", BI_F_LRU
        RegB m, "cache", BI_F_LRU
    Case "itertools"
        Set m = NewModule(nm)
        RegB m, "permutations", BI_I_PERMUTATIONS
        RegB m, "combinations", BI_I_COMBINATIONS
        RegB m, "combinations_with_replacement", BI_I_COMBWR
        RegB m, "product", BI_I_PRODUCT
        RegB m, "chain", BI_I_CHAIN
        RegB m, "accumulate", BI_I_ACCUMULATE
    Case "dataclasses"
        Set m = NewModule(nm)
        RegB m, "dataclass", BI_DC_DATACLASS
        RegB m, "field", BI_DC_FIELD
        RegB m, "asdict", BI_DC_ASDICT
    Case "typing"
        Set m = NewModule(nm)
        TypingAlias m, "List", "list"
        TypingAlias m, "Dict", "dict"
        TypingAlias m, "Set", "set"
        TypingAlias m, "FrozenSet", "frozenset"
        TypingAlias m, "Tuple", "tuple"
        TypingAlias m, "Type", "type"
        names = Split("Any Optional Union Callable Iterable Iterator Sequence Mapping MutableMapping Generator NoReturn Literal Final ClassVar Hashable Sized Collection Deque DefaultDict Counter", " ")
        Set v = gObjectClass
        For ti = 0 To UBound(names)
            NameSet m, names(ti), v
        Next
        RegB m, "TypeVar", BI_TY_TYPEVAR
        RegB m, "NewType", BI_TY_TYPEVAR
        RegB m, "cast", BI_TY_TYPEVAR
    Case "__future__"
        Set m = NewModule(nm)
        v = True
        NameSet m, "annotations", v
        NameSet m, "division", v
        NameSet m, "print_function", v
    Case "excel"
        Set m = NewModule(nm)
        RegB m, "cell", BI_X_CELL
        RegB m, "set_cell", BI_X_SETCELL
        RegB m, "get_range", BI_X_GETRANGE
        RegB m, "set_range", BI_X_SETRANGE
        RegB m, "msgbox", BI_X_MSGBOX
        RegB m, "evaluate", BI_X_EVAL
        RegB m, "sheet_names", BI_X_SHEETNAMES
        RegB m, "clear_range", BI_X_CLEAR
    Case Else
        Set m = Nothing
    End Select
    Set BuildModule = m
End Function

' ---------------------------------------------------------------------------
'  methods of built-in types
' ---------------------------------------------------------------------------
Public Function HasBuiltinMethod(ByRef v As Variant, ByVal nm As String) As Boolean
    Dim lst As String, o As PyObj
    Select Case VarType(v)
    Case vbString
        lst = STR_METHODS
    Case vbObject
        Set o = v
        Select Case o.Kind
        Case OK_LIST: lst = LIST_METHODS
        Case OK_TUPLE: lst = " index count "
        Case OK_DICT: lst = DICT_METHODS
        Case OK_SET: lst = SET_METHODS
        End Select
    Case vbDouble
        lst = " is_integer "
    Case vbLong, vbDecimal, vbInteger, vbBoolean
        lst = " bit_length "
    End Select
    If Len(lst) > 0 Then HasBuiltinMethod = (InStr(1, lst, " " & nm & " ", vbBinaryCompare) > 0)
End Function

Public Function TypeHasMethod(ByVal code As Long, ByVal nm As String) As Boolean
    Dim lst As String
    Select Case code
    Case BI_STR: lst = STR_METHODS
    Case BI_LIST: lst = LIST_METHODS
    Case BI_DICT: lst = DICT_METHODS
    Case BI_SET: lst = SET_METHODS
    Case BI_TUPLE: lst = " index count "
    Case BI_INT: lst = " bit_length "
    Case BI_FLOAT: lst = " is_integer "
    End Select
    If Len(lst) > 0 Then TypeHasMethod = (InStr(1, lst, " " & nm & " ", vbBinaryCompare) > 0)
End Function

Public Sub CallMethod(ByRef selfV As Variant, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim o As PyObj, f As Variant
    Select Case VarType(selfV)
    Case vbString
        StrMethod CStr(selfV), nm, args, na, kw, out
        Exit Sub
    Case vbObject
        Set o = selfV
        Select Case o.Kind
        Case OK_LIST
            ListMethod o, nm, args, na, kw, out
            Exit Sub
        Case OK_TUPLE
            TupleMethod o, nm, args, na, out
            Exit Sub
        Case OK_DICT
            DictMethod o, nm, args, na, kw, out
            Exit Sub
        Case OK_SET
            SetMethod o, nm, args, na, out
            Exit Sub
        Case OK_PROP
            Dim p As PyObj
            Set p = NewObj(OK_PROP)
            SetV p.FnV, o.FnV
            If IsObject(o.SelfV) Then
                Set p.SelfV = o.SelfV
            Else
                p.SelfV = o.SelfV
            End If
            NeedArgs nm, na, 1, 1
            If nm = "setter" Then
                If IsObject(args(0)) Then
                    Set p.SelfV = args(0)
                Else
                    p.SelfV = args(0)
                End If
            ElseIf nm = "getter" Then
                SetV p.FnV, args(0)
            End If
            Set out = p
            Exit Sub
        Case Else
            PyGetAttr selfV, nm, f
            CallValue f, args, na, kw, out
            Exit Sub
        End Select
    Case vbDouble, vbSingle
        If nm = "is_integer" Then
            out = (CDbl(selfV) = Int(CDbl(selfV)))
            Exit Sub
        End If
    Case vbLong, vbInteger, vbDecimal, vbBoolean
        If nm = "bit_length" Then
            Dim x As Variant
            x = IntOf(selfV)
            If x < 0 Then x = IntNeg(x)
            If x = 0 Then
                out = 0&
            Else
                out = Len(IntToBase(x, 2))
            End If
            Exit Sub
        End If
    End Select
    ThrowErr "AttributeError", "'" & TypeNameOf(selfV) & "' object has no attribute '" & nm & "'"
End Sub

Private Function StripChars(ByVal s As String, ByRef args() As Variant, ByVal na As Long, ByVal mode As Long) As String
    Dim chars As String, useWs As Boolean, i As Long, j As Long, n As Long
    useWs = True
    If na > 0 Then
        If Not IsEmpty(args(0)) Then
            chars = ArgStr(args(0), "strip arg")
            useWs = False
        End If
    End If
    n = Len(s)
    i = 1
    j = n
    If (mode And 1) <> 0 Then
        Do While i <= n
            If useWs Then
                If Not IsWs(Mid$(s, i, 1)) Then Exit Do
            Else
                If InStr(1, chars, Mid$(s, i, 1), vbBinaryCompare) = 0 Then Exit Do
            End If
            i = i + 1
        Loop
    End If
    If (mode And 2) <> 0 Then
        Do While j >= i
            If useWs Then
                If Not IsWs(Mid$(s, j, 1)) Then Exit Do
            Else
                If InStr(1, chars, Mid$(s, j, 1), vbBinaryCompare) = 0 Then Exit Do
            End If
            j = j - 1
        Loop
    End If
    If j < i Then
        StripChars = vbNullString
    Else
        StripChars = Mid$(s, i, j - i + 1)
    End If
End Function

Private Function SplitStr(ByVal s As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByVal fromRight As Boolean) As PyObj
    Dim r As PyObj, sepV As Variant, sep As String, maxs As Long, v As Variant, i As Long, j As Long, n As Long
    Dim cnt As Long, p As Long, x As Variant, ls As Long, pieces As PyObj, arr() As String
    Set r = NewList()
    maxs = -1
    If Not ArgOr(args, na, 0, kw, "sep", sepV) Then sepV = Empty
    If ArgOr(args, na, 1, kw, "maxsplit", v) Then maxs = IntToLong(v, "maxsplit")
    n = Len(s)
    If IsEmpty(sepV) Then
        If Not fromRight Then
            i = 1
            Do
                Do While i <= n
                    If Not IsWs(Mid$(s, i, 1)) Then Exit Do
                    i = i + 1
                Loop
                If i > n Then Exit Do
                If maxs >= 0 And cnt >= maxs Then
                    x = StripChars(Mid$(s, i), args, 0, 2)
                    r.Push x
                    Exit Do
                End If
                j = i
                Do While j <= n
                    If IsWs(Mid$(s, j, 1)) Then Exit Do
                    j = j + 1
                Loop
                x = Mid$(s, i, j - i)
                r.Push x
                cnt = cnt + 1
                i = j
            Loop
        Else
            j = n
            Do
                Do While j >= 1
                    If Not IsWs(Mid$(s, j, 1)) Then Exit Do
                    j = j - 1
                Loop
                If j < 1 Then Exit Do
                If maxs >= 0 And cnt >= maxs Then
                    x = StripChars(Left$(s, j), args, 0, 1)
                    r.InsertAt 0, x
                    Exit Do
                End If
                i = j
                Do While i >= 1
                    If IsWs(Mid$(s, i, 1)) Then Exit Do
                    i = i - 1
                Loop
                x = Mid$(s, i + 1, j - i)
                r.InsertAt 0, x
                cnt = cnt + 1
                j = i
            Loop
        End If
        Set SplitStr = r
        Exit Function
    End If
    sep = ArgStr(sepV, "sep")
    If Len(sep) = 0 Then ThrowErr "ValueError", "empty separator"
    ls = Len(sep)
    If Not fromRight Then
        i = 1
        Do
            If maxs >= 0 And cnt >= maxs Then Exit Do
            p = InStr(i, s, sep, vbBinaryCompare)
            If p = 0 Then Exit Do
            x = Mid$(s, i, p - i)
            r.Push x
            cnt = cnt + 1
            i = p + ls
        Loop
        x = Mid$(s, i)
        r.Push x
    Else
        j = n
        Do
            If maxs >= 0 And cnt >= maxs Then Exit Do
            If j < ls Then Exit Do
            p = InStrRev(s, sep, j, vbBinaryCompare)
            If p = 0 Then Exit Do
            x = Mid$(s, p + ls, j - p - ls + 1)
            r.InsertAt 0, x
            cnt = cnt + 1
            j = p - 1
        Loop
        x = Left$(s, j)
        r.InsertAt 0, x
    End If
    Set SplitStr = r
End Function

Private Sub SubRange(ByVal n As Long, ByRef args() As Variant, ByVal na As Long, ByVal first As Long, ByRef st As Long, ByRef en As Long)
    Dim vs As Variant, ve As Variant, vst As Variant, stp As Long, cnt As Long
    If na > first Then SetV vs, args(first)
    If na > first + 1 Then SetV ve, args(first + 1)
    SliceIndices n, vs, ve, vst, st, stp, cnt
    If cnt < 0 Then cnt = 0
    en = st + cnt
End Sub

Private Function FindIn(ByVal s As String, ByVal subs As String, ByRef args() As Variant, ByVal na As Long, ByVal rev As Boolean) As Long
    Dim st As Long, en As Long, p As Long, seg As String
    SubRange Len(s), args, na, 1, st, en
    If en - st < Len(subs) Then
        FindIn = -1
        Exit Function
    End If
    seg = Mid$(s, st + 1, en - st)
    If Len(subs) = 0 Then
        If rev Then
            FindIn = en
        Else
            FindIn = st
        End If
        Exit Function
    End If
    If rev Then
        p = InStrRev(seg, subs, -1, vbBinaryCompare)
    Else
        p = InStr(1, seg, subs, vbBinaryCompare)
    End If
    If p = 0 Then
        FindIn = -1
    Else
        FindIn = st + p - 1
    End If
End Function

Private Function CountIn(ByVal s As String, ByVal subs As String) As Long
    Dim p As Long, n As Long
    If Len(subs) = 0 Then
        CountIn = Len(s) + 1
        Exit Function
    End If
    p = InStr(1, s, subs, vbBinaryCompare)
    Do While p > 0
        n = n + 1
        p = InStr(p + Len(subs), s, subs, vbBinaryCompare)
    Loop
    CountIn = n
End Function

Private Function AllChars(ByVal s As String, ByVal kind As Long) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case kind
        Case 1
            If Not IsDigitCh(c) Then Exit Function
        Case 2
            If Not IsAlphaC(c) Then Exit Function
        Case 3
            If Not (IsAlphaC(c) Or IsDigitCh(c)) Then Exit Function
        Case 4
            If Not IsWs(c) Then Exit Function
        End Select
    Next
    AllChars = True
End Function

Private Function TitleCase(ByVal s As String) As String
    Dim i As Long, c As String, prev As Boolean, parts() As String
    If Len(s) = 0 Then Exit Function
    ReDim parts(0 To Len(s) - 1)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If UCase$(c) <> LCase$(c) Then
            If prev Then
                c = LCase$(c)
            Else
                c = UCase$(c)
            End If
            prev = True
        Else
            prev = False
        End If
        parts(i - 1) = c
    Next
    TitleCase = Join(parts, vbNullString)
End Function

Private Function PrefixMatch(ByVal s As String, ByRef pv As Variant, ByVal atEnd As Boolean) As Boolean
    Dim t As PyObj, i As Long, x As Variant, pre As String
    If IsKind(pv, OK_TUPLE) Then
        Set t = pv
        For i = 0 To t.Count - 1
            t.GetInto i, x
            If PrefixMatch(s, x, atEnd) Then
                PrefixMatch = True
                Exit Function
            End If
        Next
        Exit Function
    End If
    pre = ArgStr(pv, IIf(atEnd, "endswith", "startswith") & " arg")
    If Len(pre) > Len(s) Then Exit Function
    If atEnd Then
        PrefixMatch = (Right$(s, Len(pre)) = pre)
    Else
        PrefixMatch = (Left$(s, Len(pre)) = pre)
    End If
End Function

Private Sub StrMethod(ByVal s As String, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim r As PyObj, i As Long, n As Long, t As String, x As Variant, k As Long, cnt As Long
    Dim parts() As String, lst As PyObj, wid As Long, fill As String, marg As Long, lft As Long, st As Long, en As Long
    Dim y As Variant, p As Long, sub_ As String
    Select Case nm
    Case "upper"
        out = UCase$(s)
    Case "lower", "casefold"
        out = LCase$(s)
    Case "strip"
        out = StripChars(s, args, na, 3)
    Case "lstrip"
        out = StripChars(s, args, na, 1)
    Case "rstrip"
        out = StripChars(s, args, na, 2)
    Case "split"
        Set out = SplitStr(s, args, na, kw, False)
    Case "rsplit"
        Set out = SplitStr(s, args, na, kw, True)
    Case "splitlines"
        Set r = NewList()
        t = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
        If Len(t) > 0 Then
            If Right$(t, 1) = vbLf Then t = Left$(t, Len(t) - 1)
            parts = Split(t, vbLf)
            For i = 0 To UBound(parts)
                x = parts(i)
                r.Push x
            Next
        End If
        Set out = r
    Case "join"
        NeedArgs "join", na, 1, 1
        Set lst = SeqOf(args(0))
        If lst.Count = 0 Then
            out = vbNullString
        Else
            ReDim parts(0 To lst.Count - 1)
            For i = 0 To lst.Count - 1
                lst.GetInto i, x
                If VarType(x) <> vbString Then ThrowErr "TypeError", "sequence item " & i & ": expected str instance, " & TypeNameOf(x) & " found"
                parts(i) = x
            Next
            out = Join(parts, s)
        End If
    Case "replace"
        NeedArgs "replace", na, 2, 3
        cnt = -1
        If na = 3 Then cnt = IntToLong(args(2), "count")
        t = ArgStr(args(0), "replace() argument 1")
        sub_ = ArgStr(args(1), "replace() argument 2")
        If Len(t) = 0 Then
            n = Len(s)
            ReDim parts(0 To n)
            For i = 0 To n
                If cnt < 0 Or i < cnt Then parts(i) = sub_
                If i < n Then parts(i) = parts(i) & Mid$(s, i + 1, 1)
            Next
            out = Join(parts, vbNullString)
        ElseIf cnt = 0 Then
            out = s
        Else
            out = Replace(s, t, sub_, 1, cnt, vbBinaryCompare)
        End If
    Case "find", "index"
        NeedArgs nm, na, 1, 3
        k = FindIn(s, ArgStr(args(0), "substring"), args, na, False)
        If k < 0 And nm = "index" Then ThrowErr "ValueError", "substring not found"
        out = k
    Case "rfind", "rindex"
        NeedArgs nm, na, 1, 3
        k = FindIn(s, ArgStr(args(0), "substring"), args, na, True)
        If k < 0 And nm = "rindex" Then ThrowErr "ValueError", "substring not found"
        out = k
    Case "count"
        NeedArgs "count", na, 1, 3
        SubRange Len(s), args, na, 1, st, en
        out = CountIn(Mid$(s, st + 1, en - st), ArgStr(args(0), "substring"))
    Case "startswith", "endswith"
        NeedArgs nm, na, 1, 3
        SubRange Len(s), args, na, 1, st, en
        out = PrefixMatch(Mid$(s, st + 1, en - st), args(0), nm = "endswith")
    Case "isdigit", "isnumeric", "isdecimal"
        out = AllChars(s, 1)
    Case "isalpha"
        out = AllChars(s, 2)
    Case "isalnum"
        out = AllChars(s, 3)
    Case "isspace"
        out = AllChars(s, 4)
    Case "isupper"
        out = (UCase$(s) = s And LCase$(s) <> s)
    Case "islower"
        out = (LCase$(s) = s And UCase$(s) <> s)
    Case "istitle"
        out = (TitleCase(s) = s And UCase$(s) <> LCase$(s))
    Case "title"
        out = TitleCase(s)
    Case "capitalize"
        If Len(s) = 0 Then
            out = s
        Else
            out = UCase$(Left$(s, 1)) & LCase$(Mid$(s, 2))
        End If
    Case "swapcase"
        n = Len(s)
        If n = 0 Then
            out = s
        Else
            ReDim parts(0 To n - 1)
            For i = 1 To n
                t = Mid$(s, i, 1)
                If UCase$(t) = t Then
                    parts(i - 1) = LCase$(t)
                Else
                    parts(i - 1) = UCase$(t)
                End If
            Next
            out = Join(parts, vbNullString)
        End If
    Case "center", "ljust", "rjust"
        NeedArgs nm, na, 1, 2
        wid = IntToLong(args(0), "width")
        fill = " "
        If na = 2 Then
            fill = ArgStr(args(1), "fillchar")
            If Len(fill) <> 1 Then ThrowErr "TypeError", "The fill character must be exactly one character long"
        End If
        marg = wid - Len(s)
        If marg <= 0 Then
            out = s
        ElseIf nm = "ljust" Then
            out = s & String$(marg, fill)
        ElseIf nm = "rjust" Then
            out = String$(marg, fill) & s
        Else
            lft = marg \ 2 + (marg And wid And 1)
            out = String$(lft, fill) & s & String$(marg - lft, fill)
        End If
    Case "zfill"
        NeedArgs "zfill", na, 1, 1
        wid = IntToLong(args(0), "width")
        If Len(s) >= wid Then
            out = s
        ElseIf Left$(s, 1) = "-" Or Left$(s, 1) = "+" Then
            out = Left$(s, 1) & String$(wid - Len(s), "0") & Mid$(s, 2)
        Else
            out = String$(wid - Len(s), "0") & s
        End If
    Case "format"
        out = StrFormatMethod(s, args, na, kw)
    Case "partition", "rpartition"
        NeedArgs nm, na, 1, 1
        t = ArgStr(args(0), "separator")
        If Len(t) = 0 Then ThrowErr "ValueError", "empty separator"
        If nm = "partition" Then p = InStr(1, s, t, vbBinaryCompare) Else p = InStrRev(s, t, -1, vbBinaryCompare)
        Set r = NewObj(OK_TUPLE)
        If p = 0 Then
            If nm = "partition" Then
                x = s
                r.Push x
                x = vbNullString
                r.Push x
                r.Push x
            Else
                x = vbNullString
                r.Push x
                r.Push x
                x = s
                r.Push x
            End If
        Else
            x = Left$(s, p - 1)
            r.Push x
            x = t
            r.Push x
            x = Mid$(s, p + Len(t))
            r.Push x
        End If
        Set out = r
    Case "removeprefix"
        t = ArgStr(args(0), "prefix")
        If Len(t) > 0 And Left$(s, Len(t)) = t Then
            out = Mid$(s, Len(t) + 1)
        Else
            out = s
        End If
    Case "removesuffix"
        t = ArgStr(args(0), "suffix")
        If Len(t) > 0 And Right$(s, Len(t)) = t And Len(t) <= Len(s) Then
            out = Left$(s, Len(s) - Len(t))
        Else
            out = s
        End If
    Case "expandtabs"
        k = 8
        If na > 0 Then k = IntToLong(args(0), "tabsize")
        n = 0
        t = vbNullString
        For i = 1 To Len(s)
            x = Mid$(s, i, 1)
            If x = vbTab Then
                If k > 0 Then
                    t = t & Space$(k - (n Mod k))
                    n = n + k - (n Mod k)
                End If
            ElseIf x = vbLf Then
                t = t & x
                n = 0
            Else
                t = t & x
                n = n + 1
            End If
        Next
        out = t
    Case "encode"
        out = s
    Case Else
        ThrowErr "AttributeError", "'str' object has no attribute '" & nm & "'"
    End Select
End Sub

Private Sub ListMethod(ByVal o As PyObj, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim i As Long, n As Long, x As Variant, src As PyObj, keyFn As Variant, rev As Boolean, st As Long, en As Long
    out = Empty
    Select Case nm
    Case "append"
        NeedArgs "append", na, 1, 1
        o.Push args(0)
    Case "appendleft"
        NeedArgs "appendleft", na, 1, 1
        o.InsertAt 0, args(0)
    Case "extend"
        NeedArgs "extend", na, 1, 1
        Set src = NewListFrom(args(0))
        o.Reserve o.Count + src.Count
        For i = 0 To src.Count - 1
            src.GetInto i, x
            o.Push x
        Next
    Case "extendleft"
        Set src = NewListFrom(args(0))
        For i = 0 To src.Count - 1
            src.GetInto i, x
            o.InsertAt 0, x
        Next
    Case "insert"
        NeedArgs "insert", na, 2, 2
        i = IntToLong(args(0), "index")
        If i < 0 Then i = i + o.Count
        If i < 0 Then i = 0
        If i > o.Count Then i = o.Count
        o.InsertAt i, args(1)
    Case "pop", "popleft"
        If o.Count = 0 Then ThrowErr "IndexError", "pop from empty list"
        If nm = "popleft" Then
            i = 0
        ElseIf na > 0 Then
            i = NormIndex(args(0), o.Count, "pop")
        Else
            i = o.Count - 1
        End If
        o.GetInto i, out
        o.RemoveAt i
    Case "remove"
        NeedArgs "remove", na, 1, 1
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If PyEq(x, args(0)) Then
                o.RemoveAt i
                Exit Sub
            End If
        Next
        ThrowErr "ValueError", "list.remove(x): x not in list"
    Case "index"
        NeedArgs "index", na, 1, 3
        SubRange o.Count, args, na, 1, st, en
        For i = st To en - 1
            o.GetInto i, x
            If PyEq(x, args(0)) Then
                out = i
                Exit Sub
            End If
        Next
        ThrowErr "ValueError", PyRepr(args(0)) & " is not in list"
    Case "count"
        NeedArgs "count", na, 1, 1
        n = 0
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If PyEq(x, args(0)) Then n = n + 1
        Next
        out = n
    Case "sort"
        NeedArgs "sort", na, 0, 0
        SortArgs kw, keyFn, rev
        SortList o, keyFn, rev
    Case "reverse"
        o.ReverseItems
    Case "rotate"
        n = 1
        If na > 0 Then n = IntToLong(args(0), "n")
        If o.Count > 0 Then
            n = n Mod o.Count
            If n < 0 Then n = n + o.Count
            For i = 1 To n
                o.GetInto o.Count - 1, x
                o.RemoveAt o.Count - 1
                o.InsertAt 0, x
            Next
        End If
    Case "clear"
        o.ClearItems
    Case "copy"
        Set out = NewListFrom(ToVar(o))
    Case Else
        ThrowErr "AttributeError", "'list' object has no attribute '" & nm & "'"
    End Select
End Sub

Private Sub TupleMethod(ByVal o As PyObj, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim i As Long, n As Long, x As Variant
    Select Case nm
    Case "count"
        NeedArgs "count", na, 1, 1
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If PyEq(x, args(0)) Then n = n + 1
        Next
        out = n
    Case "index"
        NeedArgs "index", na, 1, 1
        For i = 0 To o.Count - 1
            o.GetInto i, x
            If PyEq(x, args(0)) Then
                out = i
                Exit Sub
            End If
        Next
        ThrowErr "ValueError", "tuple.index(x): x not in tuple"
    Case Else
        ThrowErr "AttributeError", "'tuple' object has no attribute '" & nm & "'"
    End Select
End Sub

Private Sub DictMethod(ByVal o As PyObj, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim r As PyObj, i As Long, k As Variant, v As Variant, e As Long, n As Long, keyFn As Variant, it As PyObj, x As Variant
    Select Case nm
    Case "keys"
        Set out = DictKeysList(o)
    Case "values", "items"
        Set r = NewList()
        r.Reserve o.DCount
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.ValInto i, v
                If nm = "values" Then
                    r.Push v
                Else
                    o.KeyInto i, k
                    r.Push ToVar(NewTuple2(k, v))
                End If
            End If
        Next
        Set out = r
    Case "get"
        NeedArgs "get", na, 1, 2
        If Not DictGet(o, args(0), out) Then
            If na = 2 Then
                SetV out, args(1)
            Else
                out = Empty
            End If
        End If
    Case "pop"
        NeedArgs "pop", na, 1, 2
        e = DictFind(o, args(0))
        If e >= 0 Then
            o.ValInto e, out
            o.DelAt e
        ElseIf na = 2 Then
            SetV out, args(1)
        Else
            RaiseExc MakeExcV("KeyError", args(0))
        End If
    Case "popitem"
        For i = o.EntryCount - 1 To 0 Step -1
            If o.LiveAt(i) Then
                o.KeyInto i, k
                o.ValInto i, v
                o.DelAt i
                Set out = NewTuple2(k, v)
                Exit Sub
            End If
        Next
        ThrowErr "KeyError", "popitem(): dictionary is empty"
    Case "setdefault"
        NeedArgs "setdefault", na, 1, 2
        If Not DictGet(o, args(0), out) Then
            If na = 2 Then
                SetV out, args(1)
            Else
                out = Empty
            End If
            DictSet o, args(0), out
        End If
    Case "update"
        If na > 0 Then
            If o.Flags = 2 Then
                CounterAdd o, args(0), 1
            Else
                DictUpdateFrom o, args(0)
            End If
        End If
        If Not (kw Is Nothing) Then DictUpdate o, kw
        out = Empty
    Case "clear"
        o.DClear
        out = Empty
    Case "copy"
        Set r = CopyDict(o, OK_DICT)
        r.Flags = o.Flags
        SetV r.FnV, o.FnV
        Set out = r
    Case "most_common"
        Set r = NewList()
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.KeyInto i, k
                o.ValInto i, v
                r.Push ToVar(NewTuple2(k, v))
            End If
        Next
        Set keyFn = NewObj(OK_BUILTIN)
        SortByItem1Desc r
        If na > 0 Then
            If Not IsEmpty(args(0)) Then
                n = IntToLong(args(0), "n")
                Do While r.Count > n And r.Count > 0
                    r.RemoveAt r.Count - 1
                Loop
            End If
        End If
        Set out = r
    Case "elements"
        Set r = NewList()
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.KeyInto i, k
                o.ValInto i, v
                For n = 1 To IntToLong(v, "count")
                    r.Push k
                Next
            End If
        Next
        Set out = r
    Case "total"
        x = 0&
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.ValInto i, v
                BinOp OP_ADD, x, v, k
                SetV x, k
            End If
        Next
        SetV out, x
    Case Else
        ThrowErr "AttributeError", "'dict' object has no attribute '" & nm & "'"
    End Select
End Sub

Public Function CounterRepr(ByVal o As PyObj) As String
    Dim r As PyObj, i As Long, k As Variant, v As Variant, parts() As String, t As PyObj
    Set r = NewList()
    For i = 0 To o.EntryCount - 1
        If o.LiveAt(i) Then
            o.KeyInto i, k
            o.ValInto i, v
            r.Push ToVar(NewTuple2(k, v))
        End If
    Next
    SortByItem1Desc r
    ReDim parts(0 To r.Count - 1)
    For i = 0 To r.Count - 1
        r.GetInto i, v
        Set t = v
        t.GetInto 0, k
        t.GetInto 1, v
        parts(i) = PyRepr(k) & ": " & PyRepr(v)
    Next
    CounterRepr = "{" & Join(parts, ", ") & "}"
End Function

Private Sub SortByItem1Desc(ByVal r As PyObj)
    ' stable sort of (key, count) tuples by count, descending
    Dim n As Long, items() As Variant, keys() As Variant, idx() As Long, tmp() As Long, i As Long, res() As Variant, t As PyObj, x As Variant
    n = r.Count
    If n < 2 Then Exit Sub
    r.GetArray items
    ReDim keys(0 To n - 1)
    For i = 0 To n - 1
        Set t = items(i)
        t.GetInto 1, x
        SetV keys(i), x
    Next
    ReDim idx(0 To n - 1)
    ReDim tmp(0 To n - 1)
    For i = 0 To n - 1
        idx(i) = i
    Next
    MergeSortIdx idx, tmp, keys, 0, n - 1, True
    ReDim res(0 To n - 1)
    For i = 0 To n - 1
        SetV res(i), items(idx(i))
    Next
    r.SetArray res, n
End Sub

Private Function SetFrom(ByRef v As Variant) As PyObj
    Dim r As PyObj, it As PyObj, x As Variant
    If IsKind(v, OK_SET) Then
        Set SetFrom = v
        Exit Function
    End If
    Set r = NewObj(OK_SET)
    Set it = GetIter(v)
    Do While IterNext(it, x)
        SetAdd r, x
    Loop
    Set SetFrom = r
End Function

Private Sub SetMethod(ByVal o As PyObj, ByVal nm As String, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim r As PyObj, i As Long, k As Variant, b As PyObj, e As Variant, op As Long
    out = Empty
    Select Case nm
    Case "add"
        NeedArgs "add", na, 1, 1
        SetAdd o, args(0)
    Case "remove"
        NeedArgs "remove", na, 1, 1
        If Not DictDel(o, args(0)) Then RaiseExc MakeExcV("KeyError", args(0))
    Case "discard"
        NeedArgs "discard", na, 1, 1
        DictDel o, args(0)
    Case "pop"
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.KeyInto i, out
                o.DelAt i
                Exit Sub
            End If
        Next
        ThrowErr "KeyError", "pop from an empty set"
    Case "clear"
        o.DClear
    Case "copy"
        Set out = CopyDict(o, OK_SET)
    Case "union", "intersection", "difference", "symmetric_difference", "update", "intersection_update", "difference_update", "symmetric_difference_update"
        Select Case nm
        Case "union", "update": op = OP_BITOR
        Case "intersection", "intersection_update": op = OP_BITAND
        Case "difference", "difference_update": op = OP_SUB
        Case Else: op = OP_BITXOR
        End Select
        Set r = o
        For i = 0 To na - 1
            Set b = SetFrom(args(i))
            BinOp op, ToVar(r), ToVar(b), e
            Set r = e
        Next
        If na = 0 Then Set r = CopyDict(o, OK_SET)
        If Right$(nm, 7) = "_update" Or nm = "update" Then
            o.DClear
            For i = 0 To r.EntryCount - 1
                If r.LiveAt(i) Then
                    r.KeyInto i, k
                    o.SetH r.KSAt(i), r.HAt(i), k, e
                End If
            Next
            out = Empty
        Else
            Set out = r
        End If
    Case "issubset"
        out = SetSubset(o, SetFrom(args(0)))
    Case "issuperset"
        out = SetSubset(SetFrom(args(0)), o)
    Case "isdisjoint"
        Set b = SetFrom(args(0))
        out = True
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                If b.FindH(o.KSAt(i), o.HAt(i)) >= 0 Then
                    out = False
                    Exit For
                End If
            End If
        Next
    Case Else
        ThrowErr "AttributeError", "'set' object has no attribute '" & nm & "'"
    End Select
End Sub
