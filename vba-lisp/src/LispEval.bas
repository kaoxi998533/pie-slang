Attribute VB_Name = "LispEval"
Option Explicit
' ===========================================================================
'  LispEval - the evaluator.  LEval runs in a loop so that calls in tail
'  position (if/cond/case/and/or/when/unless/let*/begin bodies, procedure
'  bodies, named let ...) do not grow the VBA call stack.
'
'  Entry points:
'    LispRunProgram(src, echo)  run a program, return its output (+ error)
'    LispEvalCode(...)          used by the LISP() worksheet function
' ===========================================================================

Public Const LISP_DEFAULT_MAX_DEPTH As Long = 2000

' ---------------------------------------------------------------------------
'  setup / entry points
' ---------------------------------------------------------------------------
Private Sub SpecialSym(ByVal nm As String, ByVal code As Long)
    Dim s As LispObj
    Set s = LIntern(nm)
    s.Code = code
End Sub

Public Sub LispReset()
    LReleaseAll
    lgNextId = 0
    lgGensym = 0
    Set lgSymtab = LNewObj(LK_ENV)
    Set lgVoid = LNewObj(LK_VOID)
    Set lgEof = LNewObj(LK_EOF)
    Set lgGlobal = LNewEnv(Nothing)
    lgDepth = 0
    lgLine = 0
    If lgMaxDepth <= 0 Then lgMaxDepth = LISP_DEFAULT_MAX_DEPTH
    lgDecSep = Mid$(Format$(1.5, "0.0"), 2, 1)
    LOutReset
    SpecialSym "quote", LSF_QUOTE
    SpecialSym "quasiquote", LSF_QUASIQUOTE
    SpecialSym "unquote", LSF_UNQUOTE
    SpecialSym "unquote-splicing", LSF_UNQUOTESPLICING
    SpecialSym "if", LSF_IF
    SpecialSym "define", LSF_DEFINE
    SpecialSym "set!", LSF_SET
    SpecialSym "lambda", LSF_LAMBDA
    SpecialSym "named-lambda", LSF_LAMBDA
    SpecialSym "begin", LSF_BEGIN
    SpecialSym "let", LSF_LET
    SpecialSym "let*", LSF_LETSTAR
    SpecialSym "letrec", LSF_LETREC
    SpecialSym "letrec*", LSF_LETREC
    SpecialSym "let-syntax", LSF_LET
    SpecialSym "letrec-syntax", LSF_LETREC
    SpecialSym "cond", LSF_COND
    SpecialSym "case", LSF_CASE
    SpecialSym "and", LSF_AND
    SpecialSym "or", LSF_OR
    SpecialSym "when", LSF_WHEN
    SpecialSym "unless", LSF_UNLESS
    SpecialSym "do", LSF_DO
    SpecialSym "delay", LSF_DELAY
    SpecialSym "make-promise-lazy", LSF_DELAYFORCE
    SpecialSym "delay-force", LSF_DELAYFORCE
    SpecialSym "cons-stream", LSF_CONSSTREAM
    SpecialSym "define-syntax", LSF_DEFSYNTAX
    SpecialSym "syntax-rules", LSF_SYNTAXRULES
    SpecialSym "define-macro", LSF_DEFMACRO
    SpecialSym "defmacro", LSF_DEFMACRO
    SpecialSym "guard", LSF_GUARD
    SpecialSym "define-record-type", LSF_DEFRECORD
    SpecialSym "assert", LSF_ASSERT
    SpecialSym "case-lambda", LSF_CASELAMBDA
    SpecialSym "let-values", LSF_LETVALUES
    SpecialSym "let*-values", LSF_LETVALUES
    SpecialSym "define-values", LSF_DEFVALUES
    SpecialSym "use-modules", LSF_IGNORE
    SpecialSym "import", LSF_IGNORE
    SpecialSym "require", LSF_IGNORE
    Set lsQuote = LIntern("quote")
    Set lsQuasi = LIntern("quasiquote")
    Set lsUnquote = LIntern("unquote")
    Set lsUnquoteSplicing = LIntern("unquote-splicing")
    Set lsElse = LIntern("else")
    Set lsArrow = LIntern("=>")
    Set lsEllipsis = LIntern("...")
    Set lsUnderscore = LIntern("_")
    Set lsLambda = LIntern("lambda")
    Set lsBegin = LIntern("begin")
    Set lsDefine = LIntern("define")
    Set lsDot = LIntern(".")
    LInitBuiltins
End Sub

Public Sub LispFinish()
    LReleaseAll
    Set lgSymtab = Nothing
End Sub

Private Function ErrorText(ByRef c As Variant) As String
    Dim o As LispObj, s As String
    If LIsKind(c, LK_ERROR) Then
        Set o = c
        If o.Flags = 99 Then Exit Function    ' (exit)
        s = "Error: " & LCondReport(c)
        If o.I1 > 0 Then s = s & "  [line " & o.I1 & "]"
    Else
        s = "Error: uncaught exception: " & LToStr(c, True)
        If lgLine > 0 Then s = s & "  [line " & lgLine & "]"
    End If
    ErrorText = s
End Function

' Run a whole program and return everything it printed. With echo=True the
' value of every top-level expression is shown as "=> value" (REPL style).
Public Function LispRunProgram(ByVal src As String, Optional ByVal echo As Boolean = False) As String
    Dim forms As Variant, x As Variant, fo As LispObj, v As Variant, cond As Variant, msg As String, outS As String
    LispReset
    On Error GoTo EH
    LSetV forms, LReadAll(src)
    LSetV x, forms
    Do While LIsPair(x)
        Set fo = x
        lgLine = fo.Ln
        lgDepth = 0
        LEval fo.Car, lgGlobal, v
        If echo Then EchoValue v
        LSetV x, fo.Cdr
    Loop
    LispRunProgram = LOutText()
    LispFinish
    Exit Function
EH:
    If Err.Number = LESCAPE Then
        msg = "Error: a continuation was invoked outside of its extent (only escaping continuations are supported)"
    Else
        LSetV cond, LCapture()
    End If
    Resume Report
Report:
    On Error Resume Next
    If Len(msg) = 0 Then msg = ErrorText(cond)
    outS = LOutText()
    If Len(msg) > 0 Then
        If Len(outS) > 0 And Right$(outS, 1) <> vbLf Then outS = outS & vbLf
        outS = outS & msg & vbLf
    End If
    LispRunProgram = outS
    LispFinish
End Function

Private Sub EchoValue(ByRef v As Variant)
    Dim o As LispObj
    If IsObject(v) Then
        Set o = v
        If o.Kind = LK_VOID Then Exit Sub
        If o.Kind = LK_VALUES And o.Count = 0 Then Exit Sub
    End If
    If Not LOutAtLineStart() Then LOutWrite vbLf
    LOutWrite "=> " & LToStr(v, True) & vbLf
End Sub

' Evaluate code for a worksheet formula: the arguments are bound to x, y, z
' and the list args; the result is the value of the last expression.
' The caller must call LispFinish after converting the result.
Public Function LispEvalCode(ByVal src As String, ByRef argv() As Variant, ByVal nArgs As Long, _
                             ByRef result As Variant, ByRef printed As String, ByRef errText As String) As Boolean
    Dim forms As Variant, x As Variant, fo As LispObj, i As Long, cond As Variant, s As LispObj, lst As Variant
    LispReset
    On Error GoTo EH
    For i = 0 To nArgs - 1
        If i < 3 Then
            Set s = LIntern(Mid$("xyz", i + 1, 1))
            lgGlobal.SetRef s, s.Id, argv(i)
        End If
    Next
    LSetV lst, LListFromArray(argv, 0, nArgs - 1)
    Set s = LIntern("args")
    lgGlobal.SetRef s, s.Id, lst
    LSetV forms, LReadAll(src)
    result = Empty
    Set result = lgVoid
    LSetV x, forms
    Do While LIsPair(x)
        Set fo = x
        lgDepth = 0
        LEval fo.Car, lgGlobal, result
        LSetV x, fo.Cdr
    Loop
    printed = LOutText()
    LispEvalCode = True
    Exit Function
EH:
    If Err.Number = LESCAPE Then
        errText = "continuation invoked outside of its extent"
    Else
        LSetV cond, LCapture()
        errText = LCondReport(cond)
    End If
    printed = LOutText()
End Function

' ---------------------------------------------------------------------------
'  variables
' ---------------------------------------------------------------------------
Public Sub LLookup(ByVal env As LispObj, ByVal s As LispObj, ByRef out As Variant)
    Dim e As LispObj, i As Long
    Set e = env
    Do
        i = e.FindRef(s, s.Id)
        If i >= 0 Then
            e.ValInto i, out
            Exit Sub
        End If
        Set e = e.Env
    Loop Until e Is Nothing
    LThrow2 "unbound variable:", s
End Sub

Private Function TryLookup(ByVal env As LispObj, ByVal s As LispObj, ByRef out As Variant) As Boolean
    Dim e As LispObj, i As Long
    Set e = env
    Do While Not (e Is Nothing)
        i = e.FindRef(s, s.Id)
        If i >= 0 Then
            e.ValInto i, out
            TryLookup = True
            Exit Function
        End If
        Set e = e.Env
    Loop
End Function

Public Sub LDefine(ByVal env As LispObj, ByVal s As LispObj, ByRef v As Variant)
    env.SetRef s, s.Id, v
End Sub

Private Sub SetVar(ByVal env As LispObj, ByVal s As LispObj, ByRef v As Variant)
    Dim e As LispObj, i As Long
    Set e = env
    Do While Not (e Is Nothing)
        i = e.FindRef(s, s.Id)
        If i >= 0 Then
            e.SetValAt i, v
            Exit Sub
        End If
        Set e = e.Env
    Loop
    LThrow2 "set!: unbound variable:", s
End Sub

' ---------------------------------------------------------------------------
'  small list helpers for syntax
' ---------------------------------------------------------------------------
Private Sub SynErr(ByRef form As Variant)
    LThrow2 "bad syntax:", form
End Sub

Private Function Nth(ByVal o As LispObj, ByVal n As Long) As Variant
    ' n-th element (0-based) of the list o; Empty if too short
    Dim x As Variant, i As Long, p As LispObj
    Nth = Empty
    LSetV x, o
    For i = 1 To n
        If Not LIsPair(x) Then Exit Function
        Set p = x
        LSetV x, p.Cdr
    Next
    If LIsPair(x) Then
        Set p = x
        If IsObject(p.Car) Then Set Nth = p.Car Else Nth = p.Car
    End If
End Function

Private Function NthCdr(ByVal o As LispObj, ByVal n As Long) As Variant
    Dim x As Variant, i As Long, p As LispObj
    NthCdr = Empty
    LSetV x, o
    For i = 1 To n
        If Not LIsPair(x) Then Exit Function
        Set p = x
        LSetV x, p.Cdr
    Next
    If IsObject(x) Then Set NthCdr = x Else NthCdr = x
End Function

Private Function NeedSym(ByRef v As Variant, ByRef form As Variant) As LispObj
    If Not LIsSym(v) Then SynErr form
    Set NeedSym = v
End Function

' evaluate a sequence, value of the last expression (void if empty)
Public Sub LEvalSeq(ByRef body As Variant, ByVal env As LispObj, ByRef out As Variant)
    Dim x As Variant, p As LispObj
    Set out = lgVoid
    LSetV x, body
    Do While LIsPair(x)
        Set p = x
        LEval p.Car, env, out
        LSetV x, p.Cdr
    Loop
End Sub

' evaluate all but the last expression; the last one is returned in nx
Private Function TailSeq(ByRef body As Variant, ByVal env As LispObj, ByRef out As Variant, ByRef nx As Variant) As Boolean
    Dim x As Variant, p As LispObj, v As Variant
    LSetV x, body
    If Not LIsPair(x) Then
        Set out = lgVoid
        Exit Function
    End If
    Do
        Set p = x
        If LIsPair(p.Cdr) Then
            LEval p.Car, env, v
            LSetV x, p.Cdr
        Else
            LSetV nx, p.Car
            TailSeq = True
            Exit Function
        End If
    Loop
End Function

' ---------------------------------------------------------------------------
'  procedures
' ---------------------------------------------------------------------------
Public Function LMakeLambda(ByRef params As Variant, ByRef body As Variant, ByVal env As LispObj, ByVal nm As String) As LispObj
    Dim fo As LispObj, x As Variant, p As LispObj
    If Not LIsPair(body) Then LThrow "lambda: the body is empty"
    Set fo = LNewObj(LK_LAMBDA)
    fo.SName = nm
    Set fo.Env = env
    fo.SetCdr body
    fo.SetCar params
    LSetV x, params
    Do While LIsPair(x)
        Set p = x
        If Not LIsSym(p.Car) Then LThrow2 "lambda: bad parameter list", params
        fo.Push p.Car
        LSetV x, p.Cdr
    Loop
    If LIsSym(x) Then
        Set fo.Ref = x
    ElseIf Not IsEmpty(x) Then
        LThrow2 "lambda: bad parameter list", params
    End If
    fo.I1 = fo.Count
    If Not (env Is lgGlobal) Then LKeep env
    Set LMakeLambda = fo
End Function

Private Sub ArityError(ByVal fo As LispObj, ByVal na As Long)
    Dim nm As String, want As String
    nm = fo.SName
    If Len(nm) = 0 Then nm = "#<procedure>"
    If fo.Ref Is Nothing Then
        want = CStr(fo.I1)
    Else
        want = "at least " & fo.I1
    End If
    LThrow nm & ": wrong number of arguments (expected " & want & ", got " & na & ")"
End Sub

Private Function BindArgs(ByVal fo As LispObj, ByRef args() As Variant, ByVal na As Long) As LispObj
    Dim ne As LispObj, i As Long, n As Long, sv As Variant, s As LispObj, rest As Variant
    n = fo.I1
    If na < n Then ArityError fo, na
    If na > n Then
        If fo.Ref Is Nothing Then ArityError fo, na
    End If
    Set ne = LNewEnv(fo.Env)
    For i = 0 To n - 1
        fo.GetInto i, sv
        Set s = sv
        ne.SetRef s, s.Id, args(i)
    Next
    If Not (fo.Ref Is Nothing) Then
        LSetV rest, LListFromArray(args, n, na - 1)
        ne.SetRef fo.Ref, fo.Ref.Id, rest
    End If
    Set BindArgs = ne
End Function

Private Function PickCase(ByVal fo As LispObj, ByVal na As Long) As LispObj
    Dim i As Long, cv As Variant, c As LispObj
    For i = 0 To fo.Count - 1
        fo.GetInto i, cv
        Set c = cv
        If na = c.I1 Then
            Set PickCase = c
            Exit Function
        End If
        If na > c.I1 And Not (c.Ref Is Nothing) Then
            Set PickCase = c
            Exit Function
        End If
    Next
    LThrow "case-lambda: no clause accepts " & na & " argument(s)"
End Function

' bind a formals list (as in lambda) to values; used by let-values / define-values
Private Sub BindFormals(ByRef formals As Variant, ByRef vals() As Variant, ByVal n As Long, ByVal env As LispObj)
    Dim x As Variant, p As LispObj, i As Long, s As LispObj, rest As Variant
    LSetV x, formals
    Do While LIsPair(x)
        Set p = x
        If i >= n Then LThrow "let-values: too few values"
        Set s = NeedSym(p.Car, formals)
        env.SetRef s, s.Id, vals(i)
        i = i + 1
        LSetV x, p.Cdr
    Loop
    If LIsSym(x) Then
        Set s = x
        LSetV rest, LListFromArray(vals, i, n - 1)
        env.SetRef s, s.Id, rest
    ElseIf i <> n Then
        LThrow "let-values: too many values"
    End If
End Sub

Private Sub ValuesToArray(ByRef v As Variant, ByRef arr() As Variant, ByRef n As Long)
    Dim o As LispObj, i As Long, x As Variant
    If LIsKind(v, LK_VALUES) Then
        Set o = v
        n = o.Count
        ReDim arr(0 To IIf(n > 0, n - 1, 0))
        For i = 0 To n - 1
            o.GetInto i, x
            If IsObject(x) Then
                Set arr(i) = x
            Else
                arr(i) = x
            End If
        Next
    Else
        n = 1
        ReDim arr(0 To 0)
        If IsObject(v) Then
            Set arr(0) = v
        Else
            arr(0) = v
        End If
    End If
End Sub

' apply a procedure (not in tail position)
Public Sub LApply(ByRef f As Variant, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim fo As LispObj, ne As LispObj
    If Not IsObject(f) Then LThrow2 "not a procedure:", f
    Set fo = f
    Do
        Select Case fo.Kind
        Case LK_LAMBDA
            If fo.Flags = 4 Then
                Set fo = PickCase(fo, na)
            Else
                Set ne = BindArgs(fo, args, na)
                LEvalSeq fo.Cdr, ne, out
                Exit Sub
            End If
        Case LK_PRIM, LK_RECPROC
            LCallPrim fo, args, na, out
            Exit Sub
        Case LK_CONT
            LInvokeCont fo, args, na
            Exit Sub
        Case Else
            LThrow2 "not a procedure:", f
        End Select
    Loop
End Sub

Public Sub LApply0(ByRef f As Variant, ByRef out As Variant)
    Dim a() As Variant
    ReDim a(0 To 0)
    LApply f, a, 0, out
End Sub

Public Sub LApply1(ByRef f As Variant, ByRef x As Variant, ByRef out As Variant)
    Dim a() As Variant
    ReDim a(0 To 0)
    If IsObject(x) Then
        Set a(0) = x
    Else
        a(0) = x
    End If
    LApply f, a, 1, out
End Sub

Public Sub LApply2(ByRef f As Variant, ByRef x As Variant, ByRef y As Variant, ByRef out As Variant)
    Dim a() As Variant
    ReDim a(0 To 1)
    If IsObject(x) Then
        Set a(0) = x
    Else
        a(0) = x
    End If
    If IsObject(y) Then
        Set a(1) = y
    Else
        a(1) = y
    End If
    LApply f, a, 2, out
End Sub

' ---------------------------------------------------------------------------
'  the evaluator
' ---------------------------------------------------------------------------
Public Sub LEval(ByVal x0 As Variant, ByVal env0 As LispObj, ByRef out As Variant)
    Dim o As LispObj, hs As LispObj, fo As LispObj, ro As LispObj, ne As LispObj
    Dim hd As Variant, f As Variant, rest As Variant, v As Variant, t As Variant
    Dim args() As Variant, na As Long, cap As Long, i As Long, lst As Variant
    Dim x As Variant, env As LispObj
    If IsObject(x0) Then Set x = x0 Else x = x0
    Set env = env0
    lgDepth = lgDepth + 1
    If lgDepth > lgMaxDepth Then LThrow "maximum recursion depth exceeded"
    Do
        If Not IsObject(x) Then
            out = x
            Exit Do
        End If
        Set o = x
        If o.Kind = LK_SYM Then
            LLookup env, o, out
            Exit Do
        End If
        If o.Kind <> LK_PAIR Then
            Set out = o
            Exit Do
        End If
        If o.Ln > 0 Then lgLine = o.Ln
        LSetV hd, o.Car
        ' ---- special forms and the operator ----
        If IsObject(hd) Then
            Set hs = hd
            If hs.Kind = LK_SYM Then
                If hs.Code <> 0 Then
                    If SpecialForm(hs.Code, o, env, out, x, env) Then GoTo NextIter
                    Exit Do
                End If
                LLookup env, hs, f
            Else
                LEval hd, env, f
            End If
        Else
            f = hd
        End If
        If Not IsObject(f) Then LThrow2 "not a procedure:", f
        Set fo = f
        If fo.Kind = LK_MACRO Then
            ExpandMacro fo, o, env, t
            MemoizeExpansion o, t
            GoTo NextIter
        End If
        ' ---- evaluate the operands ----
        na = 0
        LSetV rest, o.Cdr
        Do While IsObject(rest)
            Set ro = rest
            If ro.Kind <> LK_PAIR Then Exit Do
            If na >= cap Then
                If cap = 0 Then
                    cap = 8
                    ReDim args(0 To 7)
                Else
                    cap = cap * 2
                    ReDim Preserve args(0 To cap - 1)
                End If
            End If
            LSetV v, ro.Car
            If IsObject(v) Then
                Set hs = v
                If hs.Kind = LK_SYM Then
                    LLookup env, hs, t
                ElseIf hs.Kind = LK_PAIR Then
                    LEval v, env, t
                Else
                    Set t = v
                End If
                If IsObject(t) Then
                    Set args(na) = t
                Else
                    args(na) = t
                End If
            Else
                args(na) = v
            End If
            na = na + 1
            LSetV rest, ro.Cdr
        Loop
        If cap = 0 Then
            cap = 1
            ReDim args(0 To 0)
        End If
        ' ---- apply ----
ApplyF:
        Select Case fo.Kind
        Case LK_LAMBDA
            If fo.Flags = 4 Then
                Set fo = PickCase(fo, na)
                GoTo ApplyF
            End If
            Set ne = BindArgs(fo, args, na)
            LSetV rest, fo.Cdr
            Do
                Set ro = rest
                If LIsPair(ro.Cdr) Then
                    LEval ro.Car, ne, v
                    LSetV rest, ro.Cdr
                Else
                    LSetV x, ro.Car
                    Exit Do
                End If
            Loop
            Set env = ne
        Case LK_PRIM
            If fo.Code = LP_APPLY Then
                ' (apply f a b ... lst): keep proper tail calls through apply
                If na < 1 Then LThrow "apply: expected at least 1 argument"
                If Not LIsProc(args(0)) Then LThrow2 "apply: not a procedure:", args(0)
                Set fo = args(0)
                SpreadApplyArgs args, na
                cap = UBound(args) + 1
                GoTo ApplyF
            End If
            LCallPrim fo, args, na, out
            Exit Do
        Case LK_RECPROC
            LCallPrim fo, args, na, out
            Exit Do
        Case LK_CONT
            LInvokeCont fo, args, na
        Case Else
            LThrow2 "not a procedure:", f
        End Select
NextIter:
    Loop
    lgDepth = lgDepth - 1
End Sub

' args = (f a b ... lst)  ->  args = (a b ... items of lst)
Public Sub SpreadApplyArgs(ByRef args() As Variant, ByRef na As Long)
    Dim nw() As Variant, n As Long, i As Long, x As Variant, p As LispObj
    ReDim nw(0 To 7)
    For i = 1 To na - 2
        If n > UBound(nw) Then
            ReDim Preserve nw(0 To 2 * n + 1)
        End If
        If IsObject(args(i)) Then
            Set nw(n) = args(i)
        Else
            nw(n) = args(i)
        End If
        n = n + 1
    Next
    If na >= 2 Then
        LSetV x, args(na - 1)
        Do While LIsPair(x)
            Set p = x
            If n > UBound(nw) Then
                ReDim Preserve nw(0 To 2 * n + 1)
            End If
            If IsObject(p.Car) Then
                Set nw(n) = p.Car
            Else
                nw(n) = p.Car
            End If
            n = n + 1
            LSetV x, p.Cdr
        Loop
        If Not IsEmpty(x) Then LThrow2 "apply: the last argument must be a list, got", args(na - 1)
    End If
    args = nw
    na = n
End Sub

' replace a macro use by its expansion, so that it is expanded only once
Private Sub MemoizeExpansion(ByVal o As LispObj, ByRef expansion As Variant)
    Dim p As LispObj, e As Variant
    If LIsPair(expansion) Then
        Set p = expansion
        o.SetCar p.Car
        o.SetCdr p.Cdr
    Else
        Set o.Car = lsBegin
        Set o.Cdr = LCons(expansion, e)
    End If
End Sub

' ---------------------------------------------------------------------------
'  special forms: return True for a tail call (nx / nenv set), False when
'  the value has been stored in out
' ---------------------------------------------------------------------------
Private Function SpecialForm(ByVal code As Long, ByVal o As LispObj, ByVal env As LispObj, ByRef out As Variant, _
                             ByRef nx As Variant, ByRef nenv As LispObj) As Boolean
    Dim t As Variant, v As Variant, x As Variant, p As LispObj, s As LispObj, ne As LispObj, cl As LispObj
    Dim n As Long, i As Long, fo As LispObj, a As Variant, e As Variant
    n = LLength(o)
    If n < 0 Then SynErr o
    Select Case code
    Case LSF_QUOTE
        If n <> 2 Then SynErr o
        LSetV out, Nth(o, 1)
    Case LSF_IF
        If n < 3 Or n > 4 Then SynErr o
        LEval Nth(o, 1), env, t
        If LTruthy(t) Then
            LSetV nx, Nth(o, 2)
            SpecialForm = True
        ElseIf n = 4 Then
            LSetV nx, Nth(o, 3)
            SpecialForm = True
        Else
            Set out = lgVoid
        End If
    Case LSF_DEFINE
        If n < 2 Then SynErr o
        LSetV t, Nth(o, 1)
        If LIsPair(t) Then
            Set p = t
            If LIsPair(p.Car) Then LThrow2 "define: curried definitions are not supported", o
            Set s = NeedSym(p.Car, o)
            Set v = LMakeLambda(p.Cdr, NthCdr(o, 2), env, s.SName)
        Else
            Set s = NeedSym(t, o)
            If n = 2 Then
                Set v = lgVoid
            ElseIf n = 3 Then
                LEval Nth(o, 2), env, v
                NameProc v, s.SName
            Else
                SynErr o
            End If
        End If
        env.SetRef s, s.Id, v
        Set out = lgVoid
    Case LSF_SET
        If n <> 3 Then SynErr o
        Set s = NeedSym(Nth(o, 1), o)
        LEval Nth(o, 2), env, v
        SetVar env, s, v
        Set out = lgVoid
    Case LSF_LAMBDA
        If n < 3 Then SynErr o
        Set s = o.Car
        LSetV t, Nth(o, 1)
        If s.SName = "named-lambda" Then
            ' (named-lambda (name . params) body...)
            If Not LIsPair(t) Then SynErr o
            Set p = t
            Set out = LMakeLambda(p.Cdr, NthCdr(o, 2), env, LToStr(p.Car, False))
        Else
            Set out = LMakeLambda(t, NthCdr(o, 2), env, vbNullString)
        End If
    Case LSF_BEGIN
        SpecialForm = TailSeq(o.Cdr, env, out, nx)
    Case LSF_LET
        If n < 3 Then SynErr o
        LSetV t, Nth(o, 1)
        If LIsSym(t) Then
            SpecialForm = NamedLet(o, env, out, nx, nenv)
            Exit Function
        End If
        Set ne = LNewEnv(env)
        LSetV x, t
        Do While LIsPair(x)
            Set p = x
            BindingInit p.Car, env, ne, o
            LSetV x, p.Cdr
        Loop
        If Not IsEmpty(x) Then SynErr o
        Set nenv = ne
        SpecialForm = TailSeq(NthCdr(o, 2), ne, out, nx)
    Case LSF_LETSTAR, LSF_LETREC
        If n < 3 Then SynErr o
        Set ne = LNewEnv(env)
        LSetV x, Nth(o, 1)
        Do While LIsPair(x)
            Set p = x
            BindingInit p.Car, ne, ne, o
            LSetV x, p.Cdr
        Loop
        If Not IsEmpty(x) Then SynErr o
        Set nenv = ne
        SpecialForm = TailSeq(NthCdr(o, 2), ne, out, nx)
    Case LSF_COND
        LSetV x, o.Cdr
        Do While LIsPair(x)
            Set p = x
            If Not LIsPair(p.Car) Then SynErr o
            Set cl = p.Car
            If IsObject(cl.Car) Then
                If cl.Car Is lsElse Then
                    SpecialForm = TailSeq(cl.Cdr, env, out, nx)
                    Exit Function
                End If
            End If
            LEval cl.Car, env, t
            If LTruthy(t) Then
                If IsEmpty(cl.Cdr) Then
                    LSetV out, t
                ElseIf IsArrowClause(cl) Then
                    LEval Nth(cl, 2), env, v
                    LApply1 v, t, out
                Else
                    SpecialForm = TailSeq(cl.Cdr, env, out, nx)
                End If
                Exit Function
            End If
            LSetV x, p.Cdr
        Loop
        Set out = lgVoid
    Case LSF_CASE
        If n < 2 Then SynErr o
        LEval Nth(o, 1), env, t
        LSetV x, NthCdr(o, 2)
        Do While LIsPair(x)
            Set p = x
            If Not LIsPair(p.Car) Then SynErr o
            Set cl = p.Car
            If CaseMatches(cl.Car, t) Then
                If IsArrowClause(cl) Then
                    LEval Nth(cl, 2), env, v
                    LApply1 v, t, out
                Else
                    SpecialForm = TailSeq(cl.Cdr, env, out, nx)
                End If
                Exit Function
            End If
            LSetV x, p.Cdr
        Loop
        Set out = lgVoid
    Case LSF_AND
        If n = 1 Then
            out = True
            Exit Function
        End If
        LSetV x, o.Cdr
        Do
            Set p = x
            If Not LIsPair(p.Cdr) Then
                LSetV nx, p.Car
                SpecialForm = True
                Exit Function
            End If
            LEval p.Car, env, t
            If Not LTruthy(t) Then
                out = False
                Exit Function
            End If
            LSetV x, p.Cdr
        Loop
    Case LSF_OR
        If n = 1 Then
            out = False
            Exit Function
        End If
        LSetV x, o.Cdr
        Do
            Set p = x
            If Not LIsPair(p.Cdr) Then
                LSetV nx, p.Car
                SpecialForm = True
                Exit Function
            End If
            LEval p.Car, env, t
            If LTruthy(t) Then
                LSetV out, t
                Exit Function
            End If
            LSetV x, p.Cdr
        Loop
    Case LSF_WHEN, LSF_UNLESS
        If n < 2 Then SynErr o
        LEval Nth(o, 1), env, t
        If LTruthy(t) = (code = LSF_WHEN) Then
            SpecialForm = TailSeq(NthCdr(o, 2), env, out, nx)
        Else
            Set out = lgVoid
        End If
    Case LSF_DO
        SpecialForm = EvalDo(o, env, out, nx, nenv)
    Case Else
        SpecialForm = SpecialForm2(code, o, n, env, out, nx, nenv)
    End Select
End Function

' (test => receiver) clause?
Private Function IsArrowClause(ByVal cl As LispObj) As Boolean
    Dim v As Variant
    If LLength(cl) <> 3 Then Exit Function
    LSetV v, Nth(cl, 1)
    If IsObject(v) Then IsArrowClause = (v Is lsArrow)
End Function

Private Sub NameProc(ByRef v As Variant, ByVal nm As String)
    Dim fo As LispObj
    If IsObject(v) Then
        Set fo = v
        If fo.Kind = LK_LAMBDA Or fo.Kind = LK_MACRO Then
            If Len(fo.SName) = 0 Then fo.SName = nm
        End If
    End If
End Sub

' one (var init) binding of let / let* / letrec
Private Sub BindingInit(ByRef b As Variant, ByVal evalEnv As LispObj, ByVal bindEnv As LispObj, ByRef form As Variant)
    Dim p As LispObj, s As LispObj, v As Variant
    If LIsSym(b) Then
        Set s = b
        Set v = lgVoid
    Else
        If Not LIsPair(b) Then SynErr form
        Set p = b
        Set s = NeedSym(p.Car, form)
        If LIsPair(p.Cdr) Then
            LEval LCar(p.Cdr), evalEnv, v
            NameProc v, s.SName
        Else
            Set v = lgVoid
        End If
    End If
    bindEnv.SetRef s, s.Id, v
End Sub

Private Function NamedLet(ByVal o As LispObj, ByVal env As LispObj, ByRef out As Variant, ByRef nx As Variant, ByRef nenv As LispObj) As Boolean
    Dim nm As LispObj, x As Variant, p As LispObj, b As LispObj, args() As Variant, na As Long, e As Variant
    Dim params As Variant, ptail As LispObj, cell As LispObj, loopEnv As LispObj, fo As LispObj, v As Variant, fv As Variant
    Set nm = Nth(o, 1)
    ReDim args(0 To 7)
    LSetV x, Nth(o, 2)
    Do While LIsPair(x)
        Set p = x
        If Not LIsPair(p.Car) Then SynErr o
        Set b = p.Car
        If na > UBound(args) Then
            ReDim Preserve args(0 To 2 * na + 1)
        End If
        LEval LCar(b.Cdr), env, v
        If IsObject(v) Then
            Set args(na) = v
        Else
            args(na) = v
        End If
        Set cell = LCons(NeedSym(b.Car, o), e)
        If ptail Is Nothing Then
            Set params = cell
        Else
            Set ptail.Cdr = cell
        End If
        Set ptail = cell
        na = na + 1
        LSetV x, p.Cdr
    Loop
    Set loopEnv = LNewEnv(env)
    Set fo = LMakeLambda(params, NthCdr(o, 3), loopEnv, nm.SName)
    Set fv = fo
    loopEnv.SetRef nm, nm.Id, fv
    Set nenv = BindArgs(fo, args, na)
    NamedLet = TailSeq(fo.Cdr, nenv, out, nx)
End Function

Private Function CaseMatches(ByRef datums As Variant, ByRef key As Variant) As Boolean
    Dim x As Variant, p As LispObj
    If IsObject(datums) Then
        If datums Is lsElse Then
            CaseMatches = True
            Exit Function
        End If
    End If
    LSetV x, datums
    Do While LIsPair(x)
        Set p = x
        If LEqv(p.Car, key) Then
            CaseMatches = True
            Exit Function
        End If
        LSetV x, p.Cdr
    Loop
End Function

Private Function EvalDo(ByVal o As LispObj, ByVal env As LispObj, ByRef out As Variant, ByRef nx As Variant, ByRef nenv As LispObj) As Boolean
    ' (do ((var init step)...) (test res...) body...)
    Dim specs() As Variant, ns As Long, x As Variant, p As LispObj, sp As LispObj, ne As LispObj, ne2 As LispObj
    Dim v As Variant, t As Variant, s As LispObj, i As Long, vals() As Variant, tst As LispObj, body As Variant
    If LLength(o) < 3 Then SynErr o
    ReDim specs(0 To 7)
    Set ne = LNewEnv(env)
    LSetV x, Nth(o, 1)
    Do While LIsPair(x)
        Set p = x
        If Not LIsPair(p.Car) Then SynErr o
        Set sp = p.Car
        If ns > UBound(specs) Then
            ReDim Preserve specs(0 To 2 * ns + 1)
        End If
        Set specs(ns) = sp
        ns = ns + 1
        Set s = NeedSym(sp.Car, o)
        LEval Nth(sp, 1), env, v
        ne.SetRef s, s.Id, v
        LSetV x, p.Cdr
    Loop
    If Not LIsPair(Nth(o, 2)) Then SynErr o
    Set tst = Nth(o, 2)
    LSetV body, NthCdr(o, 3)
    ReDim vals(0 To IIf(ns > 0, ns - 1, 0))
    Do
        LEval tst.Car, ne, t
        If LTruthy(t) Then
            Set nenv = ne
            EvalDo = TailSeq(tst.Cdr, ne, out, nx)
            Exit Function
        End If
        LEvalSeq body, ne, v
        For i = 0 To ns - 1
            Set sp = specs(i)
            If LLength(sp) >= 3 Then
                LEval Nth(sp, 2), ne, v
            Else
                LLookup ne, sp.Car, v
            End If
            If IsObject(v) Then
                Set vals(i) = v
            Else
                vals(i) = v
            End If
        Next
        Set ne2 = LNewEnv(env)
        For i = 0 To ns - 1
            Set sp = specs(i)
            Set s = sp.Car
            ne2.SetRef s, s.Id, vals(i)
        Next
        Set ne = ne2
    Loop
End Function

Private Function SpecialForm2(ByVal code As Long, ByVal o As LispObj, ByVal n As Long, ByVal env As LispObj, _
                              ByRef out As Variant, ByRef nx As Variant, ByRef nenv As LispObj) As Boolean
    Dim t As Variant, v As Variant, x As Variant, p As LispObj, s As LispObj, m As LispObj, e As Variant
    Dim ne As LispObj, vals() As Variant, nv As Long, b As LispObj, fo As LispObj
    Select Case code
    Case LSF_QUASIQUOTE
        If n <> 2 Then SynErr o
        Quasi Nth(o, 1), 1, env, out
    Case LSF_UNQUOTE, LSF_UNQUOTESPLICING
        LThrow2 "unquote outside of quasiquote:", o
    Case LSF_DELAY, LSF_DELAYFORCE
        If n <> 2 Then SynErr o
        Set p = LNewObj(LK_PROMISE)
        p.SetCar Nth(o, 1)
        Set p.Env = env
        If code = LSF_DELAYFORCE Then p.I1 = 1
        If Not (env Is lgGlobal) Then LKeep env
        Set out = p
    Case LSF_CONSSTREAM
        If n <> 3 Then SynErr o
        LEval Nth(o, 1), env, v
        Set p = LNewObj(LK_PROMISE)
        p.SetCar Nth(o, 2)
        Set p.Env = env
        If Not (env Is lgGlobal) Then LKeep env
        Set t = p
        Set out = LCons(v, t)
    Case LSF_SYNTAXRULES
        Set out = MakeSyntaxRules(o, env)
    Case LSF_DEFSYNTAX
        If n <> 3 Then SynErr o
        Set s = NeedSym(Nth(o, 1), o)
        LEval Nth(o, 2), env, v
        If Not LIsKind(v, LK_MACRO) Then
            If LIsProc(v) Then
                ' a transformer procedure taking the whole form
                Set m = LNewObj(LK_MACRO)
                m.Flags = 3
                Set m.Ref = v
                Set v = m
            Else
                LThrow2 "define-syntax: expected syntax-rules, got", v
            End If
        End If
        NameProc v, s.SName
        env.SetRef s, s.Id, v
        Set out = lgVoid
    Case LSF_DEFMACRO
        ' (define-macro (name . params) body...) | (define-macro name transformer)
        ' (defmacro name params body...)
        If n < 3 Then SynErr o
        LSetV t, Nth(o, 1)
        Set m = LNewObj(LK_MACRO)
        m.Flags = 2
        If LIsPair(t) Then
            Set p = t
            Set s = NeedSym(p.Car, o)
            Set m.Ref = LMakeLambda(p.Cdr, NthCdr(o, 2), env, s.SName)
        Else
            Set s = NeedSym(t, o)
            If n = 3 Then
                LEval Nth(o, 2), env, v
                If Not LIsProc(v) Then LThrow2 "define-macro: expected a procedure, got", v
                Set m.Ref = v
            Else
                Set m.Ref = LMakeLambda(Nth(o, 2), NthCdr(o, 3), env, s.SName)
            End If
        End If
        m.SName = s.SName
        Set v = m
        env.SetRef s, s.Id, v
        Set out = lgVoid
    Case LSF_GUARD
        EvalGuard o, env, out
    Case LSF_DEFRECORD
        DefineRecord o, env
        Set out = lgVoid
    Case LSF_ASSERT
        If n < 2 Then SynErr o
        LEval Nth(o, 1), env, t
        If Not LTruthy(t) Then LThrow2 "assertion failed:", Nth(o, 1)
        Set out = lgVoid
    Case LSF_IGNORE
        Set out = lgVoid
    Case LSF_CASELAMBDA
        Set fo = LNewObj(LK_LAMBDA)
        fo.Flags = 4
        LSetV x, o.Cdr
        Do While LIsPair(x)
            Set p = x
            If Not LIsPair(p.Car) Then SynErr o
            Set b = p.Car
            fo.Push ToV(LMakeLambda(b.Car, b.Cdr, env, vbNullString))
            LSetV x, p.Cdr
        Loop
        Set out = fo
    Case LSF_LETVALUES
        If n < 3 Then SynErr o
        Set ne = LNewEnv(env)
        LSetV x, Nth(o, 1)
        Do While LIsPair(x)
            Set p = x
            If Not LIsPair(p.Car) Then SynErr o
            Set b = p.Car
            LEval LCar(b.Cdr), ne, v
            ValuesToArray v, vals, nv
            BindFormals b.Car, vals, nv, ne
            LSetV x, p.Cdr
        Loop
        Set nenv = ne
        SpecialForm2 = TailSeq(NthCdr(o, 2), ne, out, nx)
    Case LSF_DEFVALUES
        If n <> 3 Then SynErr o
        LEval Nth(o, 2), env, v
        ValuesToArray v, vals, nv
        BindFormals Nth(o, 1), vals, nv, env
        Set out = lgVoid
    Case Else
        LThrow2 "unknown special form:", o
    End Select
End Function

Private Function ToV(ByVal o As LispObj) As Variant
    Set ToV = o
End Function

' ---------------------------------------------------------------------------
'  quasiquote
' ---------------------------------------------------------------------------
Private Sub Quasi(ByRef x As Variant, ByVal depth As Long, ByVal env As LispObj, ByRef out As Variant)
    Dim o As LispObj, hd As Variant, h As LispObj, a As Variant, d As Variant, sp As Variant, e As Variant
    Dim vec As LispObj, lst As Variant, i As Long, item As Variant, r As LispObj, y As Variant, q As LispObj
    If LIsKind(x, LK_VECTOR) Then
        Set vec = x
        lst = Empty
        For i = vec.Count - 1 To 0 Step -1
            vec.GetInto i, item
            Set lst = LCons(item, lst)
        Next
        Quasi lst, depth, env, y
        Set r = LNewObj(LK_VECTOR)
        Do While LIsPair(y)
            Set q = y
            r.Push q.Car
            LSetV y, q.Cdr
        Loop
        Set out = r
        Exit Sub
    End If
    If Not LIsPair(x) Then
        LSetV out, x
        Exit Sub
    End If
    Set o = x
    LSetV hd, o.Car
    If IsObject(hd) Then
        If hd Is lsUnquote Then
            If depth = 1 Then
                LEval Nth(o, 1), env, out
            Else
                Quasi Nth(o, 1), depth - 1, env, a
                Set out = LList2(lsUnquote, a)
            End If
            Exit Sub
        End If
        If hd Is lsQuasi Then
            Quasi Nth(o, 1), depth + 1, env, a
            Set out = LList2(lsQuasi, a)
            Exit Sub
        End If
        If LIsPair(hd) Then
            Set h = hd
            If IsObject(h.Car) Then
                If h.Car Is lsUnquoteSplicing And depth = 1 Then
                    LEval Nth(h, 1), env, sp
                    Quasi o.Cdr, depth, env, d
                    LSetV out, LAppend2(sp, d)
                    Exit Sub
                End If
            End If
        End If
    End If
    Quasi o.Car, depth, env, a
    Quasi o.Cdr, depth, env, d
    Set out = LCons(a, d)
End Sub

' copy of list a followed by b
Public Function LAppend2(ByRef a As Variant, ByRef b As Variant) As Variant
    Dim head As LispObj, tail As LispObj, cell As LispObj, x As Variant, p As LispObj, e As Variant
    LSetV x, a
    Do While LIsPair(x)
        Set p = x
        Set cell = LCons(p.Car, e)
        If head Is Nothing Then
            Set head = cell
        Else
            Set tail.Cdr = cell
        End If
        Set tail = cell
        LSetV x, p.Cdr
    Loop
    If Not IsEmpty(x) Then LThrow2 "append: expected a proper list, got", a
    If head Is Nothing Then
        If IsObject(b) Then Set LAppend2 = b Else LAppend2 = b
    Else
        tail.SetCdr b
        Set LAppend2 = head
    End If
End Function

' ---------------------------------------------------------------------------
'  macros
' ---------------------------------------------------------------------------
Private Function MakeSyntaxRules(ByVal o As LispObj, ByVal env As LispObj) As LispObj
    ' (syntax-rules (literal...) (pattern template)...)
    ' (syntax-rules ellipsis (literal...) (pattern template)...)
    Dim m As LispObj, x As Variant
    Set m = LNewObj(LK_MACRO)
    m.Flags = 1
    Set m.Env = env
    Set m.Ref = lsEllipsis
    LSetV x, o.Cdr
    If LIsSym(LCar(x)) Then
        Set m.Ref = LCar(x)
        LSetV x, LCdr(x)
    End If
    If Not (IsEmpty(LCar(x)) Or LIsPair(LCar(x))) Then SynErr o
    m.SetCar LCar(x)
    m.SetCdr LCdr(x)
    If Not (env Is lgGlobal) Then LKeep env
    Set MakeSyntaxRules = m
End Function

Private Sub ExpandMacro(ByVal m As LispObj, ByVal form As LispObj, ByVal env As LispObj, ByRef out As Variant)
    Dim args() As Variant, n As Long, x As Variant, rule As LispObj, b As LispObj, pat As Variant, ren As LispObj
    Dim fv As Variant, fm As Variant
    Select Case m.Flags
    Case 2
        LListToArray form.Cdr, args, n, m.SName
        Set fv = m.Ref
        LApply fv, args, n, out
    Case 3
        Set fv = m.Ref
        Set fm = form
        LApply1 fv, fm, out
    Case Else
        LSetV x, m.Cdr
        Do While LIsPair(x)
            Set rule = LCar(x)
            LSetV pat, rule.Car
            Set b = LNewEnv(Nothing)
            If LIsPair(pat) Then
                If SrMatch(LCdr(pat), form.Cdr, m, b) Then
                    Set ren = LNewEnv(Nothing)
                    CollectBinders LCar(rule.Cdr), b, m, ren
                    SrExpand LCar(rule.Cdr), b, m, ren, out
                    Exit Sub
                End If
            End If
            LSetV x, LCdr(x)
        Loop
        LThrow2 "no syntax rule matches", form
    End Select
End Sub

Private Function IsLiteral(ByVal s As LispObj, ByVal m As LispObj) As Boolean
    Dim x As Variant, p As LispObj
    LSetV x, m.Car
    Do While LIsPair(x)
        Set p = x
        If IsObject(p.Car) Then
            If p.Car Is s Then
                IsLiteral = True
                Exit Function
            End If
        End If
        LSetV x, p.Cdr
    Loop
End Function

' is the element after the head of list o the ellipsis?
Private Function EllipsisNext(ByVal o As LispObj, ByVal m As LispObj) As Boolean
    Dim q As LispObj
    If LIsPair(o.Cdr) Then
        Set q = o.Cdr
        If IsObject(q.Car) Then EllipsisNext = (q.Car Is m.Ref)
    End If
End Function

Private Sub PatternVars(ByRef pat As Variant, ByVal m As LispObj, ByVal acc As LispObj)
    Dim p As LispObj, s As LispObj, i As Long, x As Variant
    If LIsSym(pat) Then
        Set s = pat
        If s Is m.Ref Or s Is lsUnderscore Then Exit Sub
        If IsLiteral(s, m) Then Exit Sub
        acc.Push pat
    ElseIf LIsPair(pat) Then
        Set p = pat
        PatternVars p.Car, m, acc
        PatternVars p.Cdr, m, acc
    ElseIf LIsKind(pat, LK_VECTOR) Then
        Set p = pat
        For i = 0 To p.Count - 1
            p.GetInto i, x
            PatternVars x, m, acc
        Next
    End If
End Sub

Private Function VecToList(ByVal vec As LispObj) As Variant
    Dim i As Long, item As Variant, lst As Variant
    For i = vec.Count - 1 To 0 Step -1
        vec.GetInto i, item
        Set lst = LCons(item, lst)
    Next
    If IsObject(lst) Then Set VecToList = lst Else VecToList = lst
End Function

Private Function SrMatch(ByRef pat As Variant, ByRef form As Variant, ByVal m As LispObj, ByVal b As LispObj) As Boolean
    Dim po As LispObj, fo As LispObj, p As Variant, tailPat As Variant, tailLen As Long, x As Variant, q As LispObj
    Dim items() As Variant, cnt As Long, ftail As Variant, reps As Long, vars As LispObj, multis() As LispObj
    Dim k As Long, j As Long, bk As LispObj, sv As Variant, s As LispObj, v As Variant, remn As Variant, mv As Variant, mo As LispObj
    If LIsSym(pat) Then
        Set po = pat
        If po Is lsUnderscore Then
            SrMatch = True
        ElseIf IsLiteral(po, m) Then
            If IsObject(form) Then SrMatch = (form Is po)
        Else
            b.SetRef po, po.Id, form
            SrMatch = True
        End If
        Exit Function
    End If
    If LIsPair(pat) Then
        Set po = pat
        If EllipsisNext(po, m) Then
            LSetV p, po.Car
            LSetV tailPat, NthCdr(po, 2)
            tailLen = 0
            LSetV x, tailPat
            Do While LIsPair(x)
                tailLen = tailLen + 1
                LSetV x, LCdr(x)
            Loop
            ReDim items(0 To 7)
            cnt = 0
            LSetV x, form
            Do While LIsPair(x)
                Set q = x
                If cnt > UBound(items) Then
                    ReDim Preserve items(0 To 2 * cnt + 1)
                End If
                If IsObject(q.Car) Then
                    Set items(cnt) = q.Car
                Else
                    items(cnt) = q.Car
                End If
                cnt = cnt + 1
                LSetV x, q.Cdr
            Loop
            LSetV ftail, x
            reps = cnt - tailLen
            If reps < 0 Then Exit Function
            Set vars = LNewObj(LK_VECTOR)
            PatternVars p, m, vars
            If vars.Count > 0 Then
                ReDim multis(0 To vars.Count - 1)
                For j = 0 To vars.Count - 1
                    Set multis(j) = LNewObj(LK_MULTI)
                Next
            End If
            For k = 0 To reps - 1
                Set bk = LNewEnv(Nothing)
                If Not SrMatch(p, items(k), m, bk) Then Exit Function
                For j = 0 To vars.Count - 1
                    vars.GetInto j, sv
                    Set s = sv
                    v = Empty
                    TryLookup bk, s, v
                    Set mo = multis(j)
                    mo.Push v
                Next
            Next
            For j = 0 To vars.Count - 1
                vars.GetInto j, sv
                Set s = sv
                Set mv = multis(j)
                b.SetRef s, s.Id, mv
            Next
            LSetV remn, ftail
            For k = cnt - 1 To reps Step -1
                Set remn = LCons(items(k), remn)
            Next
            SrMatch = SrMatch(tailPat, remn, m, b)
            Exit Function
        End If
        If Not LIsPair(form) Then Exit Function
        Set fo = form
        If Not SrMatch(po.Car, fo.Car, m, b) Then Exit Function
        SrMatch = SrMatch(po.Cdr, fo.Cdr, m, b)
        Exit Function
    End If
    If IsEmpty(pat) Then
        SrMatch = IsEmpty(form)
        Exit Function
    End If
    If LIsKind(pat, LK_VECTOR) Then
        If Not LIsKind(form, LK_VECTOR) Then Exit Function
        SrMatch = SrMatch(VecToList(pat), VecToList(form), m, b)
        Exit Function
    End If
    SrMatch = LEqual(pat, form)
End Function

Private Sub MultiVars(ByRef t As Variant, ByVal b As LispObj, ByVal m As LispObj, ByVal acc As LispObj)
    ' symbols in template t that are bound to ellipsis matches
    Dim p As LispObj, v As Variant, i As Long, x As Variant
    If LIsSym(t) Then
        Set p = t
        If TryLookup(b, p, v) Then
            If LIsKind(v, LK_MULTI) Then acc.Push t
        End If
    ElseIf LIsPair(t) Then
        Set p = t
        MultiVars p.Car, b, m, acc
        MultiVars p.Cdr, b, m, acc
    ElseIf LIsKind(t, LK_VECTOR) Then
        Set p = t
        For i = 0 To p.Count - 1
            p.GetInto i, x
            MultiVars x, b, m, acc
        Next
    End If
End Sub

' Record in ren the identifiers that the template binds with let / let* /
' letrec / named let / lambda / do (pattern variables excluded).  They are
' renamed on expansion, so they cannot capture variables of the macro user.
Private Sub AddBinder(ByRef v As Variant, ByVal b As LispObj, ByVal m As LispObj, ByVal ren As LispObj)
    Dim s As LispObj, x As Variant
    If Not LIsSym(v) Then Exit Sub
    Set s = v
    If s Is m.Ref Or s Is lsUnderscore Or s.Code <> 0 Then Exit Sub
    If TryLookup(b, s, x) Then Exit Sub
    If ren.FindRef(s, s.Id) >= 0 Then Exit Sub
    x = True
    ren.SetRef s, s.Id, x
End Sub

Private Sub CollectBinders(ByRef t As Variant, ByVal b As LispObj, ByVal m As LispObj, ByVal ren As LispObj)
    Dim o As LispObj, hs As LispObj, x As Variant, p As LispObj, bd As Variant, q As LispObj
    If Not LIsPair(t) Then Exit Sub
    Set o = t
    If LIsSym(o.Car) Then
        Set hs = o.Car
        Select Case hs.SName
        Case "let", "let*", "letrec", "letrec*", "do", "let-values", "let*-values"
            LSetV x, o.Cdr
            If LIsPair(x) Then
                Set p = x
                If LIsSym(p.Car) Then
                    AddBinder p.Car, b, m, ren
                    LSetV x, p.Cdr
                End If
            End If
            If LIsPair(x) Then
                Set p = x
                LSetV bd, p.Car
                Do While LIsPair(bd)
                    Set q = bd
                    If LIsPair(q.Car) Then
                        AddBinder LCar(q.Car), b, m, ren
                    Else
                        AddBinder q.Car, b, m, ren
                    End If
                    LSetV bd, q.Cdr
                Loop
            End If
        Case "lambda"
            If LIsPair(o.Cdr) Then
                LSetV bd, LCar(o.Cdr)
                Do While LIsPair(bd)
                    Set q = bd
                    AddBinder q.Car, b, m, ren
                    LSetV bd, q.Cdr
                Loop
                AddBinder bd, b, m, ren
            End If
        End Select
    End If
    LSetV x, t
    Do While LIsPair(x)
        Set p = x
        CollectBinders p.Car, b, m, ren
        LSetV x, p.Cdr
    Loop
End Sub

Private Sub SrExpand(ByRef t As Variant, ByVal b As LispObj, ByVal m As LispObj, ByVal ren As LispObj, ByRef out As Variant)
    Dim po As LispObj, s As LispObj, v As Variant, vars As LispObj, k As Long, j As Long, n As Long, bk As LispObj
    Dim sv As Variant, mo As LispObj, res As Variant, head As LispObj, tail As LispObj, cell As LispObj, e As Variant
    Dim a As Variant, d As Variant, item As Variant, after As Variant, fresh As LispObj, vec As LispObj, y As Variant
    If LIsSym(t) Then
        Set s = t
        If TryLookup(b, s, v) Then
            If LIsKind(v, LK_MULTI) Then LThrow2 "syntax-rules: pattern variable used without an ellipsis:", s
            LSetV out, v
            Exit Sub
        End If
        ' variables bound by the template itself get fresh names (hygiene)
        If TryLookup(ren, s, v) Then
            If Not LIsSym(v) Then
                Set fresh = LNewObj(LK_SYM)
                fresh.SName = s.SName
                fresh.H = s.H
                Set v = fresh
                ren.SetRef s, s.Id, v
            End If
            LSetV out, v
            Exit Sub
        End If
        Set out = s
        Exit Sub
    End If
    If LIsPair(t) Then
        Set po = t
        If IsObject(po.Car) Then
            If po.Car Is m.Ref And LIsPair(po.Cdr) Then
                ' (... ...) escapes the ellipsis
                LSetV out, LCar(po.Cdr)
                Exit Sub
            End If
        End If
        If EllipsisNext(po, m) Then
            Set vars = LNewObj(LK_VECTOR)
            MultiVars po.Car, b, m, vars
            If vars.Count = 0 Then LThrow2 "syntax-rules: no pattern variable before the ellipsis in", t
            n = -1
            For j = 0 To vars.Count - 1
                vars.GetInto j, sv
                Set s = sv
                TryLookup b, s, v
                Set mo = v
                If n < 0 Or mo.Count < n Then n = mo.Count
            Next
            For k = 0 To n - 1
                Set bk = LNewEnv(b)
                For j = 0 To vars.Count - 1
                    vars.GetInto j, sv
                    Set s = sv
                    TryLookup b, s, v
                    Set mo = v
                    mo.GetInto k, item
                    bk.SetRef s, s.Id, item
                Next
                SrExpand po.Car, bk, m, ren, item
                Set cell = LCons(item, e)
                If head Is Nothing Then
                    Set head = cell
                Else
                    Set tail.Cdr = cell
                End If
                Set tail = cell
            Next
            LSetV after, NthCdr(po, 2)
            SrExpand after, b, m, ren, d
            If head Is Nothing Then
                LSetV out, d
            Else
                tail.SetCdr d
                Set out = head
            End If
            Exit Sub
        End If
        SrExpand po.Car, b, m, ren, a
        SrExpand po.Cdr, b, m, ren, d
        Set out = LCons(a, d)
        Exit Sub
    End If
    If LIsKind(t, LK_VECTOR) Then
        Set po = t
        SrExpand VecToList(po), b, m, ren, y
        Set vec = LNewObj(LK_VECTOR)
        Do While LIsPair(y)
            Set cell = y
            vec.Push cell.Car
            LSetV y, cell.Cdr
        Loop
        Set out = vec
        Exit Sub
    End If
    LSetV out, t
End Sub

' ---------------------------------------------------------------------------
'  records:  (define-record-type <point> (make-point x y) point? (x point-x set-point-x!) ...)
' ---------------------------------------------------------------------------
Private Function RecProc(ByVal code As Long, ByVal rt As LispObj, ByVal nm As String, ByVal idx As Long) As LispObj
    Dim p As LispObj
    Set p = LNewObj(LK_RECPROC)
    p.Code = code
    Set p.Ref = rt
    p.SName = nm
    p.I1 = idx
    Set RecProc = p
End Function

Private Function FieldIndex(ByVal rt As LispObj, ByVal s As LispObj) As Long
    Dim i As Long, v As Variant
    For i = 0 To rt.Count - 1
        rt.GetInto i, v
        If v Is s Then
            FieldIndex = i
            Exit Function
        End If
    Next
    LThrow2 "define-record-type: unknown field", s
End Function

Private Sub DefineRecord(ByVal o As LispObj, ByVal env As LispObj)
    Dim rt As LispObj, nm As String, x As Variant, p As LispObj, f As LispObj, s As LispObj, ctor As Variant
    Dim cp As LispObj, pr As LispObj, v As Variant, y As Variant, q As LispObj, f2 As Long
    If LLength(o) < 3 Then SynErr o
    nm = LToStr(Nth(o, 1), False)
    If Left$(nm, 1) = "<" And Right$(nm, 1) = ">" And Len(nm) > 2 Then nm = Mid$(nm, 2, Len(nm) - 2)
    Set rt = LNewObj(LK_RECTYPE)
    rt.SName = nm
    ' fields
    LSetV x, NthCdr(o, 4)
    Do While LIsPair(x)
        Set p = x
        If LIsSym(p.Car) Then
            rt.Push p.Car
        Else
            rt.Push LCar(p.Car)
        End If
        LSetV x, p.Cdr
    Loop
    If LIsSym(Nth(o, 1)) Then
        Set s = Nth(o, 1)
        Set v = rt
        env.SetRef s, s.Id, v
    End If
    ' constructor
    LSetV ctor, Nth(o, 2)
    If LIsPair(ctor) Then
        Set cp = ctor
        Set s = NeedSym(cp.Car, o)
        Set pr = RecProc(1, rt, s.SName, 0)
        LSetV y, cp.Cdr
        Do While LIsPair(y)
            Set q = y
            pr.Push CLng(FieldIndex(rt, NeedSym(q.Car, o)))
            LSetV y, q.Cdr
        Loop
        pr.I1 = pr.Count
        Set v = pr
        env.SetRef s, s.Id, v
    ElseIf LIsSym(ctor) Then
        Set s = ctor
        Set pr = RecProc(1, rt, s.SName, 0)
        For f2 = 0 To rt.Count - 1
            pr.Push f2
        Next
        pr.I1 = pr.Count
        Set v = pr
        env.SetRef s, s.Id, v
    End If
    ' predicate
    If LIsSym(Nth(o, 3)) Then
        Set s = Nth(o, 3)
        Set v = RecProc(2, rt, s.SName, 0)
        env.SetRef s, s.Id, v
    End If
    ' accessors and modifiers
    LSetV x, NthCdr(o, 4)
    Do While LIsPair(x)
        Set p = x
        If LIsPair(p.Car) Then
            Set f = p.Car
            If LIsSym(Nth(f, 1)) Then
                Set s = Nth(f, 1)
                Set v = RecProc(3, rt, s.SName, FieldIndex(rt, f.Car))
                env.SetRef s, s.Id, v
            End If
            If LIsSym(Nth(f, 2)) Then
                Set s = Nth(f, 2)
                Set v = RecProc(4, rt, s.SName, FieldIndex(rt, f.Car))
                env.SetRef s, s.Id, v
            End If
        End If
        LSetV x, p.Cdr
    Loop
End Sub

' ---------------------------------------------------------------------------
'  errors, continuations, dynamic-wind, promises
' ---------------------------------------------------------------------------
Private Sub EvalGuard(ByVal o As LispObj, ByVal env As LispObj, ByRef out As Variant)
    ' (guard (var clause...) body...)
    Dim spec As LispObj, var As LispObj, cond As Variant, saveDepth As Long, ne As LispObj
    Dim x As Variant, cl As LispObj, p As LispObj, t As Variant, f As Variant
    If Not LIsPair(Nth(o, 1)) Then SynErr o
    Set spec = Nth(o, 1)
    Set var = NeedSym(spec.Car, o)
    saveDepth = lgDepth
    On Error GoTo EH
    LEvalSeq NthCdr(o, 2), env, out
    Exit Sub
EH:
    If Err.Number = LESCAPE Then Err.Raise LESCAPE
    LSetV cond, LCapture()
    lgDepth = saveDepth
    Resume Handle
Handle:
    On Error GoTo 0
    Set ne = LNewEnv(env)
    ne.SetRef var, var.Id, cond
    LSetV x, spec.Cdr
    Do While LIsPair(x)
        Set p = x
        If Not LIsPair(p.Car) Then SynErr o
        Set cl = p.Car
        If IsObject(cl.Car) Then
            If cl.Car Is lsElse Then
                LEvalSeq cl.Cdr, ne, out
                Exit Sub
            End If
        End If
        LEval cl.Car, ne, t
        If LTruthy(t) Then
            If IsEmpty(cl.Cdr) Then
                LSetV out, t
            ElseIf IsArrowClause(cl) Then
                LEval Nth(cl, 2), ne, f
                LApply1 f, t, out
            Else
                LEvalSeq cl.Cdr, ne, out
            End If
            Exit Sub
        End If
        LSetV x, p.Cdr
    Loop
    LRaise cond
End Sub

' (with-exception-handler handler thunk): calls (handler condition) when thunk raises
Public Sub LWithHandler(ByRef handler As Variant, ByRef thunk As Variant, ByRef out As Variant)
    Dim cond As Variant, saveDepth As Long
    saveDepth = lgDepth
    On Error GoTo EH
    LApply0 thunk, out
    Exit Sub
EH:
    If Err.Number = LESCAPE Then Err.Raise LESCAPE
    LSetV cond, LCapture()
    lgDepth = saveDepth
    Resume Handle
Handle:
    On Error GoTo 0
    LApply1 handler, cond, out
End Sub

Public Sub LCallCC(ByRef f As Variant, ByRef out As Variant)
    Dim k As LispObj, saveDepth As Long, kv As Variant, num As Long, desc As String
    Set k = LNewObj(LK_CONT)
    Set kv = k
    saveDepth = lgDepth
    On Error GoTo EH
    LApply1 f, kv, out
    k.Flags = 1
    Exit Sub
EH:
    k.Flags = 1
    num = Err.Number
    If num = LESCAPE Then
        If lgEscK Is k Then
            LSetV out, lgEscVal
            Set lgEscK = Nothing
            lgEscVal = Empty
            lgDepth = saveDepth
            Exit Sub
        End If
    End If
    desc = Err.Description
    Err.Raise num, "LispVBA", desc
End Sub

Public Sub LInvokeCont(ByVal k As LispObj, ByRef args() As Variant, ByVal na As Long)
    Dim v As LispObj, i As Long
    If k.Flags = 1 Then LThrow "continuation invoked after its extent ended (only escaping continuations are supported)"
    If na = 1 Then
        LSetV lgEscVal, args(0)
    Else
        Set v = LNewObj(LK_VALUES)
        For i = 0 To na - 1
            v.Push args(i)
        Next
        Set lgEscVal = v
    End If
    Set lgEscK = k
    Err.Raise LESCAPE, "LispVBA", "continuation"
End Sub

Public Sub LDynamicWind(ByRef before As Variant, ByRef thunk As Variant, ByRef after As Variant, ByRef out As Variant)
    Dim r As Variant, num As Long, desc As String, saveErr As Variant, saveK As LispObj, saveV As Variant
    LApply0 before, r
    On Error GoTo EH
    LApply0 thunk, out
    On Error GoTo 0
    LApply0 after, r
    Exit Sub
EH:
    num = Err.Number
    desc = Err.Description
    LSetV saveErr, lgErr
    Set saveK = lgEscK
    LSetV saveV, lgEscVal
    Resume Unwind
Unwind:
    On Error GoTo 0
    LApply0 after, r
    LSetV lgErr, saveErr
    Set lgEscK = saveK
    LSetV lgEscVal, saveV
    Err.Raise num, "LispVBA", desc
End Sub

Public Sub LForce(ByRef pv As Variant, ByRef out As Variant)
    Dim p As LispObj, v As Variant, q As LispObj
    If Not LIsKind(pv, LK_PROMISE) Then
        LSetV out, pv
        Exit Sub
    End If
    Set p = pv
    Do
        If p.Flags = 1 Then
            LSetV out, p.Cdr
            Exit Sub
        End If
        LEval p.Car, p.Env, v
        If p.Flags = 1 Then
            LSetV out, p.Cdr
            Exit Sub
        End If
        If p.I1 = 1 And LIsKind(v, LK_PROMISE) Then
            ' delay-force: continue with the inner promise
            Set q = v
            If q.Flags = 1 Then
                p.SetCdr q.Cdr
                p.Flags = 1
            Else
                p.SetCar q.Car
                Set p.Env = q.Env
                p.I1 = q.I1
            End If
        Else
            p.SetCdr v
            p.Flags = 1
            p.Car = Empty
            Set p.Env = Nothing
        End If
    Loop
End Sub
