Attribute VB_Name = "PyEval"
Option Explicit
' ===========================================================================
'  PyEval - tree-walking evaluator: statements, expressions, function calls,
'  scopes, classes, exceptions and the public entry points.
'
'  Entry points:
'    PyRunProgram(src)  -> runs a program, returns everything it printed
'                          (plus a traceback if it failed)
'    PyEvalCode(...)    -> used by the PY() worksheet function
' ===========================================================================

Public Const PY_DEFAULT_MAX_DEPTH As Long = 300

' ---------------------------------------------------------------------------
'  entry points
' ---------------------------------------------------------------------------
Public Sub PyReset()
    ReleaseAll
    gNextId = 0
    If gMaxDepth <= 0 Then gMaxDepth = PY_DEFAULT_MAX_DEPTH
    ReDim gFrameName(0 To gMaxDepth + 2)
    ReDim gFrameLine(0 To gMaxDepth + 2)
    gDepth = 0
    gFrameName(0) = "<module>"
    gRet = Empty
    Set gExc = Nothing
    Set gCurExc = Nothing
    OutReset
    gDecSep = Mid$(Format$(1.5, "0.0"), 2, 1)
    InitBuiltins
    Set gGlobals = NewObj(OK_ENV)
    NameSet gGlobals, "__name__", "__main__"
    Set gModules = NewObj(OK_ENV)
End Sub

' free the interpreter state of the last run
Public Sub PyFinish()
    ReleaseAll
    Set gBuiltins = Nothing
End Sub

' run a whole program; returns the printed output (and the traceback, if any)
Public Function PyRunProgram(ByVal src As String) As String
    Dim prog As PyNode, exc As PyObj, msg As String, st As Long
    PyReset
    On Error GoTo EH
    Set prog = PyParse(src)
    st = ExecBlock(prog, gGlobals)
    On Error GoTo 0
    PyRunProgram = OutText()
    PyFinish
    Exit Function
EH:
    Set exc = CaptureExc()
    Resume Report
Report:
    On Error GoTo EH2
    msg = FormatException(exc)
    On Error GoTo 0
    PyRunProgram = OutText() & msg
    PyFinish
    Exit Function
EH2:
    msg = exc.Cls.PName & " (error while formatting the exception)" & vbLf
    Resume Report2
Report2:
    On Error GoTo 0
    PyRunProgram = OutText() & msg
    PyFinish
End Function

' Evaluate code for a worksheet formula.  The arguments are bound to x, y, z
' and the list args.  A pure expression returns its value; a program returns
' the variable "result" if it defines one (hasResult=True).
' Returns False on error (errText = "ExcType: message").
' The caller must call PyFinish after converting the result.
Public Function PyEvalCode(ByVal src As String, ByRef argv() As Variant, ByVal nArgs As Long, _
                           ByRef result As Variant, ByRef hasResult As Boolean, _
                           ByRef printed As String, ByRef errText As String) As Boolean
    Dim node As PyNode, lst As PyObj, i As Long, exc As PyObj, st As Long, nm As String
    PyReset
    On Error GoTo EH
    Set lst = NewList()
    For i = 0 To nArgs - 1
        lst.Push argv(i)
        If i < 3 Then
            nm = Mid$("xyz", i + 1, 1)
            NameSet gGlobals, nm, argv(i)
        End If
    Next
    NameSet gGlobals, "args", lst
    Set node = PyTryParseExpr(src)
    If node Is Nothing Then
        Set node = PyParse(src)
        st = ExecBlock(node, gGlobals)
        hasResult = NameGet(gGlobals, "result", result)
    Else
        EvalInto node, gGlobals, result
        hasResult = True
    End If
    printed = OutText()
    PyEvalCode = True
    Exit Function
EH:
    Set exc = CaptureExc()
    Resume Rep
Rep:
    On Error Resume Next
    errText = exc.Cls.PName
    errText = ExcLine(exc)
    printed = OutText()
    On Error GoTo 0
End Function

Public Function FormatException(ByVal exc As PyObj) As String
    Dim v As Variant, t As PyObj, x As Variant
    If IsSubclassOf(exc.Cls, ExcClass("SystemExit")) Then
        If NameGet(exc, "args", v) Then
            If IsKind(v, OK_TUPLE) Then
                Set t = v
                If t.Count > 0 Then
                    t.GetInto 0, x
                    If VarType(x) = vbString Then FormatException = x & vbLf
                End If
            End If
        End If
        Exit Function
    End If
    If IsSubclassOf(exc.Cls, ExcClass("SyntaxError")) Then
        FormatException = "  File ""<python>"", line " & exc.I1 & vbLf & ExcLine(exc) & vbLf
        Exit Function
    End If
    FormatException = FormatTraceback(exc)
End Function

Public Function FormatTraceback(ByVal e As PyObj) As String
    Dim i As Long, d As Long, ln As Long, parts() As String, k As Long
    d = e.I2
    If d > gMaxDepth + 1 Then d = gMaxDepth + 1
    If d < 0 Then d = 0
    ReDim parts(0 To d + 3)
    parts(0) = "Traceback (most recent call last):"
    k = 1
    For i = 0 To d
        If d > 24 And i >= 3 And i < d - 10 Then
            If i = 3 Then
                parts(k) = "  ... (" & (d - 13) & " more frames)"
                k = k + 1
            End If
        Else
            If i = d Then
                ln = e.I1
            Else
                ln = gFrameLine(i)
            End If
            parts(k) = "  File ""<python>"", line " & ln & ", in " & gFrameName(i)
            k = k + 1
        End If
    Next
    parts(k) = ExcLine(e)
    ReDim Preserve parts(0 To k)
    FormatTraceback = Join(parts, vbLf) & vbLf
End Function

' ---------------------------------------------------------------------------
'  statements
' ---------------------------------------------------------------------------
Public Function ExecBlock(ByVal blk As PyNode, ByVal env As PyObj) As Long
    Dim i As Long, st As Long
    For i = 0 To blk.N - 1
        st = ExecStmt(blk.Kid(i), env)
        If st <> ST_NORMAL Then
            ExecBlock = st
            Exit Function
        End If
    Next
End Function

Public Function ExecStmt(ByVal nd As PyNode, ByVal env As PyObj) As Long
    Dim v As Variant, i As Long
    gFrameLine(gDepth) = nd.Ln
    Select Case nd.Kind
    Case NK_EXPR
        EvalInto nd.A, env, v
    Case NK_ASSIGN
        EvalInto nd.A, env, v
        For i = 0 To nd.N - 1
            AssignTo nd.Kid(i), v, env
        Next
    Case NK_IF
        EvalInto nd.A, env, v
        If Truthy(v) Then
            ExecStmt = ExecBlock(nd.B, env)
        ElseIf Not (nd.C Is Nothing) Then
            If nd.C.Kind = NK_IF Then
                ExecStmt = ExecStmt(nd.C, env)
            Else
                ExecStmt = ExecBlock(nd.C, env)
            End If
        End If
    Case NK_FOR
        ExecStmt = ExecFor(nd, env)
    Case NK_WHILE
        ExecStmt = ExecWhile(nd, env)
    Case NK_RETURN
        If nd.A Is Nothing Then
            gRet = Empty
        Else
            EvalInto nd.A, env, v
            If IsObject(v) Then Set gRet = v Else gRet = v
        End If
        ExecStmt = ST_RETURN
    Case NK_AUGASSIGN
        ExecAugAssign nd, env
    Case NK_PASS
    Case NK_BREAK
        ExecStmt = ST_BREAK
    Case NK_CONTINUE
        ExecStmt = ST_CONTINUE
    Case NK_DEF
        ExecDef nd, env
    Case NK_CLASS
        ExecClass nd, env
    Case NK_TRY
        ExecStmt = ExecTry(nd, env)
    Case NK_RAISE
        ExecRaise nd, env
    Case NK_GLOBAL, NK_NONLOCAL
        ExecDecl nd, env
    Case NK_DEL
        For i = 0 To nd.N - 1
            DeleteTarget nd.Kid(i), env
        Next
    Case NK_ASSERT
        EvalInto nd.A, env, v
        If Not Truthy(v) Then
            If nd.B Is Nothing Then
                RaiseExc NewExcInst(ExcClass("AssertionError"), NewObj(OK_TUPLE))
            Else
                EvalInto nd.B, env, v
                RaiseExc MakeExcV("AssertionError", v)
            End If
        End If
    Case NK_YIELD
        ExecYield nd, env
    Case NK_ANNASSIGN
        If Not (nd.B Is Nothing) Then
            EvalInto nd.B, env, v
            AssignTo nd.A, v, env
        End If
        If env.Kind = OK_CLASS And nd.A.Kind = NK_NAME Then RecordAnnotation env, nd.A.S
    Case NK_IMPORT
        ExecImport nd, env
    Case NK_FROMIMPORT
        ExecFromImport nd, env
    Case NK_BLOCK
        ExecStmt = ExecBlock(nd, env)
    Case Else
        ThrowErr "SystemError", "unknown statement kind " & nd.Kind
    End Select
End Function

Private Function ExecWhile(ByVal nd As PyNode, ByVal env As PyObj) As Long
    Dim v As Variant, st As Long
    Do
        EvalInto nd.A, env, v
        If Not Truthy(v) Then
            If Not (nd.C Is Nothing) Then ExecWhile = ExecBlock(nd.C, env)
            Exit Function
        End If
        st = ExecBlock(nd.B, env)
        If st = ST_BREAK Then Exit Function
        If st = ST_RETURN Then
            ExecWhile = ST_RETURN
            Exit Function
        End If
        gFrameLine(gDepth) = nd.Ln
    Loop
End Function

Private Function ExecFor(ByVal nd As PyNode, ByVal env As PyObj) As Long
    Dim seq As Variant, it As PyObj, x As Variant, st As Long, so As PyObj, tgt As PyNode
    Dim cur As Long, n As Long, k As Long, stp As Long, cv As Variant
    EvalInto nd.B, env, seq
    Set tgt = nd.A
    If IsObject(seq) Then
        Set so = seq
        If so.Kind = OK_RANGE And tgt.Kind = NK_NAME Then
            n = RangeLen(so)
            cur = so.I1
            stp = so.I3
            For k = 1 To n
                cv = cur
                StoreName env, tgt.S, tgt.H, cv
                st = ExecBlock(nd.C, env)
                If st = ST_BREAK Then Exit Function
                If st = ST_RETURN Then
                    ExecFor = ST_RETURN
                    Exit Function
                End If
                If k < n Then cur = cur + stp
            Next
            If Not (nd.D Is Nothing) Then ExecFor = ExecBlock(nd.D, env)
            Exit Function
        End If
    End If
    Set it = GetIter(seq)
    Do While IterNext(it, x)
        AssignTo tgt, x, env
        st = ExecBlock(nd.C, env)
        If st = ST_BREAK Then Exit Function
        If st = ST_RETURN Then
            ExecFor = ST_RETURN
            Exit Function
        End If
    Loop
    If Not (nd.D Is Nothing) Then ExecFor = ExecBlock(nd.D, env)
End Function

Private Function ExecTry(ByVal nd As PyNode, ByVal env As PyObj) As Long
    Dim st As Long, stF As Long, saveDepth As Long, i As Long, phase As Long
    Dim exc As PyObj, pending As PyObj, savedCur As PyObj, h As PyNode
    Dim tv As Variant, matched As Boolean, saveRet As Variant, ev As Variant
    saveDepth = gDepth
    Set savedCur = gCurExc
    phase = 0
    On Error GoTo EH
    st = ExecBlock(nd.B, env)
    If st = ST_NORMAL Then
        If Not (nd.C Is Nothing) Then
            phase = 1
            st = ExecBlock(nd.C, env)
        End If
    End If
    On Error GoTo 0
    GoTo RunFinally
EH:
    Set exc = CaptureExc()
    gDepth = saveDepth
    If phase = 0 Then
        Resume Handle
    End If
    Resume Pend
Handle:
    On Error GoTo EH2
    matched = False
    For i = 0 To nd.N - 1
        Set h = nd.Kid(i)
        If h.A Is Nothing Then
            matched = True
        Else
            EvalInto h.A, env, tv
            matched = ExcMatches(exc, tv)
        End If
        If matched Then
            If Len(h.S) > 0 Then
                Set ev = exc
                StoreName env, h.S, h.H, ev
            End If
            Set gCurExc = exc
            st = ExecBlock(h.B, env)
            Set gCurExc = savedCur
            Exit For
        End If
    Next
    On Error GoTo 0
    If Not matched Then Set pending = exc
    GoTo RunFinally
EH2:
    Set pending = CaptureExc()
    gDepth = saveDepth
    Set gCurExc = savedCur
    Resume RunFinally
Pend:
    Set pending = exc
RunFinally:
    On Error GoTo 0
    If Not (nd.D Is Nothing) Then
        If st = ST_RETURN Then SetV saveRet, gRet
        stF = ExecBlock(nd.D, env)
        If stF <> ST_NORMAL Then
            ExecTry = stF
            Exit Function
        End If
        If st = ST_RETURN Then SetV gRet, saveRet
    End If
    If Not (pending Is Nothing) Then RaiseExc pending
    ExecTry = st
End Function

Private Sub ExecYield(ByVal nd As PyNode, ByVal env As PyObj)
    Dim acc As Variant, lst As PyObj, v As Variant, it As PyObj, x As Variant
    If Not FindInChain(env, "__yield__", acc) Then ThrowErr "SyntaxError", "'yield' outside function"
    Set lst = acc
    If nd.A Is Nothing Then
        v = Empty
        lst.Push v
    ElseIf nd.Op = 1 Then
        EvalInto nd.A, env, v
        Set it = GetIter(v)
        Do While IterNext(it, x)
            lst.Push x
        Loop
    Else
        EvalInto nd.A, env, v
        lst.Push v
    End If
    If lst.Count > 1000000 Then ThrowErr "MemoryError", "generator produced more than 1000000 values (generators are evaluated eagerly)"
End Sub

Private Sub RecordAnnotation(ByVal cls As PyObj, ByVal nm As String)
    Dim ann As Variant, d As PyObj, e As Variant
    If Not NameGet(cls, "__annotations__", ann) Then
        Set d = NewDict()
        Set ann = d
        NameSet cls, "__annotations__", ann
    End If
    Set d = ann
    DictSetStr d, nm, e
End Sub

Public Function NewExcInst(ByVal cls As PyObj, ByVal argsTuple As PyObj) As PyObj
    Dim e As PyObj
    Set e = NewObj(OK_INST)
    Set e.Cls = cls
    NameSet e, "args", argsTuple
    Set NewExcInst = e
End Function

Private Sub ExecRaise(ByVal nd As PyNode, ByVal env As PyObj)
    Dim v As Variant, o As PyObj, a() As Variant, r As Variant
    If nd.A Is Nothing Then
        If gCurExc Is Nothing Then ThrowErr "RuntimeError", "No active exception to reraise"
        RaiseExc gCurExc
    End If
    EvalInto nd.A, env, v
    If IsObject(v) Then
        Set o = v
        If o.Kind = OK_CLASS Then
            ReDim a(0 To 0)
            Instantiate o, a, 0, Nothing, r
            Set o = r
        End If
        If o.Kind = OK_INST Then
            If IsSubclassOf(o.Cls, ExcClass("BaseException")) Then
                o.I3 = 0
                RaiseExc o
            End If
        End If
    End If
    ThrowErr "TypeError", "exceptions must derive from BaseException"
End Sub

Private Sub ExecDecl(ByVal nd As PyNode, ByVal env As PyObj)
    Dim i As Long, k As PyNode, v As Variant
    If env Is gGlobals Then Exit Sub
    If env.Decl Is Nothing Then Set env.Decl = NewObj(OK_ENV)
    If nd.Kind = NK_GLOBAL Then
        v = 1&
    Else
        v = 2&
    End If
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        env.Decl.SetH k.S, k.H, k.S, v
    Next
End Sub

Private Sub ApplyDecorators(ByVal decos As PyNode, ByVal env As PyObj, ByRef v As Variant)
    Dim i As Long, d As Variant, r As Variant
    If decos Is Nothing Then Exit Sub
    For i = decos.N - 1 To 0 Step -1
        EvalInto decos.Kid(i), env, d
        CallValue1 d, v, r
        SetV v, r
    Next
End Sub

Private Sub ExecDef(ByVal nd As PyNode, ByVal env As PyObj)
    Dim f As Variant
    MakeFunction nd, env, f
    ApplyDecorators nd.C, env, f
    StoreName env, nd.S, nd.H, f
End Sub

Private Sub MakeFunction(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim fo As PyObj, i As Long, pn As PyNode, v As Variant, defs As PyObj
    Set fo = NewObj(OK_FUNC)
    fo.PName = nd.S
    Set fo.Node = nd
    If env.Kind = OK_CLASS Then
        Set fo.Env = env.Env
        Set fo.Cls = env
    Else
        Set fo.Env = env
        If Not (env Is gGlobals) Then KeepObj env
    End If
    If nd.N > 0 Then
        Set defs = NewList()
        For i = 0 To nd.N - 1
            Set pn = nd.Kid(i)
            If pn.A Is Nothing Then
                Set v = gNoDefault
            Else
                EvalInto pn.A, env, v
            End If
            defs.Push v
        Next
        Set fo.Defaults = defs
    End If
    Set out = fo
End Sub

Private Sub ExecClass(ByVal nd As PyNode, ByVal env As PyObj)
    Dim cls As PyObj, i As Long, b As Variant, bo As PyObj, v As Variant
    Set cls = NewObj(OK_CLASS)
    cls.PName = nd.S
    Set cls.Env = env
    KeepObj cls
    For i = 0 To nd.N - 1
        EvalInto nd.Kid(i), env, b
        If Not IsObject(b) Then ThrowErr "TypeError", "bases must be classes"
        Set bo = b
        If bo.Kind = OK_BUILTIN Then
            If bo.PName = "object" Then GoTo NextBase
            ThrowErr "TypeError", "inheriting from the built-in type '" & bo.PName & "' is not supported"
        End If
        If bo.Kind <> OK_CLASS Then ThrowErr "TypeError", "bases must be classes"
        If Not (bo Is gObjectClass) Then cls.Push b
        If (bo.Flags And 4) <> 0 Then cls.Flags = cls.Flags Or 4
NextBase:
    Next
    ExecBlock nd.B, cls
    For i = 0 To cls.EntryCount - 1
        If cls.LiveAt(i) Then
            cls.ValInto i, v
            If IsKind(v, OK_PROP) Then cls.Flags = cls.Flags Or 4
        End If
    Next
    Set v = cls
    ApplyDecorators nd.C, env, v
    StoreName env, nd.S, nd.H, v
End Sub

Private Sub ExecAugAssign(ByVal nd As PyNode, ByVal env As PyObj)
    Dim t As PyNode, cur As Variant, rhs As Variant, res As Variant, obj As Variant, idx As Variant
    Dim vs As Variant, ve As Variant, vst As Variant
    Set t = nd.A
    Select Case t.Kind
    Case NK_NAME
        LoadName env, t.S, t.H, cur
        EvalInto nd.B, env, rhs
        InplaceOp nd.Op, cur, rhs, res
        StoreName env, t.S, t.H, res
    Case NK_SUBSCR
        EvalInto t.A, env, obj
        If t.B.Kind = NK_SLICE Then
            EvalSliceParts t.B, env, vs, ve, vst
            DoSlice obj, vs, ve, vst, cur
            EvalInto nd.B, env, rhs
            InplaceOp nd.Op, cur, rhs, res
            SetSlice obj, vs, ve, vst, res
        Else
            EvalInto t.B, env, idx
            GetItem obj, idx, cur
            EvalInto nd.B, env, rhs
            InplaceOp nd.Op, cur, rhs, res
            SetItem obj, idx, res
        End If
    Case NK_ATTR
        EvalInto t.A, env, obj
        PyGetAttr obj, t.S, cur
        EvalInto nd.B, env, rhs
        InplaceOp nd.Op, cur, rhs, res
        PySetAttr obj, t.S, res
    Case Else
        ThrowErr "SyntaxError", "illegal expression for augmented assignment"
    End Select
End Sub

Private Sub InplaceOp(ByVal op As Long, ByRef cur As Variant, ByRef rhs As Variant, ByRef out As Variant)
    Dim o As PyObj, src As PyObj, i As Long, x As Variant
    If op = OP_ADD Then
        If IsKind(cur, OK_LIST) Then
            Set o = cur
            Set src = NewListFrom(rhs)
            For i = 0 To src.Count - 1
                src.GetInto i, x
                o.Push x
            Next
            Set out = o
            Exit Sub
        End If
    End If
    BinOp op, cur, rhs, out
End Sub

Private Sub DeleteTarget(ByVal t As PyNode, ByVal env As PyObj)
    Dim obj As Variant, idx As Variant, i As Long, vs As Variant, ve As Variant, vst As Variant
    Select Case t.Kind
    Case NK_NAME
        If Not env.DelH(t.S, t.H) Then
            If Not gGlobals.DelH(t.S, t.H) Then ThrowErr "NameError", "name '" & t.S & "' is not defined"
        End If
    Case NK_SUBSCR
        EvalInto t.A, env, obj
        If t.B.Kind = NK_SLICE Then
            EvalSliceParts t.B, env, vs, ve, vst
            DelSlice obj, vs, ve, vst
        Else
            EvalInto t.B, env, idx
            DelItem obj, idx
        End If
    Case NK_ATTR
        EvalInto t.A, env, obj
        PyDelAttr obj, t.S
    Case NK_TUPLE, NK_LIST
        For i = 0 To t.N - 1
            DeleteTarget t.Kid(i), env
        Next
    Case Else
        ThrowErr "SyntaxError", "cannot delete expression"
    End Select
End Sub

Private Function ImportModule(ByVal nm As String) As PyObj
    Dim v As Variant, m As PyObj
    If NameGet(gModules, nm, v) Then
        Set ImportModule = v
        Exit Function
    End If
    Set m = BuildModule(nm)
    If m Is Nothing Then ThrowErr "ModuleNotFoundError", "No module named '" & nm & "'"
    Set v = m
    NameSet gModules, nm, v
    Set ImportModule = m
End Function

Private Sub ExecImport(ByVal nd As PyNode, ByVal env As PyObj)
    Dim i As Long, k As PyNode, m As PyObj, v As Variant, bindName As String
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        Set m = ImportModule(k.S)
        If VarType(k.V) = vbString Then
            bindName = k.V
        Else
            bindName = k.S
            If InStr(1, bindName, ".") > 0 Then bindName = Left$(bindName, InStr(1, bindName, ".") - 1)
        End If
        Set v = m
        StoreName env, bindName, HashStr(bindName), v
    Next
End Sub

Private Sub ExecFromImport(ByVal nd As PyNode, ByVal env As PyObj)
    Dim i As Long, k As PyNode, m As PyObj, v As Variant, j As Long, key As Variant, mv As Variant, bindName As String
    Set m = ImportModule(nd.S)
    Set mv = m
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        If k.S = "*" Then
            For j = 0 To m.EntryCount - 1
                If m.LiveAt(j) Then
                    m.KeyInto j, key
                    If Left$(key, 1) <> "_" Then
                        m.ValInto j, v
                        StoreName env, CStr(key), HashStr(CStr(key)), v
                    End If
                End If
            Next
        Else
            If Not TryGetAttr(mv, k.S, v) Then ThrowErr "ImportError", "cannot import name '" & k.S & "' from '" & nd.S & "'"
            If VarType(k.V) = vbString Then
                bindName = k.V
            Else
                bindName = k.S
            End If
            StoreName env, bindName, HashStr(bindName), v
        End If
    Next
End Sub

' ---------------------------------------------------------------------------
'  names and targets
' ---------------------------------------------------------------------------
Public Sub LoadName(ByVal env As PyObj, ByVal nm As String, ByVal h As Long, ByRef out As Variant)
    Dim e As PyObj, i As Long
    Set e = env
    Do
        i = e.FindH(nm, h)
        If i >= 0 Then
            e.ValInto i, out
            Exit Sub
        End If
        Set e = e.Env
    Loop Until e Is Nothing
    i = gBuiltins.FindH(nm, h)
    If i >= 0 Then
        gBuiltins.ValInto i, out
        Exit Sub
    End If
    ThrowErr "NameError", "name '" & nm & "' is not defined"
End Sub

Private Function FindInChain(ByVal env As PyObj, ByVal nm As String, ByRef out As Variant) As Boolean
    Dim e As PyObj, i As Long, h As Long
    h = HashStr(nm)
    Set e = env
    Do While Not (e Is Nothing)
        i = e.FindH(nm, h)
        If i >= 0 Then
            e.ValInto i, out
            FindInChain = True
            Exit Function
        End If
        Set e = e.Env
    Loop
End Function

Public Sub StoreName(ByVal env As PyObj, ByVal nm As String, ByVal h As Long, ByRef v As Variant)
    Dim i As Long, e As PyObj, dv As Variant
    If Not (env.Decl Is Nothing) Then
        i = env.Decl.FindH(nm, h)
        If i >= 0 Then
            env.Decl.ValInto i, dv
            If dv = 1 Then
                gGlobals.SetH nm, h, nm, v
                Exit Sub
            End If
            Set e = env.Env
            Do While Not (e Is Nothing)
                If e Is gGlobals Then Exit Do
                If e.FindH(nm, h) >= 0 Then
                    e.SetH nm, h, nm, v
                    Exit Sub
                End If
                Set e = e.Env
            Loop
            ThrowErr "SyntaxError", "no binding for nonlocal '" & nm & "' found"
        End If
    End If
    env.SetH nm, h, nm, v
End Sub

Public Sub AssignTo(ByVal t As PyNode, ByRef v As Variant, ByVal env As PyObj)
    Dim obj As Variant, idx As Variant, vs As Variant, ve As Variant, vst As Variant
    Select Case t.Kind
    Case NK_NAME
        StoreName env, t.S, t.H, v
    Case NK_TUPLE, NK_LIST
        Unpack t, v, env
    Case NK_SUBSCR
        EvalInto t.A, env, obj
        If t.B.Kind = NK_SLICE Then
            EvalSliceParts t.B, env, vs, ve, vst
            SetSlice obj, vs, ve, vst, v
        Else
            EvalInto t.B, env, idx
            SetItem obj, idx, v
        End If
    Case NK_ATTR
        EvalInto t.A, env, obj
        PySetAttrH obj, t.S, t.H, v
    Case NK_STARRED
        ThrowErr "SyntaxError", "starred assignment target must be in a list or tuple"
    Case Else
        ThrowErr "SyntaxError", "cannot assign to expression"
    End Select
End Sub

Private Sub Unpack(ByVal t As PyNode, ByRef v As Variant, ByVal env As PyObj)
    Dim lst As PyObj, n As Long, i As Long, starAt As Long, x As Variant, rest As PyObj, after As Long
    Dim rv As Variant
    If IsKind(v, OK_TUPLE) Then
        Set lst = v
    Else
        Set lst = NewListFrom(v)
    End If
    n = lst.Count
    starAt = -1
    For i = 0 To t.N - 1
        If t.Kid(i).Kind = NK_STARRED Then starAt = i
    Next
    If starAt < 0 Then
        If n <> t.N Then
            If n > t.N Then
                ThrowErr "ValueError", "too many values to unpack (expected " & t.N & ")"
            Else
                ThrowErr "ValueError", "not enough values to unpack (expected " & t.N & ", got " & n & ")"
            End If
        End If
        For i = 0 To t.N - 1
            lst.GetInto i, x
            AssignTo t.Kid(i), x, env
        Next
    Else
        after = t.N - starAt - 1
        If n < t.N - 1 Then ThrowErr "ValueError", "not enough values to unpack (expected at least " & (t.N - 1) & ", got " & n & ")"
        For i = 0 To starAt - 1
            lst.GetInto i, x
            AssignTo t.Kid(i), x, env
        Next
        Set rest = NewList()
        For i = starAt To n - after - 1
            lst.GetInto i, x
            rest.Push x
        Next
        Set rv = rest
        AssignTo t.Kid(starAt).A, rv, env
        For i = 0 To after - 1
            lst.GetInto n - after + i, x
            AssignTo t.Kid(starAt + 1 + i), x, env
        Next
    End If
End Sub

' ---------------------------------------------------------------------------
'  expressions
' ---------------------------------------------------------------------------
Public Sub EvalInto(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim a As Variant, b As Variant
    Select Case nd.Kind
    Case NK_NAME
        LoadName env, nd.S, nd.H, out
    Case NK_CONST
        If IsObject(nd.V) Then
            Set out = nd.V
        Else
            out = nd.V
        End If
    Case NK_BINOP
        EvalInto nd.A, env, a
        EvalInto nd.B, env, b
        BinOp nd.Op, a, b, out
    Case NK_COMPARE
        EvalInto nd.A, env, a
        EvalInto nd.B, env, b
        out = CompareOp(nd.Op, a, b)
    Case NK_CALL
        EvalCall nd, env, out
    Case NK_ATTR
        EvalInto nd.A, env, a
        PyGetAttrH a, nd.S, nd.H, out
    Case NK_SUBSCR
        EvalInto nd.A, env, a
        If nd.B.Kind = NK_SLICE Then
            Dim vs As Variant, ve As Variant, vst As Variant
            EvalSliceParts nd.B, env, vs, ve, vst
            DoSlice a, vs, ve, vst, out
        Else
            EvalInto nd.B, env, b
            GetItem a, b, out
        End If
    Case NK_AND
        EvalInto nd.A, env, a
        If Truthy(a) Then
            EvalInto nd.B, env, out
        Else
            SetV out, a
        End If
    Case NK_OR
        EvalInto nd.A, env, a
        If Truthy(a) Then
            SetV out, a
        Else
            EvalInto nd.B, env, out
        End If
    Case NK_NOT
        EvalInto nd.A, env, a
        out = Not Truthy(a)
    Case NK_UNARY
        EvalInto nd.A, env, a
        UnaryOp nd.Op, a, out
    Case NK_IFEXP
        EvalInto nd.A, env, a
        If Truthy(a) Then
            EvalInto nd.B, env, out
        Else
            EvalInto nd.C, env, out
        End If
    Case NK_CHAIN
        out = EvalChain(nd, env)
    Case NK_LIST, NK_TUPLE, NK_SET
        EvalSeqLiteral nd, env, out
    Case NK_DICT
        EvalDict nd, env, out
    Case NK_LISTCOMP, NK_GENEXP, NK_SETCOMP, NK_DICTCOMP
        EvalComp nd, env, out
    Case NK_LAMBDA
        MakeFunction nd, env, out
    Case NK_FSTRING
        out = EvalFString(nd, env)
    Case NK_WALRUS
        EvalInto nd.A, env, out
        StoreName env, nd.S, nd.H, out
    Case NK_STARRED
        ThrowErr "SyntaxError", "can't use starred expression here"
    Case Else
        ThrowErr "SystemError", "cannot evaluate node kind " & nd.Kind
    End Select
End Sub

Public Sub EvalSliceParts(ByVal sl As PyNode, ByVal env As PyObj, ByRef vs As Variant, ByRef ve As Variant, ByRef vst As Variant)
    If sl.A Is Nothing Then
        vs = Empty
    Else
        EvalInto sl.A, env, vs
    End If
    If sl.B Is Nothing Then
        ve = Empty
    Else
        EvalInto sl.B, env, ve
    End If
    If sl.C Is Nothing Then
        vst = Empty
    Else
        EvalInto sl.C, env, vst
    End If
End Sub

Private Function EvalChain(ByVal nd As PyNode, ByVal env As PyObj) As Boolean
    Dim a As Variant, b As Variant, i As Long
    EvalInto nd.Kid(0), env, a
    For i = 1 To nd.N - 1
        EvalInto nd.Kid(i), env, b
        If Not CompareOp(Asc(Mid$(nd.S, i, 1)), a, b) Then Exit Function
        SetV a, b
    Next
    EvalChain = True
End Function

Private Sub EvalSeqLiteral(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim r As PyObj, i As Long, k As PyNode, v As Variant, lst As PyObj, j As Long, x As Variant, kind As Long
    Select Case nd.Kind
    Case NK_LIST: kind = OK_LIST
    Case NK_TUPLE: kind = OK_TUPLE
    Case Else: kind = OK_SET
    End Select
    Set r = NewObj(kind)
    If kind <> OK_SET Then r.Reserve nd.N
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        If k.Kind = NK_STARRED Then
            EvalInto k.A, env, v
            Set lst = SeqOf(v)
            For j = 0 To lst.Count - 1
                lst.GetInto j, x
                If kind = OK_SET Then
                    SetAdd r, x
                Else
                    r.Push x
                End If
            Next
        Else
            EvalInto k, env, v
            If kind = OK_SET Then
                SetAdd r, v
            Else
                r.Push v
            End If
        End If
    Next
    Set out = r
End Sub

Private Sub EvalDict(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim d As PyObj, i As Long, k As Variant, v As Variant, kn As PyNode
    Set d = NewDict()
    i = 0
    Do While i < nd.N
        Set kn = nd.Kid(i)
        If kn.Kind = NK_DSTAR Then
            EvalInto kn.A, env, v
            If Not IsKind(v, OK_DICT) Then ThrowErr "TypeError", "'" & TypeNameOf(v) & "' object is not a mapping"
            DictUpdate d, v
        Else
            EvalInto kn, env, k
            EvalInto nd.Kid(i + 1), env, v
            DictSet d, k, v
        End If
        i = i + 2
    Loop
    Set out = d
End Sub

Private Sub EvalComp(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim cenv As PyObj, res As PyObj
    Set cenv = NewObj(OK_ENV)
    If env.Kind = OK_CLASS Then
        Set cenv.Env = env.Env
    Else
        Set cenv.Env = env
    End If
    Select Case nd.Kind
    Case NK_SETCOMP: Set res = NewObj(OK_SET)
    Case NK_DICTCOMP: Set res = NewDict()
    Case Else: Set res = NewList()
    End Select
    CompLoop nd, 0, cenv, res
    Set out = res
End Sub

Private Sub CompLoop(ByVal nd As PyNode, ByVal k As Long, ByVal env As PyObj, ByVal res As PyObj)
    Dim cl As PyNode, it As PyObj, x As Variant, v As Variant, kv As Variant
    If k >= nd.N Then
        EvalInto nd.A, env, v
        Select Case res.Kind
        Case OK_LIST
            res.Push v
        Case OK_SET
            SetAdd res, v
        Case OK_DICT
            EvalInto nd.B, env, kv
            DictSet res, v, kv
        End Select
        Exit Sub
    End If
    Set cl = nd.Kid(k)
    If cl.Kind = NK_COMPFOR Then
        EvalInto cl.B, env, x
        Set it = GetIter(x)
        Do While IterNext(it, v)
            AssignTo cl.A, v, env
            CompLoop nd, k + 1, env, res
        Loop
    Else
        EvalInto cl.A, env, v
        If Truthy(v) Then CompLoop nd, k + 1, env, res
    End If
End Sub

Private Function EvalFString(ByVal nd As PyNode, ByVal env As PyObj) As String
    Dim parts() As String, i As Long, k As PyNode, v As Variant, spec As String
    If nd.N = 0 Then Exit Function
    ReDim parts(0 To nd.N - 1)
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        If k.Kind = NK_CONST Then
            parts(i) = k.V
        Else
            EvalInto k.A, env, v
            If k.S = "r" Or k.S = "a" Then
                v = PyRepr(v)
            ElseIf k.S = "s" Then
                v = PyStr(v)
            End If
            If Not (k.B Is Nothing) Then
                spec = EvalFString(k.B, env)
            ElseIf VarType(k.V) = vbString Then
                spec = k.V
            Else
                spec = vbNullString
            End If
            If Len(spec) = 0 Then
                parts(i) = PyStr(v)
            Else
                parts(i) = FormatValue(v, spec)
            End If
        End If
    Next
    EvalFString = Join(parts, vbNullString)
End Function

' ---------------------------------------------------------------------------
'  calls
' ---------------------------------------------------------------------------
Private Sub EvalCall(ByVal nd As PyNode, ByVal env As PyObj, ByRef out As Variant)
    Dim fnode As PyNode, f As Variant, selfV As Variant, args() As Variant, na As Long, kw As PyObj
    Dim o As PyObj, mf As Variant, mode As Long, fo As PyObj
    Set fnode = nd.A
    ' mode 0: call f;  1: call user method fo with self;  2: built-in method of selfV
    If fnode.Kind = NK_ATTR Then
        EvalInto fnode.A, env, selfV
        If IsObject(selfV) Then
            Set o = selfV
            Select Case o.Kind
            Case OK_INST
                If o.FindH(fnode.S, fnode.H) < 0 Then
                    If ClassLookup(o.Cls, fnode.S, fnode.H, mf) Then
                        If IsKind(mf, OK_FUNC) Then
                            Set fo = mf
                            If fo.Flags = 0 Then mode = 1
                        End If
                    End If
                End If
                If mode = 0 Then PyGetAttrH selfV, fnode.S, fnode.H, f
            Case OK_LIST, OK_DICT, OK_SET, OK_TUPLE
                mode = 2
            Case Else
                PyGetAttrH selfV, fnode.S, fnode.H, f
            End Select
        Else
            mode = 2
        End If
    ElseIf fnode.Kind = NK_NAME Then
        If nd.N = 0 Then
            If fnode.S = "super" Then
                MakeSuper env, out
                Exit Sub
            End If
        End If
        LoadName env, fnode.S, fnode.H, f
    Else
        EvalInto fnode, env, f
    End If
    If mode = 1 Then
        EvalArgs nd, env, args, na, kw, 1
        SetV args(0), selfV
        CallFunc fo, args, na, kw, out
    Else
        EvalArgs nd, env, args, na, kw, 0
        If mode = 2 Then
            CallMethod selfV, fnode.S, args, na, kw, out
        Else
            CallValue f, args, na, kw, out
        End If
    End If
End Sub

Private Sub EvalArgs(ByVal nd As PyNode, ByVal env As PyObj, ByRef args() As Variant, ByRef na As Long, _
                     ByRef kw As PyObj, ByVal offset As Long)
    Dim i As Long, k As PyNode, v As Variant, lst As PyObj, j As Long, d As PyObj, kk As Variant
    ReDim args(0 To nd.N + offset)
    na = offset
    For i = 0 To nd.N - 1
        Set k = nd.Kid(i)
        Select Case k.Kind
        Case NK_KWARG
            If kw Is Nothing Then Set kw = NewDict()
            EvalInto k.A, env, v
            DictSetStr kw, k.S, v
        Case NK_STARRED
            EvalInto k.A, env, v
            Set lst = SeqOf(v)
            If na + lst.Count > UBound(args) Then
                ReDim Preserve args(0 To na + lst.Count + nd.N)
            End If
            For j = 0 To lst.Count - 1
                lst.GetInto j, v
                If IsObject(v) Then Set args(na) = v Else args(na) = v
                na = na + 1
            Next
        Case NK_DSTAR
            EvalInto k.A, env, v
            If Not IsKind(v, OK_DICT) Then ThrowErr "TypeError", "argument after ** must be a mapping, not " & TypeNameOf(v)
            Set d = v
            If kw Is Nothing Then Set kw = NewDict()
            For j = 0 To d.EntryCount - 1
                If d.LiveAt(j) Then
                    d.KeyInto j, kk
                    If VarType(kk) <> vbString Then ThrowErr "TypeError", "keywords must be strings"
                    d.ValInto j, v
                    DictSetStr kw, CStr(kk), v
                End If
            Next
        Case Else
            If na > UBound(args) Then
                ReDim Preserve args(0 To na * 2 + 1)
            End If
            EvalInto k, env, v
            If IsObject(v) Then Set args(na) = v Else args(na) = v
            na = na + 1
        End Select
    Next
End Sub

Public Sub CallValue(ByRef f As Variant, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim o As PyObj, a2() As Variant, i As Long, sv As Variant, fv As Variant, cm As Variant
    If Not IsObject(f) Then ThrowErr "TypeError", "'" & TypeNameOf(f) & "' object is not callable"
    Set o = f
    Select Case o.Kind
    Case OK_FUNC
        CallFunc o, args, na, kw, out
    Case OK_BUILTIN
        CallBuiltin o, args, na, kw, out
    Case OK_BOUND
        ReDim a2(0 To na)
        SetV sv, o.SelfV
        SetV a2(0), sv
        For i = 0 To na - 1
            SetV a2(i + 1), args(i)
        Next
        SetV fv, o.FnV
        CallValue fv, a2, na + 1, kw, out
    Case OK_BMETHOD
        If o.Flags = 1 Then
            ' unbound method of a built-in type, e.g. str.lower(s)
            If na = 0 Then ThrowErr "TypeError", "unbound method " & o.PName & "() needs an argument"
            ReDim a2(0 To na)
            For i = 1 To na - 1
                SetV a2(i - 1), args(i)
            Next
            SetV sv, args(0)
            CallMethod sv, o.PName, a2, na - 1, kw, out
        Else
            SetV sv, o.SelfV
            CallMethod sv, o.PName, args, na, kw, out
        End If
    Case OK_CLASS
        Instantiate o, args, na, kw, out
    Case OK_INST
        If ClassLookup(o.Cls, "__call__", HashStr("__call__"), cm) Then
            ReDim a2(0 To na)
            Set a2(0) = o
            For i = 0 To na - 1
                SetV a2(i + 1), args(i)
            Next
            CallValue cm, a2, na + 1, kw, out
        Else
            ThrowErr "TypeError", "'" & TypeNameOf(f) & "' object is not callable"
        End If
    Case Else
        ThrowErr "TypeError", "'" & TypeNameOf(f) & "' object is not callable"
    End Select
End Sub

Public Sub CallValue0(ByRef f As Variant, ByRef out As Variant)
    Dim args() As Variant
    ReDim args(0 To 0)
    CallValue f, args, 0, Nothing, out
End Sub

Public Sub CallValue1(ByRef f As Variant, ByRef a As Variant, ByRef out As Variant)
    Dim args() As Variant
    ReDim args(0 To 0)
    SetV args(0), a
    CallValue f, args, 1, Nothing, out
End Sub

Public Sub CallValue2(ByRef f As Variant, ByRef a As Variant, ByRef b As Variant, ByRef out As Variant)
    Dim args() As Variant
    ReDim args(0 To 1)
    SetV args(0), a
    SetV args(1), b
    CallValue f, args, 2, Nothing, out
End Sub

' call a special method defined by the class of an instance
Public Function CallDunder0(ByRef selfV As Variant, ByVal nm As String, ByRef out As Variant) As Boolean
    Dim o As PyObj, f As Variant, a() As Variant
    If Not IsObject(selfV) Then Exit Function
    Set o = selfV
    If o.Kind <> OK_INST Then Exit Function
    If Not ClassLookup(o.Cls, nm, HashStr(nm), f) Then Exit Function
    ReDim a(0 To 0)
    Set a(0) = o
    CallValue f, a, 1, Nothing, out
    CallDunder0 = True
End Function

Public Function CallDunder1(ByRef selfV As Variant, ByVal nm As String, ByRef arg As Variant, ByRef out As Variant) As Boolean
    Dim o As PyObj, f As Variant, a() As Variant
    If Not IsObject(selfV) Then Exit Function
    Set o = selfV
    If o.Kind <> OK_INST Then Exit Function
    If Not ClassLookup(o.Cls, nm, HashStr(nm), f) Then Exit Function
    ReDim a(0 To 1)
    Set a(0) = o
    SetV a(1), arg
    CallValue f, a, 2, Nothing, out
    CallDunder1 = True
End Function

Public Function CallDunder2(ByRef selfV As Variant, ByVal nm As String, ByRef arg1 As Variant, ByRef arg2 As Variant, ByRef out As Variant) As Boolean
    Dim o As PyObj, f As Variant, a() As Variant
    If Not IsObject(selfV) Then Exit Function
    Set o = selfV
    If o.Kind <> OK_INST Then Exit Function
    If Not ClassLookup(o.Cls, nm, HashStr(nm), f) Then Exit Function
    ReDim a(0 To 2)
    Set a(0) = o
    SetV a(1), arg1
    SetV a(2), arg2
    CallValue f, a, 3, Nothing, out
    CallDunder2 = True
End Function

' next() on a user-defined iterator; False on StopIteration
Public Function UserIterNext(ByVal inst As PyObj, ByRef out As Variant) As Boolean
    Dim exc As PyObj, saveDepth As Long, v As Variant
    saveDepth = gDepth
    Set v = inst
    On Error GoTo EH
    If Not CallDunder0(v, "__next__", out) Then
        On Error GoTo 0
        ThrowErr "TypeError", "'" & inst.Cls.PName & "' object is not an iterator"
    End If
    UserIterNext = True
    Exit Function
EH:
    Set exc = CaptureExc()
    gDepth = saveDepth
    If IsSubclassOf(exc.Cls, ExcClass("StopIteration")) Then
        UserIterNext = False
        Exit Function
    End If
    RaiseExc exc
End Function

Private Sub MakeSuper(ByVal env As PyObj, ByRef out As Variant)
    Dim c As Variant, s As Variant, so As PyObj
    If Not FindInChain(env, "__class__", c) Then ThrowErr "RuntimeError", "super(): no arguments"
    If Not FindInChain(env, "__selfarg__", s) Then ThrowErr "RuntimeError", "super(): no arguments"
    Set so = NewObj(OK_SUPER)
    Set so.Cls = c
    If IsObject(s) Then
        Set so.SelfV = s
    Else
        so.SelfV = s
    End If
    Set out = so
End Sub

Public Sub CallFunc(ByVal fn As PyObj, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim env As PyObj, dn As PyNode, st As Long, cv As Variant, gen As PyObj, gv As Variant
    Set dn = fn.Node
    If gDepth >= gMaxDepth Then ThrowErr "RecursionError", "maximum recursion depth exceeded"
    Set env = NewObj(OK_ENV)
    Set env.Env = fn.Env
    BindParams fn, dn, env, args, na, kw
    If (dn.Op And 2) <> 0 Then
        Set gen = NewList()
        Set gv = gen
        NameSet env, "__yield__", gv
    End If
    If (dn.Op And 1) <> 0 Then
        If Not (fn.Cls Is Nothing) Then
            Set cv = fn.Cls
            NameSet env, "__class__", cv
            If na > 0 Then NameSet env, "__selfarg__", args(0)
        End If
    End If
    gDepth = gDepth + 1
    gFrameName(gDepth) = fn.PName
    gFrameLine(gDepth) = dn.Ln
    If dn.Kind = NK_LAMBDA Then
        EvalInto dn.A, env, out
    Else
        st = ExecBlock(dn.B, env)
        If Not (gen Is Nothing) Then
            gRet = Empty
            Set out = GetIter(gv)
        ElseIf st = ST_RETURN Then
            If IsObject(gRet) Then
                Set out = gRet
            Else
                out = gRet
            End If
            gRet = Empty
        Else
            out = Empty
        End If
    End If
    gDepth = gDepth - 1
End Sub

Private Sub BindParams(ByVal fn As PyObj, ByVal dn As PyNode, ByVal env As PyObj, ByRef args() As Variant, _
                       ByVal na As Long, ByVal kw As PyObj)
    Dim np As Long, i As Long, ai As Long, pn As PyNode, starIdx As Long, dstarIdx As Long, nPos As Long
    Dim isSet() As Boolean, rest As PyObj, kwd As PyObj, e As Long, k As Variant, v As Variant, found As Boolean
    Dim j As Long, missing As String, nMissing As Long
    np = dn.N
    starIdx = -1
    dstarIdx = -1
    If np > 0 Then
        ReDim isSet(0 To np - 1)
    End If
    ai = 0
    For i = 0 To np - 1
        Set pn = dn.Kid(i)
        Select Case pn.Op
        Case 0
            nPos = nPos + 1
            If ai < na Then
                env.SetH pn.S, pn.H, pn.S, args(ai)
                isSet(i) = True
                ai = ai + 1
            End If
        Case 1
            starIdx = i
        Case 2
            dstarIdx = i
        End Select
    Next
    If ai < na And starIdx < 0 Then
        ThrowErr "TypeError", fn.PName & "() takes " & nPos & " positional argument" & IIf(nPos = 1, "", "s") & " but " & na & IIf(na = 1, " was", " were") & " given"
    End If
    If starIdx >= 0 Then
        Set rest = NewObj(OK_TUPLE)
        For j = ai To na - 1
            rest.Push args(j)
        Next
        Set pn = dn.Kid(starIdx)
        Set v = rest
        env.SetH pn.S, pn.H, pn.S, v
        isSet(starIdx) = True
    End If
    If dstarIdx >= 0 Then
        Set kwd = NewDict()
        Set pn = dn.Kid(dstarIdx)
        Set v = kwd
        env.SetH pn.S, pn.H, pn.S, v
        isSet(dstarIdx) = True
    End If
    If Not (kw Is Nothing) Then
        For e = 0 To kw.EntryCount - 1
            If kw.LiveAt(e) Then
                kw.KeyInto e, k
                kw.ValInto e, v
                found = False
                For i = 0 To np - 1
                    Set pn = dn.Kid(i)
                    If pn.Op = 0 Or pn.Op = 3 Then
                        If pn.S = k Then
                            If isSet(i) Then ThrowErr "TypeError", fn.PName & "() got multiple values for argument '" & k & "'"
                            env.SetH pn.S, pn.H, pn.S, v
                            isSet(i) = True
                            found = True
                            Exit For
                        End If
                    End If
                Next
                If Not found Then
                    If kwd Is Nothing Then ThrowErr "TypeError", fn.PName & "() got an unexpected keyword argument '" & k & "'"
                    DictSet kwd, k, v
                End If
            End If
        Next
    End If
    For i = 0 To np - 1
        If Not isSet(i) Then
            Set pn = dn.Kid(i)
            fn.Defaults.GetInto i, v
            If IsObject(v) Then
                If v Is gNoDefault Then
                    nMissing = nMissing + 1
                    If Len(missing) > 0 Then missing = missing & ", "
                    missing = missing & "'" & pn.S & "'"
                    GoTo NextParam
                End If
            End If
            env.SetH pn.S, pn.H, pn.S, v
        End If
NextParam:
    Next
    If nMissing > 0 Then
        ThrowErr "TypeError", fn.PName & "() missing " & nMissing & " required argument" & IIf(nMissing = 1, "", "s") & ": " & missing
    End If
End Sub

Public Sub Instantiate(ByVal cls As PyObj, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim inst As PyObj, f As Variant, a2() As Variant, i As Long, r As Variant
    Set inst = NewObj(OK_INST)
    Set inst.Cls = cls
    If ClassLookup(cls, "__init__", HashStr("__init__"), f) Then
        ReDim a2(0 To na)
        Set a2(0) = inst
        For i = 0 To na - 1
            SetV a2(i + 1), args(i)
        Next
        CallValue f, a2, na + 1, kw, r
        If Not IsEmpty(r) Then ThrowErr "TypeError", "__init__() should return None, not '" & TypeNameOf(r) & "'"
    ElseIf na > 0 Then
        ThrowErr "TypeError", cls.PName & "() takes no arguments"
    End If
    Set out = inst
End Sub
