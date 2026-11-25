Attribute VB_Name = "LispExcel"
Option Explicit
' ===========================================================================
'  LispExcel - Excel integration of the Lisp (Scheme subset) interpreter.
'
'  Macros
'    SetupLispSheet    creates the "Lisp" sheet (code in column A, output in
'                      column C, a Run button) and a "Data" sheet
'    RunLisp           runs the program in column A of the "Lisp" sheet;
'                      values of top-level expressions are shown as "=> value"
'    RunSelectedLisp   runs the code in the selected cell(s), shows a MsgBox
'
'  Worksheet function
'    =LISP(code, [arg1], [arg2], ...)   arguments are bound to x, y, z and
'                                       the list args; returns the value of
'                                       the last expression
'
'  Lisp procedures that talk to Excel (the optional last argument is a sheet name)
'    (cell row col [sheet]) / (cell "B3" [sheet])
'    (set-cell! row col value [sheet]) / (set-cell! "B3" value [sheet])
'    (get-range "A1:C10" [sheet])   -> list of rows (flat list for one row/column)
'    (set-range! "A1" data [sheet]) -> data: value, list (one column) or list of lists
'    (clear-range! "A1:C10" [sheet]), (sheet-names), (excel-eval "SUM(A1:A9)"), (msgbox text)
'    The default sheet is "Data" (created when missing); inside LISP() it is
'    the sheet of the formula.  Empty cells read as ().
' ===========================================================================

Private Const CODE_SHEET As String = "Lisp"
Private Const DATA_SHEET As String = "Data"
Private Const OUT_COLUMN As Long = 3

Private mBusy As Boolean
Private mUdfSheet As Object

' ---------------------------------------------------------------------------
'  macros
' ---------------------------------------------------------------------------
Public Sub RunLisp()
    Dim ws As Object, src As String, outText As String, calc As Long, t0 As Single
    If mBusy Then Exit Sub
    Set ws = FindSheet(CODE_SHEET)
    If ws Is Nothing Then
        MsgBox "Sheet '" & CODE_SHEET & "' was not found. Run the macro SetupLispSheet first.", vbExclamation, "Lisp"
        Exit Sub
    End If
    src = ReadCode(ws)
    mBusy = True
    Set mUdfSheet = Nothing
    calc = Application.Calculation
    On Error GoTo Fail
    Application.Calculation = -4135      ' xlCalculationManual
    Application.EnableCancelKey = 2      ' xlErrorHandler: Ctrl+Break stops the program
    Application.StatusBar = "Running Lisp..."
    t0 = Timer
    outText = LispRunProgram(src, True)
Done:
    On Error Resume Next
    Application.Calculation = calc
    Application.EnableCancelKey = 1
    Application.StatusBar = False
    On Error GoTo 0
    mBusy = False
    WriteOutput ws, outText
    Application.StatusBar = "Lisp finished in " & Format$(Timer - t0, "0.00") & " s"
    Exit Sub
Fail:
    outText = outText & vbLf & "Internal error " & Err.Number & ": " & Err.Description
    Resume Done
End Sub

Public Sub RunSelectedLisp()
    Dim src As String, outText As String, c As Object
    If mBusy Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    For Each c In Selection.Cells
        If Len(src) > 0 Then src = src & vbLf
        src = src & CStr(c.Formula)
    Next
    mBusy = True
    Set mUdfSheet = Nothing
    On Error GoTo Fail
    outText = LispRunProgram(src, True)
Done:
    mBusy = False
    If Len(outText) > 900 Then outText = Left$(outText, 900) & vbLf & "..."
    MsgBox outText, vbInformation, "Lisp output"
    Exit Sub
Fail:
    outText = "Internal error " & Err.Number & ": " & Err.Description
    Resume Done
End Sub

Public Sub SetupLispSheet()
    Dim ws As Object, ds As Object, lines As Variant, i As Long, btn As Object
    Set ws = FindSheet(CODE_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = CODE_SHEET
        lines = SampleProgram()
        ws.Columns(1).NumberFormat = "@"
        For i = LBound(lines) To UBound(lines)
            ws.Cells(i - LBound(lines) + 1, 1).Value = lines(i)
        Next
    End If
    ws.Columns(1).ColumnWidth = 70
    ws.Columns(2).ColumnWidth = 2
    ws.Columns(OUT_COLUMN).ColumnWidth = 60
    ws.Columns(1).Font.Name = "Consolas"
    ws.Columns(OUT_COLUMN).Font.Name = "Consolas"
    ws.Columns(OUT_COLUMN).Interior.Color = RGB(245, 245, 245)
    On Error Resume Next
    ws.Buttons.Delete
    Set btn = ws.Buttons.Add(ws.Cells(1, OUT_COLUMN + 2).Left + 4, ws.Cells(1, 1).Top + 2, 110, 26)
    btn.Caption = "Run Lisp"
    btn.OnAction = "RunLisp"
    On Error GoTo 0
    Set ds = FindSheet(DATA_SHEET)
    If ds Is Nothing Then
        Set ds = ThisWorkbook.Worksheets.Add(After:=ws)
        ds.Name = DATA_SHEET
        ds.Range("A1:B1").Value = Array("Item", "Amount")
        ds.Range("A2:B2").Value = Array("Apples", 12)
        ds.Range("A3:B3").Value = Array("Pears", 7.5)
        ds.Range("A4:B4").Value = Array("Plums", 20)
        ds.Range("A5:B5").Value = Array("Figs", 3)
    End If
    ws.Activate
End Sub

Private Function SampleProgram() As Variant
    SampleProgram = Array( _
        "; Scheme running inside Excel VBA. Edit column A, then click 'Run Lisp'.", _
        "(define (fact n) (if (= n 0) 1 (* n (fact (- n 1)))))", _
        "(fact 12)", _
        "(define (fib n)", _
        "  (let loop ((a 0) (b 1) (i 0))", _
        "    (if (= i n) a (loop b (+ a b) (+ i 1)))))", _
        "(map fib (iota 12))", _
        "", _
        "(define-syntax swap!", _
        "  (syntax-rules ()", _
        "    ((_ a b) (let ((tmp a)) (set! a b) (set! b tmp)))))", _
        "(define x 1)", _
        "(define y 2)", _
        "(swap! x y)", _
        "(list x y)", _
        "", _
        "(define rows (get-range ""A1:B5"" ""Data""))   ; header + 4 rows", _
        "(define total (apply + (map cadr (cdr rows))))", _
        "(display ""total = "") (display total) (newline)", _
        "(set-cell! ""D1"" ""Total"" ""Data"")", _
        "(set-cell! ""D2"" total ""Data"")", _
        "(sort (map car (cdr rows)) string<?)", _
        "", _
        "(guard (e (#t (display ""caught: "") (display (condition/report-string e)) (newline)))", _
        "  (error ""something went wrong:"" 42))")
End Function

' ---------------------------------------------------------------------------
'  worksheet function
' ---------------------------------------------------------------------------
Public Function LISP(ByVal code As String, ParamArray args() As Variant) As Variant
    Dim argv() As Variant, n As Long, i As Long, result As Variant, printed As String, errText As String
    If mBusy Then
        LISP = "#BUSY"
        Exit Function
    End If
    On Error GoTo Fail
    n = UBound(args) - LBound(args) + 1
    If n > 0 Then
        ReDim argv(0 To n - 1)
        For i = 0 To n - 1
            XlArgToLisp args(LBound(args) + i), argv(i)
        Next
    Else
        ReDim argv(0 To 0)
    End If
    mBusy = True
    Set mUdfSheet = Nothing
    If TypeName(Application.Caller) = "Range" Then Set mUdfSheet = Application.Caller.Worksheet
    If LispEvalCode(code, argv, n, result, printed, errText) Then
        If LIsKind(result, LK_VOID) Then
            If Right$(printed, 1) = vbLf Then printed = Left$(printed, Len(printed) - 1)
            LISP = printed
        Else
            LISP = LispToXlResult(result)
        End If
    Else
        LISP = "#Error: " & errText
    End If
    LispFinish
    mBusy = False
    Exit Function
Fail:
    LISP = "#Internal error: " & Err.Description
    mBusy = False
End Function

' ---------------------------------------------------------------------------
'  host procedures called by the interpreter
' ---------------------------------------------------------------------------
Public Function LispHostInput(ByVal prompt As String) As String
    LispHostInput = InputBox(prompt, "Lisp")
End Function

Public Sub LispHostCall(ByVal code As Long, ByRef args() As Variant, ByVal na As Long, ByRef out As Variant)
    Dim ws As Object, rng As Object, v As Variant, i As Long, head As LispObj, tail As LispObj, cell As LispObj, e As Variant
    Set out = lgVoid
    Select Case code
    Case LP_X_CELL
        If VarType(args(0)) = vbString Then
            Set ws = HostSheet(args, na, 2)
            Set rng = ws.Range(args(0))
        Else
            If na < 2 Then LThrow "cell: expected (cell row col [sheet]) or (cell ""A1"" [sheet])"
            Set ws = HostSheet(args, na, 3)
            Set rng = ws.Cells(HostLong(args(0), "cell"), HostLong(args(1), "cell"))
        End If
        XlToLisp rng.Cells(1, 1).Value2, out
    Case LP_X_SETCELL
        If VarType(args(0)) = vbString Then
            Set ws = HostSheet(args, na, 3)
            ws.Range(args(0)).Cells(1, 1).Value = LispToXlScalar(args(1))
        Else
            If na < 3 Then LThrow "set-cell!: expected (set-cell! row col value [sheet]) or (set-cell! ""A1"" value [sheet])"
            Set ws = HostSheet(args, na, 4)
            ws.Cells(HostLong(args(0), "set-cell!"), HostLong(args(1), "set-cell!")).Value = LispToXlScalar(args(2))
        End If
    Case LP_X_GETRANGE
        Set ws = HostSheet(args, na, 2)
        Set rng = ws.Range(HostStr(args(0), "get-range"))
        XlValuesToLisp rng.Value2, out
    Case LP_X_SETRANGE
        Set ws = HostSheet(args, na, 3)
        Set rng = ws.Range(HostStr(args(0), "set-range!"))
        WriteLispData rng.Cells(1, 1), args(1)
    Case LP_X_CLEAR
        Set ws = HostSheet(args, na, 2)
        ws.Range(HostStr(args(0), "clear-range!")).ClearContents
    Case LP_X_MSGBOX
        If na = 2 Then
            MsgBox LToStr(args(0), False), vbOKOnly, LToStr(args(1), False)
        Else
            MsgBox LToStr(args(0), False), vbOKOnly, "Lisp"
        End If
    Case LP_X_EVAL
        v = Application.Evaluate(HostStr(args(0), "excel-eval"))
        If IsError(v) Then LThrow2 "excel-eval: Excel could not evaluate", args(0)
        If IsArray(v) Then
            XlValuesToLisp v, out
        Else
            XlToLisp v, out
        End If
    Case LP_X_SHEETS
        For i = ThisWorkbook.Worksheets.Count To 1 Step -1
            v = ThisWorkbook.Worksheets(i).Name
            Set e = LCons(v, e)
        Next
        LSetV out, e
    Case Else
        LThrow "unknown Excel procedure"
    End Select
End Sub

Private Function HostStr(ByRef v As Variant, ByVal who As String) As String
    If VarType(v) <> vbString Then LThrow2 who & ": expected an address string such as ""A1"", got", v
    HostStr = v
End Function

Private Function HostLong(ByRef v As Variant, ByVal who As String) As Long
    If Not LIsNum(v) Then LThrow2 who & ": expected a row/column number, got", v
    HostLong = CLng(v)
End Function

' the sheet: args(pos-1) if present (a name or an index), else the default sheet
Private Function HostSheet(ByRef args() As Variant, ByVal na As Long, ByVal pos As Long) As Object
    Dim ws As Object
    If na >= pos Then
        If VarType(args(pos - 1)) = vbString Then
            Set ws = FindSheet(CStr(args(pos - 1)))
            If ws Is Nothing And StrComp(CStr(args(pos - 1)), DATA_SHEET, vbTextCompare) = 0 Then Set ws = AddDataSheet()
            If ws Is Nothing Then LThrow2 "no worksheet named", args(pos - 1)
        ElseIf LIsNum(args(pos - 1)) Then
            Set ws = ThisWorkbook.Worksheets(CLng(args(pos - 1)))
        Else
            LThrow2 "expected a sheet name, got", args(pos - 1)
        End If
    ElseIf Not (mUdfSheet Is Nothing) Then
        Set ws = mUdfSheet
    Else
        Set ws = FindSheet(DATA_SHEET)
        If ws Is Nothing Then Set ws = AddDataSheet()
    End If
    Set HostSheet = ws
End Function

Private Function AddDataSheet() As Object
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = DATA_SHEET
    Set AddDataSheet = ws
End Function

Private Function FindSheet(ByVal nm As String) As Object
    Dim ws As Object
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next
    Set FindSheet = Nothing
End Function

Private Function ReadCode(ByVal ws As Object) As String
    Dim lastRow As Long, v As Variant, i As Long, parts() As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row     ' xlUp
    If lastRow = 1 Then
        ReadCode = CStr(ws.Cells(1, 1).Formula)
        Exit Function
    End If
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Formula
    ReDim parts(0 To lastRow - 1)
    For i = 1 To lastRow
        parts(i - 1) = CStr(v(i, 1))
    Next
    ReadCode = Join(parts, vbLf)
End Function

Private Sub WriteOutput(ByVal ws As Object, ByVal txt As String)
    Dim lines() As String, n As Long, arr() As Variant, i As Long, s As String
    ws.Columns(OUT_COLUMN).ClearContents
    If Right$(txt, 1) = vbLf Then txt = Left$(txt, Len(txt) - 1)
    If Len(txt) = 0 Then Exit Sub
    lines = Split(txt, vbLf)
    n = UBound(lines) + 1
    If n > 1048575 Then n = 1048575
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        s = lines(i - 1)
        If Len(s) > 32000 Then s = Left$(s, 32000) & "..."
        arr(i, 1) = "'" & s
    Next
    ws.Range(ws.Cells(1, OUT_COLUMN), ws.Cells(n, OUT_COLUMN)).Value = arr
End Sub

' ---------------------------------------------------------------------------
'  value conversion Excel <-> Lisp
' ---------------------------------------------------------------------------
Private Sub XlToLisp(ByVal v As Variant, ByRef out As Variant)
    Select Case VarType(v)
    Case vbDouble, vbSingle, vbCurrency
        If v = Int(v) And Abs(v) <= 2147483647# Then
            out = CLng(v)
        Else
            out = CDbl(v)
        End If
    Case vbString, vbBoolean, vbLong
        out = v
    Case vbInteger, vbByte
        out = CLng(v)
    Case vbDate
        out = Format$(v, "yyyy-mm-dd hh:nn:ss")
        If Right$(out, 9) = " 00:00:00" Then out = Left$(out, 10)
    Case vbEmpty, vbNull
        out = Empty
    Case vbError
        out = "#ERROR"
    Case vbDecimal
        out = LNorm(v)
    Case Else
        out = CStr(v)
    End Select
End Sub

Private Sub XlValuesToLisp(ByVal v As Variant, ByRef out As Variant)
    Dim r As Long, c As Long, r0 As Long, r1 As Long, c0 As Long, c1 As Long, x As Variant
    Dim rows As Variant, row As Variant, e As Variant
    If Not IsArray(v) Then
        XlToLisp v, out
        Exit Sub
    End If
    r0 = LBound(v, 1)
    r1 = UBound(v, 1)
    c0 = LBound(v, 2)
    c1 = UBound(v, 2)
    rows = Empty
    If r0 = r1 Or c0 = c1 Then
        For r = r1 To r0 Step -1
            For c = c1 To c0 Step -1
                XlToLisp v(r, c), x
                Set rows = LCons(x, rows)
            Next
        Next
    Else
        For r = r1 To r0 Step -1
            row = Empty
            For c = c1 To c0 Step -1
                XlToLisp v(r, c), x
                Set row = LCons(x, row)
            Next
            Set rows = LCons(row, rows)
        Next
    End If
    LSetV out, rows
End Sub

Private Sub XlArgToLisp(ByVal a As Variant, ByRef out As Variant)
    Dim rg As Object
    If IsObject(a) Then
        Set rg = a
        If TypeName(rg) = "Range" Then
            If rg.Cells.Count = 1 Then
                XlToLisp rg.Value2, out
            Else
                XlValuesToLisp rg.Value2, out
            End If
        Else
            out = TypeName(rg)
        End If
        Exit Sub
    End If
    If IsArray(a) Then
        XlValuesToLisp a, out
        Exit Sub
    End If
    XlToLisp a, out
End Sub

Private Function LispToXlScalar(ByRef v As Variant) As Variant
    Dim o As LispObj
    Select Case VarType(v)
    Case vbEmpty, vbNull
        LispToXlScalar = Empty
    Case vbDecimal
        LispToXlScalar = CDbl(v)
    Case vbLong, vbInteger, vbDouble, vbBoolean, vbSingle, vbCurrency, vbString
        LispToXlScalar = v
    Case vbObject
        Set o = v
        If o.Kind = LK_VOID Then
            LispToXlScalar = Empty
        Else
            LispToXlScalar = LToStr(v, False)
        End If
    Case Else
        LispToXlScalar = LToStr(v, False)
    End Select
End Function

' a list as rows: True when every element is a list
Private Function IsTable(ByRef v As Variant, ByRef nr As Long, ByRef nc As Long) As Boolean
    Dim x As Variant, p As LispObj, n As Long
    nr = 0
    nc = 0
    IsTable = True
    LSetV x, v
    Do While LIsPair(x)
        Set p = x
        nr = nr + 1
        If LIsPair(p.Car) Or IsEmpty(p.Car) Then
            n = LLength(p.Car)
            If n < 0 Then IsTable = False
            If n > nc Then nc = n
        Else
            IsTable = False
        End If
        LSetV x, p.Cdr
    Loop
    If nc = 0 Then IsTable = False
End Function

Private Sub ListToArray2D(ByRef v As Variant, ByRef arr() As Variant, ByVal base0 As Long)
    Dim nr As Long, nc As Long, x As Variant, y As Variant, p As LispObj, q As LispObj, i As Long, j As Long, o As LispObj
    If LIsKind(v, LK_VECTOR) Then
        Set o = v
        ReDim arr(base0 To base0, base0 To base0 + IIf(o.Count > 0, o.Count - 1, 0))
        For j = 0 To o.Count - 1
            o.GetInto j, x
            arr(base0, base0 + j) = LispToXlScalar(x)
        Next
        Exit Sub
    End If
    If IsTable(v, nr, nc) Then
        ReDim arr(base0 To base0 + nr - 1, base0 To base0 + nc - 1)
        LSetV x, v
        i = 0
        Do While LIsPair(x)
            Set p = x
            LSetV y, p.Car
            For j = 0 To nc - 1
                If LIsPair(y) Then
                    Set q = y
                    arr(base0 + i, base0 + j) = LispToXlScalar(q.Car)
                    LSetV y, q.Cdr
                Else
                    arr(base0 + i, base0 + j) = Empty
                End If
            Next
            i = i + 1
            LSetV x, p.Cdr
        Loop
    Else
        nr = LLength(v)
        If nr <= 0 Then nr = 1
        ReDim arr(base0 To base0 + nr - 1, base0 To base0)
        LSetV x, v
        i = 0
        Do While LIsPair(x)
            Set p = x
            arr(base0 + i, base0) = LispToXlScalar(p.Car)
            i = i + 1
            LSetV x, p.Cdr
        Loop
    End If
End Sub

Private Function LispToXlResult(ByRef v As Variant) As Variant
    Dim arr() As Variant
    If LIsPair(v) Or LIsKind(v, LK_VECTOR) Then
        ListToArray2D v, arr, 0
        LispToXlResult = arr
    ElseIf IsEmpty(v) Then
        LispToXlResult = ""
    Else
        LispToXlResult = LispToXlScalar(v)
    End If
End Function

Private Sub WriteLispData(ByVal topLeft As Object, ByRef data As Variant)
    Dim arr() As Variant
    If LIsPair(data) Or LIsKind(data, LK_VECTOR) Then
        ListToArray2D data, arr, 1
        topLeft.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    Else
        topLeft.Value = LispToXlScalar(data)
    End If
End Sub
