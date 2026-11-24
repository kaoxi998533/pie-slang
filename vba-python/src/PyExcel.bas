Attribute VB_Name = "PyExcel"
Option Explicit
' ===========================================================================
'  PyExcel - Excel integration of the Python-subset interpreter.
'
'  Macros
'    SetupPythonSheets  creates the "Python" sheet (code in column A, output
'                       in column C, a Run button) and a "Data" sheet
'    RunPython          runs the program in column A of the "Python" sheet
'    RunSelectedPython  runs the code in the selected cell(s), shows a MsgBox
'
'  Worksheet function
'    =PY(code, [arg1], [arg2], ...)
'        arguments are available as x, y, z and the list args;
'        a pure expression returns its value, a program returns the
'        variable "result" (or what it printed)
'
'  Python built-ins that talk to Excel
'    cell(row, col) / cell("B3")            read a cell
'    set_cell(row, col, value) / set_cell("B3", value)
'    get_range("A1:C10")                    list of rows (flat list for one row/column)
'    set_range("A1", data)                  data: value, list or list of lists
'    clear_range("A1:C10"), sheet_names(), excel_eval("SUM(A1:A9)"), msgbox(text)
'    all of them accept sheet="Name"; the default sheet is "Data" (created
'    when missing); inside PY() it is the sheet of the formula
' ===========================================================================

Private Const CODE_SHEET As String = "Python"
Private Const DATA_SHEET As String = "Data"
Private Const OUT_COLUMN As Long = 3

Private mBusy As Boolean
Private mUdfSheet As Object

' ---------------------------------------------------------------------------
'  macros
' ---------------------------------------------------------------------------
Public Sub RunPython()
    Dim ws As Object, src As String, outText As String, calc As Long, t0 As Single
    If mBusy Then Exit Sub
    Set ws = FindSheet(CODE_SHEET)
    If ws Is Nothing Then
        MsgBox "Sheet '" & CODE_SHEET & "' was not found. Run the macro SetupPythonSheets first.", vbExclamation, "Python"
        Exit Sub
    End If
    src = ReadCode(ws)
    mBusy = True
    Set mUdfSheet = Nothing
    calc = Application.Calculation
    On Error GoTo Fail
    Application.Calculation = -4135      ' xlCalculationManual
    Application.EnableCancelKey = 2      ' xlErrorHandler: Ctrl+Break -> KeyboardInterrupt
    Application.StatusBar = "Running Python..."
    t0 = Timer
    outText = PyRunProgram(src)
Done:
    On Error Resume Next
    Application.Calculation = calc
    Application.EnableCancelKey = 1      ' xlInterrupt
    Application.StatusBar = False
    On Error GoTo 0
    mBusy = False
    WriteOutput ws, outText
    Application.StatusBar = "Python finished in " & Format$(Timer - t0, "0.00") & " s"
    Exit Sub
Fail:
    outText = outText & vbLf & "Internal error " & Err.Number & ": " & Err.Description
    Resume Done
End Sub

Public Sub RunSelectedPython()
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
    outText = PyRunProgram(src)
Done:
    mBusy = False
    If Len(outText) > 900 Then outText = Left$(outText, 900) & vbLf & "..."
    MsgBox outText, vbInformation, "Python output"
    Exit Sub
Fail:
    outText = "Internal error " & Err.Number & ": " & Err.Description
    Resume Done
End Sub

Public Sub SetupPythonSheets()
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
    btn.Caption = "Run Python"
    btn.OnAction = "RunPython"
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
        "# Python subset running inside Excel VBA. Edit column A, then click 'Run Python'.", _
        "def fib(n):", _
        "    a, b = 0, 1", _
        "    for _ in range(n):", _
        "        a, b = b, a + b", _
        "    return a", _
        "", _
        "print('fib:', [fib(i) for i in range(12)])", _
        "", _
        "class Account:", _
        "    def __init__(self, owner, balance=0):", _
        "        self.owner = owner", _
        "        self.balance = balance", _
        "    def deposit(self, amount):", _
        "        if amount <= 0:", _
        "            raise ValueError('amount must be positive')", _
        "        self.balance += amount", _
        "        return self", _
        "    def __repr__(self):", _
        "        return f'Account({self.owner!r}, {self.balance:.2f})'", _
        "", _
        "acc = Account('Ann').deposit(100).deposit(23.5)", _
        "print(acc)", _
        "try:", _
        "    acc.deposit(-1)", _
        "except ValueError as e:", _
        "    print('error:', e)", _
        "", _
        "rows = get_range('A1:B5', sheet='Data')     # header + 4 rows", _
        "total = sum(amount for _, amount in rows[1:])", _
        "print(f'total = {total}')", _
        "words = {name: len(name) for name, _ in rows[1:]}", _
        "print(sorted(words.items(), key=lambda kv: (-kv[1], kv[0])))", _
        "set_cell('D1', 'Total', sheet='Data')", _
        "set_cell('D2', total, sheet='Data')")
End Function

' ---------------------------------------------------------------------------
'  worksheet function
' ---------------------------------------------------------------------------
Public Function PY(ByVal code As String, ParamArray args() As Variant) As Variant
    Dim argv() As Variant, n As Long, i As Long, result As Variant, hasResult As Boolean
    Dim printed As String, errText As String
    If mBusy Then
        PY = "#BUSY"
        Exit Function
    End If
    On Error GoTo Fail
    n = UBound(args) - LBound(args) + 1
    If n > 0 Then
        ReDim argv(0 To n - 1)
        For i = 0 To n - 1
            XlArgToPy args(LBound(args) + i), argv(i)
        Next
    Else
        ReDim argv(0 To 0)
    End If
    mBusy = True
    Set mUdfSheet = Nothing
    If TypeName(Application.Caller) = "Range" Then Set mUdfSheet = Application.Caller.Worksheet
    If PyEvalCode(code, argv, n, result, hasResult, printed, errText) Then
        If hasResult And Not IsEmpty(result) Then
            PY = PyToXlResult(result)
        Else
            If Right$(printed, 1) = vbLf Then printed = Left$(printed, Len(printed) - 1)
            PY = printed
        End If
    Else
        PY = "#" & errText
    End If
    PyFinish
    mBusy = False
    Exit Function
Fail:
    PY = "#Internal error: " & Err.Description
    mBusy = False
End Function

' ---------------------------------------------------------------------------
'  host functions called by the interpreter
' ---------------------------------------------------------------------------
Public Function HostInput(ByVal prompt As String) As String
    HostInput = InputBox(prompt, "Python input()")
End Function

Public Sub HostCall(ByVal code As Long, ByRef args() As Variant, ByVal na As Long, ByVal kw As PyObj, ByRef out As Variant)
    Dim ws As Object, rng As Object, v As Variant, addr As String, i As Long, lst As PyObj
    out = Empty
    Select Case code
    Case BI_X_CELL
        Set ws = HostSheet(kw)
        If na = 1 Then
            Set rng = ws.Range(HostStr(args(0), "cell address"))
        ElseIf na = 2 Then
            Set rng = ws.Cells(IntToLong(args(0), "row"), IntToLong(args(1), "column"))
        Else
            ThrowErr "TypeError", "cell() takes (row, col) or (address)"
        End If
        XlToPy rng.Cells(1, 1).Value2, out
    Case BI_X_SETCELL
        Set ws = HostSheet(kw)
        If na = 2 Then
            Set rng = ws.Range(HostStr(args(0), "cell address"))
            rng.Cells(1, 1).Value = PyToXlScalar(args(1))
        ElseIf na = 3 Then
            Set rng = ws.Cells(IntToLong(args(0), "row"), IntToLong(args(1), "column"))
            rng.Value = PyToXlScalar(args(2))
        Else
            ThrowErr "TypeError", "set_cell() takes (row, col, value) or (address, value)"
        End If
    Case BI_X_GETRANGE
        If na <> 1 Then ThrowErr "TypeError", "get_range() takes exactly one argument (address)"
        Set ws = HostSheet(kw)
        Set rng = ws.Range(HostStr(args(0), "range address"))
        XlValuesToPy rng.Value2, out
    Case BI_X_SETRANGE
        If na <> 2 Then ThrowErr "TypeError", "set_range() takes (address, data)"
        Set ws = HostSheet(kw)
        Set rng = ws.Range(HostStr(args(0), "range address"))
        WritePyData rng.Cells(1, 1), args(1)
    Case BI_X_CLEAR
        If na <> 1 Then ThrowErr "TypeError", "clear_range() takes exactly one argument (address)"
        Set ws = HostSheet(kw)
        ws.Range(HostStr(args(0), "range address")).ClearContents
    Case BI_X_MSGBOX
        If na = 0 Then
            MsgBox ""
        ElseIf na = 1 Then
            MsgBox PyStr(args(0)), vbOKOnly, "Python"
        Else
            MsgBox PyStr(args(0)), vbOKOnly, PyStr(args(1))
        End If
    Case BI_X_EVAL
        If na <> 1 Then ThrowErr "TypeError", "excel_eval() takes exactly one argument (formula)"
        v = Application.Evaluate(HostStr(args(0), "formula"))
        If IsError(v) Then ThrowErr "ValueError", "Excel could not evaluate " & PyRepr(args(0))
        If IsArray(v) Then
            XlValuesToPy v, out
        Else
            XlToPy v, out
        End If
    Case BI_X_SHEETNAMES
        Set lst = NewList()
        For i = 1 To ThisWorkbook.Worksheets.Count
            v = ThisWorkbook.Worksheets(i).Name
            lst.Push v
        Next
        Set out = lst
    Case Else
        ThrowErr "SystemError", "unknown Excel function"
    End Select
End Sub

Private Function HostStr(ByRef v As Variant, ByVal what As String) As String
    If VarType(v) <> vbString Then ThrowErr "TypeError", what & " must be a string such as ""A1"""
    HostStr = v
End Function

Private Function HostSheet(ByVal kw As PyObj) As Object
    Dim v As Variant, ws As Object
    If KwGet(kw, "sheet", v) Then
        If VarType(v) = vbString Then
            Set ws = FindSheet(CStr(v))
            If ws Is Nothing And StrComp(CStr(v), DATA_SHEET, vbTextCompare) = 0 Then Set ws = AddDataSheet()
            If ws Is Nothing Then ThrowErr "ValueError", "no worksheet named " & PyRepr(v)
        ElseIf IsIntType(v) Then
            Set ws = ThisWorkbook.Worksheets(IntToLong(v, "sheet"))
        Else
            ThrowErr "TypeError", "sheet must be a name or an index"
        End If
    ElseIf Not (mUdfSheet Is Nothing) Then
        Set ws = mUdfSheet
    Else
        ' default: the "Data" sheet (created on demand, so that a program never
        ' overwrites its own code on the "Python" sheet by accident)
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
    If lastRow = 1 And Len(CStr(ws.Cells(1, 1).Formula)) = 0 Then Exit Function
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

Private Sub WriteOutput(ByVal ws As Object, ByVal text As String)
    Dim lines() As String, n As Long, arr() As Variant, i As Long, s As String
    ws.Columns(OUT_COLUMN).ClearContents
    If Right$(text, 1) = vbLf Then text = Left$(text, Len(text) - 1)
    If Len(text) = 0 Then Exit Sub
    lines = Split(text, vbLf)
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
'  value conversion Excel <-> Python
' ---------------------------------------------------------------------------
Public Sub XlToPy(ByVal v As Variant, ByRef out As Variant)
    Select Case VarType(v)
    Case vbDouble, vbSingle, vbCurrency
        If v = Int(v) And Abs(v) <= 2147483647# Then
            out = CLng(v)
        Else
            out = CDbl(v)
        End If
    Case vbString, vbBoolean, vbLong, vbInteger
        out = v
    Case vbDate
        out = Format$(v, "yyyy-mm-dd hh:nn:ss")
        If Right$(out, 9) = " 00:00:00" Then out = Left$(out, 10)
    Case vbEmpty, vbNull
        out = Empty
    Case vbError
        out = "#ERROR"
    Case vbDecimal
        out = NormBig(v)
    Case Else
        out = CStr(v)
    End Select
End Sub

Private Sub XlValuesToPy(ByVal v As Variant, ByRef out As Variant)
    Dim r As Long, c As Long, r0 As Long, r1 As Long, c0 As Long, c1 As Long
    Dim rows As PyObj, row As PyObj, x As Variant
    If Not IsArray(v) Then
        XlToPy v, out
        Exit Sub
    End If
    r0 = LBound(v, 1): r1 = UBound(v, 1)
    c0 = LBound(v, 2): c1 = UBound(v, 2)
    Set rows = NewList()
    If r0 = r1 Or c0 = c1 Then
        For r = r0 To r1
            For c = c0 To c1
                XlToPy v(r, c), x
                rows.Push x
            Next
        Next
    Else
        For r = r0 To r1
            Set row = NewList()
            For c = c0 To c1
                XlToPy v(r, c), x
                row.Push x
            Next
            Set x = row
            rows.Push x
        Next
    End If
    Set out = rows
End Sub

Private Sub XlArgToPy(ByVal a As Variant, ByRef out As Variant)
    Dim rg As Object
    If IsObject(a) Then
        Set rg = a
        If TypeName(rg) = "Range" Then
            If rg.Cells.Count = 1 Then
                XlToPy rg.Value2, out
            Else
                XlValuesToPy rg.Value2, out
            End If
        Else
            out = TypeName(rg)
        End If
        Exit Sub
    End If
    If IsArray(a) Then
        If ArrayDims(a) = 2 Then
            XlValuesToPy a, out
        Else
            XlValuesToPy Arr1To2(a), out
        End If
        Exit Sub
    End If
    XlToPy a, out
End Sub

Private Function ArrayDims(ByVal a As Variant) As Long
    Dim d As Long, t As Long
    On Error GoTo Done
    For d = 1 To 60
        t = UBound(a, d)
        ArrayDims = d
    Next
Done:
End Function

Private Function Arr1To2(ByVal a As Variant) As Variant
    Dim r() As Variant, i As Long
    ReDim r(1 To 1, LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        r(1, i) = a(i)
    Next
    Arr1To2 = r
End Function

Public Function PyToXlScalar(ByRef v As Variant) As Variant
    Select Case VarType(v)
    Case vbEmpty, vbNull
        PyToXlScalar = Empty
    Case vbDecimal
        PyToXlScalar = CDbl(v)
    Case vbLong, vbInteger, vbDouble, vbBoolean, vbSingle, vbCurrency
        PyToXlScalar = v
    Case vbString
        PyToXlScalar = v
    Case Else
        PyToXlScalar = PyStr(v)
    End Select
End Function

Private Function PyToXlResult(ByRef v As Variant) As Variant
    Dim o As PyObj, arr() As Variant, i As Long, j As Long, nr As Long, nc As Long, x As Variant, row As PyObj
    Dim k As Variant, is2D As Boolean
    If Not IsObject(v) Then
        If IsEmpty(v) Then
            PyToXlResult = ""
        Else
            PyToXlResult = PyToXlScalar(v)
        End If
        Exit Function
    End If
    Set o = v
    Select Case o.Kind
    Case OK_LIST, OK_TUPLE
        nr = o.Count
        If nr = 0 Then
            PyToXlResult = ""
            Exit Function
        End If
        is2D = True
        nc = 0
        For i = 0 To nr - 1
            o.GetInto i, x
            If IsKind(x, OK_LIST) Or IsKind(x, OK_TUPLE) Then
                Set row = x
                If row.Count > nc Then nc = row.Count
            Else
                is2D = False
            End If
        Next
        If Not is2D Or nc = 0 Then nc = 1
        ReDim arr(0 To nr - 1, 0 To nc - 1)
        For i = 0 To nr - 1
            o.GetInto i, x
            If is2D Then
                Set row = x
                For j = 0 To nc - 1
                    If j < row.Count Then
                        row.GetInto j, k
                        arr(i, j) = PyToXlScalar(k)
                    Else
                        arr(i, j) = ""
                    End If
                Next
            Else
                arr(i, 0) = PyToXlScalar(x)
            End If
        Next
        PyToXlResult = arr
    Case OK_DICT
        If o.DCount = 0 Then
            PyToXlResult = ""
            Exit Function
        End If
        ReDim arr(0 To o.DCount - 1, 0 To 1)
        j = 0
        For i = 0 To o.EntryCount - 1
            If o.LiveAt(i) Then
                o.KeyInto i, k
                o.ValInto i, x
                arr(j, 0) = PyToXlScalar(k)
                arr(j, 1) = PyToXlScalar(x)
                j = j + 1
            End If
        Next
        PyToXlResult = arr
    Case Else
        PyToXlResult = PyStr(v)
    End Select
End Function

Private Sub WritePyData(ByVal topLeft As Object, ByRef data As Variant)
    Dim o As PyObj, nr As Long, nc As Long, i As Long, j As Long, x As Variant, row As PyObj, arr() As Variant
    Dim is2D As Boolean, k As Variant
    If Not (IsKind(data, OK_LIST) Or IsKind(data, OK_TUPLE)) Then
        topLeft.Value = PyToXlScalar(data)
        Exit Sub
    End If
    Set o = data
    nr = o.Count
    If nr = 0 Then Exit Sub
    is2D = True
    For i = 0 To nr - 1
        o.GetInto i, x
        If IsKind(x, OK_LIST) Or IsKind(x, OK_TUPLE) Then
            Set row = x
            If row.Count > nc Then nc = row.Count
        Else
            is2D = False
        End If
    Next
    If Not is2D Or nc = 0 Then nc = 1
    ReDim arr(1 To nr, 1 To nc)
    For i = 0 To nr - 1
        o.GetInto i, x
        If is2D Then
            Set row = x
            For j = 0 To row.Count - 1
                row.GetInto j, k
                arr(i + 1, j + 1) = PyToXlScalar(k)
            Next
        Else
            arr(i + 1, 1) = PyToXlScalar(x)
        End If
    Next
    topLeft.Resize(nr, nc).Value = arr
End Sub
