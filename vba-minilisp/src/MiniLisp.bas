Attribute VB_Name = "MiniLisp"
'==============================================================
' MiniLisp —— 写给初学者的 Lisp 解释器（Excel VBA）
'
' 整个解释器只有三步，反复执行：
'
'     源代码 --Tokenize--> 记号 --ReadExpr--> 表达式 --Eval--> 值 --ToText--> 文字
'     "(+ 1 2)"            ( + 1 2 )        列表 (+ 1 2)        3              "3"
'
' 本文件从上到下依次是：
'   第 0 部分  数据怎么表示；序对（cons）和“内存”
'   第 1 部分  在 Excel 里使用（入口）
'   第 2 部分  Tokenize   —— 词法分析：把文字切成记号
'   第 3 部分  ReadExpr   —— 语法分析：把记号变成列表
'   第 4 部分  Eval       —— 求值（解释器的心脏）
'   第 5 部分  Apply      —— 调用函数
'   第 6 部分  内置函数    —— 想加新函数就改这里
'   第 7 部分  ToText     —— 把值变回文字
'   第 8 部分  小工具函数
'
' 另一个文件 LispEnv.cls 是“环境”（变量表）。
' 扩展方法和练习见 README.md。
'==============================================================
Option Explicit

'--------------------------------------------------------------
' 第 0 部分  数据怎么表示
'
' 所有 Lisp 值都放在 VBA 的 Variant 里：
'
'   Lisp 里写的      VBA 里存的                         说明
'   42   3.5         Double                             数字统一用 Double
'   #t   #f          Boolean
'   x    +   list?   String                             符号（变量名）
'   "hi"             Array("#str", "hi")                字符串，带 #str 标签
'   (1 . 2)  (1 2)   Long：序对的编号（见下面）           序对、列表
'   ()               Long 0（变量 NIL）                  空表
'   (lambda ...)     Array("#lambda", 参数, 函数体, 环境) 用户定义的函数
'   +  car ...       Array("#builtin", "+")             内置函数
'   (没有值)          Empty                              define、display 的返回值
'
' ---- 序对和“内存” ----
' Lisp 的一切列表都由“序对”(pair，也叫 cons cell) 搭成。一个序对有两格：
' car 和 cdr。所有序对都放在两个大数组 gCar、gCdr 里：
' 第 i 号序对的 car 是 gCar(i)，cdr 是 gCdr(i)。
' 一个序对在 VBA 里就是它的编号 i（Long 类型），就像 C 语言里的指针。
' 编号 0 不用，留给空表 ()。
'
'   列表 (1 2 3) 其实是 (1 . (2 . (3 . ())))，在“内存”里是这样的：
'
'       编号      1     2     3
'       gCar      1     2     3
'       gCdr      2     3     0        <- 0 就是 ()
'
'   所以 (1 2 3) 这个值就是 Long 1：从 1 号序对出发，顺着 cdr 一直走到 0。
'   (cons 0 那个列表) 只需要新造一个序对：gCar = 0，gCdr = 1。
'
' 因为数字都是 Double，Long 只用来表示序对，两者不会混淆。
' 带标签的数组第一个元素是以 # 开头的字符串；普通符号不可能以 # 开头
' （读入时会报错），所以标签也不会和别的值混淆。
'--------------------------------------------------------------

Private Const LISP_ERROR As Long = -2147220504   ' = vbObjectError + 1000，MiniLisp 自己的错误号

Private NIL As Long             ' 空表 ()，值永远是 0。（不用 Const 是因为 LibreOffice 会把 Const 0 当成 Integer）
Private gCar() As Variant       ' 第 i 号序对的 car
Private gCdr() As Variant       ' 第 i 号序对的 cdr
Private gPairCount As Long      ' 已经用掉了几个序对
Private gMemoryReady As Boolean ' gCar / gCdr 是否已经建好

Private gGlobal As LispEnv      ' 全局环境（RunSheet / LispRepl 共用）
Private gOut As String          ' display 输出的文字先攒在这里

' 清空“内存”：之前造的序对全部作废
Private Sub ResetMemory()
    ReDim gCar(0 To 1023)
    ReDim gCdr(0 To 1023)
    gPairCount = 0
    gMemoryReady = True
End Sub

' 造一个新序对，返回它的编号
Private Function Cons(ByVal a As Variant, ByVal d As Variant) As Long
    If gPairCount >= UBound(gCar) Then
        ' 数组用完了：容量翻倍
        ReDim Preserve gCar(0 To UBound(gCar) * 2)
        ReDim Preserve gCdr(0 To UBound(gCdr) * 2)
    End If
    gPairCount = gPairCount + 1
    gCar(gPairCount) = a
    gCdr(gPairCount) = d
    Cons = gPairCount
End Function

Private Function Car(ByVal p As Variant) As Variant
    If Not IsPair(p) Then LispError "car: expected a pair, got " & ToText(p, True)
    Car = gCar(p)
End Function

Private Function Cdr(ByVal p As Variant) As Variant
    If Not IsPair(p) Then LispError "cdr: expected a pair, got " & ToText(p, True)
    Cdr = gCdr(p)
End Function

Private Sub SetCar(ByVal p As Variant, ByVal v As Variant)
    If Not IsPair(p) Then LispError "set-car!: expected a pair, got " & ToText(p, True)
    gCar(p) = v
End Sub

Private Sub SetCdr(ByVal p As Variant, ByVal v As Variant)
    If Not IsPair(p) Then LispError "set-cdr!: expected a pair, got " & ToText(p, True)
    gCdr(p) = v
End Sub

Private Function IsPair(ByVal v As Variant) As Boolean
    If VarType(v) = vbLong Then IsPair = (v <> NIL)
End Function

Private Function IsNil(ByVal v As Variant) As Boolean
    If VarType(v) = vbLong Then IsNil = (v = NIL)
End Function

' 把编号大于 mark 的序对还回去（=LISP() 用完就扔，不占内存）
Private Sub FreePairs(ByVal mark As Long)
    Dim i As Long
    For i = mark + 1 To gPairCount
        gCar(i) = Empty
        gCdr(i) = Empty
    Next
    gPairCount = mark
End Sub

'==============================================================
' 第 1 部分  在 Excel 里使用
'==============================================================

' 运行一整段程序，返回 display 输出的文字。
' showValues = True 时像 REPL 一样，把每个表达式的值也打印出来。
' 每次运行都从一个全新的全局环境开始。
Public Function LispRun(ByVal src As String, Optional ByVal showValues As Boolean = False) As String
    On Error GoTo Failed
    LispReset
    LispRun = RunAndShow(src, gGlobal, showValues)
    Exit Function
Failed:
    LispRun = "Error: " & Err.Description & vbLf
End Function

' 工作表函数：在单元格里写  =LISP("(+ 1 2)")  得到 3
' 返回最后一个表达式的值；如果它没有值，就返回 display 输出的文字。
' 每次调用都使用独立的新环境，不会影响 RunSheet 的变量。
Public Function LISP(ByVal code As String) As Variant
    Dim env As LispEnv, v As Variant, savedOut As String, mark As Long
    If Not gMemoryReady Then ResetMemory
    mark = gPairCount               ' 记住现在用到哪个序对，结束时把新造的都还回去
    savedOut = gOut
    gOut = ""
    On Error GoTo Failed
    Set env = NewGlobalEnv()
    v = EvalAll(code, env, False)
    If VarType(v) = vbDouble Or VarType(v) = vbBoolean Then
        LISP = v
    ElseIf IsEmpty(v) Then
        LISP = gOut
    Else
        LISP = ToText(v, False)
    End If
    env.Clear
    FreePairs mark
    gOut = savedOut
    Exit Function
Failed:
    LISP = "Error: " & Err.Description
    If Not (env Is Nothing) Then env.Clear
    FreePairs mark
    gOut = savedOut
End Function

' 宏：把当前工作表 A 列当作“笔记本”，每个单元格写一段代码，
' 从上到下依次运行，结果写到同一行的 B 列。前面定义的变量后面可以接着用。
Public Sub RunSheet()
    Dim ws As Object, lastRow As Long, r As Long, code As String, result As String
    Set ws = ActiveSheet
    LispReset
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row     ' -4162 = xlUp
    For r = 1 To lastRow
        If IsError(ws.Cells(r, 1).Value) Then
            code = ""
        Else
            code = CStr(ws.Cells(r, 1).Value)
        End If
        ws.Cells(r, 2).NumberFormat = "@"                   ' 按文本显示结果
        If Trim$(code) = "" Then
            ws.Cells(r, 2).Value = ""
        Else
            result = RunAndShow(code, gGlobal, True)
            If Right$(result, 1) = vbLf Then result = Left$(result, Len(result) - 1)
            ws.Cells(r, 2).Value = result
        End If
    Next
End Sub

' 宏：弹出输入框的交互式 REPL。点“取消”或什么都不输入就退出。
Public Sub LispRepl()
    Dim code As String, shown As String
    LispReset
    shown = "MiniLisp REPL. Type an expression, e.g. (+ 1 2)." & vbLf & "Cancel to quit."
    Do
        code = InputBox(shown, "MiniLisp")
        If Trim$(code) = "" Then Exit Do
        shown = "> " & code & vbLf & RunAndShow(code, gGlobal, True)
    Loop
End Sub

' 重新开始：清空内存，丢掉所有用户定义，建立新的全局环境
Public Sub LispReset()
    If Not (gGlobal Is Nothing) Then gGlobal.Clear
    ResetMemory
    Set gGlobal = NewGlobalEnv()
End Sub

' 运行 src，返回输出文字；出错时把错误信息也放进输出里（不会弹窗）
Private Function RunAndShow(ByVal src As String, ByVal env As LispEnv, ByVal showValues As Boolean) As String
    Dim v As Variant
    gOut = ""
    On Error GoTo Failed
    v = EvalAll(src, env, showValues)
    RunAndShow = gOut
    Exit Function
Failed:
    FreshLine
    RunAndShow = gOut & "Error: " & Err.Description & vbLf
End Function

' Read-Eval-Print 循环：读一个表达式 -> 求值 ->（可选）打印 -> 读下一个……
' 返回最后一个表达式的值
Private Function EvalAll(ByVal src As String, ByVal env As LispEnv, ByVal showValues As Boolean) As Variant
    Dim toks As Collection, pos As Long, expr As Variant, v As Variant
    Set toks = Tokenize(src)
    pos = 1
    Do While pos <= toks.Count
        expr = ReadExpr(toks, pos)              ' Read
        v = Eval(expr, env)                     ' Eval
        If showValues And Not IsEmpty(v) Then   ' Print
            FreshLine
            gOut = gOut & ToText(v, True) & vbLf
        End If
    Loop
    EvalAll = v
End Function

' 建立一个新的全局环境：登记所有内置函数，再运行用 Lisp 写的“序言”
Private Function NewGlobalEnv() As LispEnv
    Dim env As LispEnv, nm As Variant
    Set env = New LispEnv
    For Each nm In Split(BuiltinNames(), " ")
        env.Define nm, Array("#builtin", nm)
    Next
    EvalAll Prelude(), env, False
    Set NewGlobalEnv = env
End Function

' 序言：用 Lisp 自己定义的函数。
' 很多功能不必用 VBA 写，直接在这里用 Lisp 写就行（见 README“用 Lisp 扩展 Lisp”）。
Private Function Prelude() As String
    Dim p As String
    p = p & "(define (cadr x) (car (cdr x)))" & vbLf
    p = p & "(define (map f lst)" & vbLf
    p = p & "  (if (null? lst) '()" & vbLf
    p = p & "      (cons (f (car lst)) (map f (cdr lst)))))" & vbLf
    p = p & "(define (filter keep? lst)" & vbLf
    p = p & "  (cond ((null? lst) '())" & vbLf
    p = p & "        ((keep? (car lst)) (cons (car lst) (filter keep? (cdr lst))))" & vbLf
    p = p & "        (else (filter keep? (cdr lst)))))" & vbLf
    p = p & "(define (reverse lst)" & vbLf
    p = p & "  (define (loop lst acc)" & vbLf
    p = p & "    (if (null? lst) acc (loop (cdr lst) (cons (car lst) acc))))" & vbLf
    p = p & "  (loop lst '()))" & vbLf
    Prelude = p
End Function

'==============================================================
' 第 2 部分  Tokenize —— 把源代码切成“记号”
'
'   "(define x (+ 1 2))"   ->   (   define   x   (   +   1   2   )   )
'
' 规则：( ) ' 各自是一个记号；空白分隔记号；; 到行尾是注释；
'       "..." 是字符串记号，存成：一个双引号 + 内容（转义已处理）。
'==============================================================
Private Function Tokenize(ByVal src As String) As Collection
    Dim toks As New Collection
    Dim i As Long, n As Long, c As String, tok As String
    n = Len(src)
    i = 1
    Do While i <= n
        c = Mid$(src, i, 1)
        If c = "(" Or c = ")" Or c = "'" Then
            ' 括号和引号：单独成为一个记号
            toks.Add c
            i = i + 1
        ElseIf c = ";" Then
            ' 注释：一直跳到行尾
            Do While i <= n
                If Mid$(src, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf c = """" Then
            ' 字符串：读到下一个双引号为止
            tok = """"
            i = i + 1
            Do
                If i > n Then LispError "unterminated string"
                c = Mid$(src, i, 1)
                If c = """" Then Exit Do
                If c = "\" And i < n Then
                    ' 转义字符：\n 换行，\t 制表符，\" 双引号，\\ 反斜杠
                    i = i + 1
                    c = Mid$(src, i, 1)
                    If c = "n" Then c = vbLf
                    If c = "t" Then c = vbTab
                End If
                tok = tok & c
                i = i + 1
            Loop
            toks.Add tok
            i = i + 1
        ElseIf IsSpace(c) Then
            ' 空白：跳过
            i = i + 1
        ElseIf c = ChrW(&HFF08) Or c = ChrW(&HFF09) Then
            ' 中文输入法打出的全角括号，初学者很容易写错
            LispError "full-width parenthesis found, please use ( and )"
        Else
            ' 其它：一直读到空白或括号为止，比如 define、42、list?、.
            tok = ""
            Do While i <= n
                c = Mid$(src, i, 1)
                If IsSpace(c) Or InStr("()'"";", c) > 0 Then Exit Do
                tok = tok & c
                i = i + 1
            Loop
            toks.Add tok
        End If
    Loop
    Set Tokenize = toks
End Function

Private Function IsSpace(ByVal c As String) As Boolean
    IsSpace = (c = " " Or c = vbTab Or c = vbCr Or c = vbLf Or c = ChrW(160) Or c = ChrW(&H3000))
End Function

'==============================================================
' 第 3 部分  ReadExpr —— 把记号变成表达式
'
'   ( + 1 ( * 2 3 ) )   ->   列表 (+ 1 (* 2 3))，由序对搭成
'
' 从第 pos 个记号开始读出“一个”完整的表达式，读完后 pos 指向下一个记号。
' 遇到 ( 就交给 ReadListRest 一个个读元素，直到 ) 为止。
' 元素本身又可能是列表 —— 递归天然地处理了嵌套。
'==============================================================
Private Function ReadExpr(ByVal toks As Collection, ByRef pos As Long) As Variant
    Dim tok As String
    If pos > toks.Count Then LispError "unexpected end of input"
    tok = toks.Item(pos)
    pos = pos + 1
    If tok = "(" Then
        ReadExpr = ReadListRest(toks, pos)
    ElseIf tok = ")" Then
        LispError "unexpected ')'"
    ElseIf tok = "'" Then
        ' 'x 是 (quote x) 的简写
        ReadExpr = Cons("quote", Cons(ReadExpr(toks, pos), NIL))
    Else
        ReadExpr = ParseAtom(tok)
    End If
End Function

' 已经读过了 "("，接着读列表剩下的部分，直到 ")"：
'   a b c )      ->  (a b c)   也就是 (cons a (cons b (cons c '())))
'   a b . c )    ->  (a b . c) 点号后面的就是最后一个 cdr
Private Function ReadListRest(ByVal toks As Collection, ByRef pos As Long) As Variant
    Dim tok As String, first As Variant
    If pos > toks.Count Then LispError "missing ')'"
    tok = toks.Item(pos)
    If tok = ")" Then
        pos = pos + 1
        ReadListRest = NIL
    ElseIf tok = "." Then
        pos = pos + 1
        ReadListRest = ReadExpr(toks, pos)
        If pos > toks.Count Then LispError "missing ')'"
        If toks.Item(pos) <> ")" Then LispError "bad dotted list"
        pos = pos + 1
    Else
        first = ReadExpr(toks, pos)                             ' 先读第一个元素
        ReadListRest = Cons(first, ReadListRest(toks, pos))     ' 再读剩下的，接在后面
    End If
End Function

' 单个记号 -> 值：字符串、#t/#f、数字，其余都是符号
Private Function ParseAtom(ByVal tok As String) As Variant
    If Left$(tok, 1) = """" Then
        ParseAtom = MakeStr(Mid$(tok, 2))
    ElseIf tok = "#t" Then
        ParseAtom = True
    ElseIf tok = "#f" Then
        ParseAtom = False
    ElseIf IsNumberToken(tok) Then
        ParseAtom = CDbl(Val(tok))
    ElseIf Left$(tok, 1) = "#" Then
        LispError "unknown syntax: " & tok
    Else
        ParseAtom = tok
    End If
End Function

' 像 42、-7、3.14、.5 这样的记号才算数字（+ - . 单独出现是符号）
Private Function IsNumberToken(ByVal tok As String) As Boolean
    Dim i As Long, c As String, digits As Long, dots As Long
    For i = 1 To Len(tok)
        c = Mid$(tok, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits + 1
        ElseIf c = "." Then
            dots = dots + 1
        ElseIf (c = "-" Or c = "+") And i = 1 Then
            ' 开头的正负号，允许
        Else
            Exit Function           ' 其它字符：不是数字（返回 False）
        End If
    Next
    IsNumberToken = (digits > 0 And dots <= 1)
End Function

'==============================================================
' 第 4 部分  Eval —— 求值，解释器的心脏
'
' 只有三种情况：
'   ① 符号        -> 去环境里查它的值
'   ② 数字等       -> 值就是它自己
'   ③ 列表 (f ...) -> 如果 f 是特殊形式（if、define……）按各自的规则处理；
'                     否则是函数调用：先求值每个元素，再交给 Apply
'
' 表达式 x 是一个列表：Car(x) 是第一个元素，Nth(x, 1) 是第二个，
' Drop(x, 2) 是去掉前两个之后剩下的列表。
'==============================================================
Public Function Eval(ByVal x As Variant, ByVal env As LispEnv) As Variant
    ' ① 符号
    If IsSymbol(x) Then
        Eval = LookupVar(x, env)
        Exit Function
    End If

    ' ② 不是序对：数字、布尔、字符串、函数……
    If Not IsPair(x) Then
        If IsNil(x) Then LispError "cannot evaluate an empty list ()"
        Eval = x
        Exit Function
    End If

    ' ③ 列表
    Select Case HeadName(x)
    Case "quote":  Eval = EvalQuote(x, env)
    Case "if":     Eval = EvalIf(x, env)
    Case "define": Eval = EvalDefine(x, env)
    Case "set!":   Eval = EvalSet(x, env)
    Case "lambda": Eval = EvalLambda(x, env)
    Case "begin":  Eval = EvalBody(Cdr(x), env)
    Case "let":    Eval = EvalLet(x, env)
    Case "cond":   Eval = EvalCond(x, env)
    Case "and":    Eval = EvalAnd(x, env)
    Case "or":     Eval = EvalOr(x, env)
    Case Else:     Eval = EvalCall(x, env)      ' 不是特殊形式 -> 函数调用
    End Select
End Function

' 查变量：从当前环境一层层往外找
Private Function LookupVar(ByVal nm As String, ByVal env As LispEnv) As Variant
    Dim e As LispEnv
    Set e = env.Find(nm)
    If e Is Nothing Then LispError "unbound variable: " & nm
    LookupVar = e.GetValue(nm)
End Function

' (f a b c)：先算出 f 和每个参数的值，再交给 Apply
Private Function EvalCall(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim f As Variant, args As Variant
    f = Eval(Car(x), env)
    args = EvalArgs(Cdr(x), env)
    EvalCall = Apply(f, args)
End Function

' 对列表里的每个表达式求值，结果放进一个 VBA 数组。
' （参数用 VBA 数组而不是 Lisp 列表，这样每次调用函数都不用造新的序对）
Private Function EvalArgs(ByVal exprs As Variant, ByVal env As LispEnv) As Variant
    Dim r() As Variant, i As Long, n As Long
    n = ListLength(exprs)
    If n < 0 Then LispError "cannot evaluate a dotted list"
    If n = 0 Then
        EvalArgs = Array()
        Exit Function
    End If
    ReDim r(0 To n - 1)
    For i = 0 To n - 1
        r(i) = Eval(Car(exprs), env)
        exprs = Cdr(exprs)
    Next
    EvalArgs = r
End Function

' 依次求值列表里的多个表达式，返回最后一个的值（函数体、begin、let 都用它）
Private Function EvalBody(ByVal body As Variant, ByVal env As LispEnv) As Variant
    Dim v As Variant
    Do While IsPair(body)
        v = Eval(Car(body), env)
        body = Cdr(body)
    Loop
    EvalBody = v
End Function

'---------------- 特殊形式 ----------------
' 特殊形式和函数调用的区别：参数“不会”事先全部求值，
' 由各自的规则决定求值哪些、什么时候求值。

' (quote x)：不求值，原样返回 x
Private Function EvalQuote(ByVal x As Variant, ByVal env As LispEnv) As Variant
    CheckForm x, 2, 2, "(quote x)"
    EvalQuote = Nth(x, 1)
End Function

' (if 条件 是则 [否则])：只有 #f 算假，其它值（包括 0 和空表）都算真
Private Function EvalIf(ByVal x As Variant, ByVal env As LispEnv) As Variant
    CheckForm x, 3, 4, "(if test then [else])"
    If IsTrue(Eval(Nth(x, 1), env)) Then
        EvalIf = Eval(Nth(x, 2), env)
    ElseIf ListLength(x) = 4 Then
        EvalIf = Eval(Nth(x, 3), env)
    Else
        EvalIf = Empty
    End If
End Function

' (define 名字 值)  或  (define (函数名 参数...) 函数体...)
' 第二种写法只是 (define 函数名 (lambda (参数...) 函数体...)) 的简写
Private Function EvalDefine(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim target As Variant
    CheckForm x, 3, -1, "(define name value) or (define (name params...) body...)"
    target = Nth(x, 1)
    If IsPair(target) Then
        CheckSymbol Car(target), "define"
        env.Define Car(target), MakeLambda(Cdr(target), Drop(x, 2), env)
    Else
        CheckSymbol target, "define"
        CheckForm x, 3, 3, "(define name value)"
        env.Define target, Eval(Nth(x, 2), env)
    End If
    EvalDefine = Empty
End Function

' (set! 名字 新值)：修改一个已经存在的变量（它可能在外层环境里）
Private Function EvalSet(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim e As LispEnv, nm As Variant
    CheckForm x, 3, 3, "(set! name value)"
    nm = Nth(x, 1)
    CheckSymbol nm, "set!"
    Set e = env.Find(nm)
    If e Is Nothing Then LispError "set!: unbound variable: " & nm
    e.Define nm, Eval(Nth(x, 2), env)
    EvalSet = Empty
End Function

' (lambda (参数...) 函数体...)：造一个函数，并记住“现在所在的环境”
Private Function EvalLambda(ByVal x As Variant, ByVal env As LispEnv) As Variant
    CheckForm x, 3, -1, "(lambda (params...) body...)"
    EvalLambda = MakeLambda(Nth(x, 1), Drop(x, 2), env)
End Function

' 参数表可以是 (a b)、(a b . rest)，或者单个符号 args（收集全部参数）
Private Function MakeLambda(ByVal params As Variant, ByVal body As Variant, ByVal env As LispEnv) As Variant
    Dim p As Variant
    p = params
    Do While IsPair(p)
        CheckSymbol Car(p), "lambda"
        p = Cdr(p)
    Loop
    If Not IsNil(p) Then CheckSymbol p, "lambda"
    MakeLambda = Array("#lambda", params, body, env)
End Function

' (let ((a 1) (b 2)) 函数体...)
' 新建一个环境放 a、b，再在里面求值函数体。初值是在“外面”的环境里算的。
Private Function EvalLet(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim e As LispEnv, bindings As Variant, b As Variant
    CheckForm x, 3, -1, "(let ((name value) ...) body...)"
    bindings = Nth(x, 1)
    If ListLength(bindings) < 0 Then LispError "let: bad bindings " & ToText(bindings, True)
    Set e = New LispEnv
    Set e.Parent = env
    Do While IsPair(bindings)
        b = Car(bindings)
        If ListLength(b) <> 2 Then LispError "let: bad binding " & ToText(b, True)
        CheckSymbol Car(b), "let"
        e.Define Car(b), Eval(Nth(b, 1), env)
        bindings = Cdr(bindings)
    Loop
    EvalLet = EvalBody(Drop(x, 2), e)
End Function

' (cond (条件1 表达式...) (条件2 表达式...) (else 表达式...))
' 从上到下找第一个为真的条件，求值它后面的表达式
Private Function EvalCond(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim clauses As Variant, clause As Variant, test As Variant
    clauses = Cdr(x)
    Do While IsPair(clauses)
        clause = Car(clauses)
        If ListLength(clause) < 1 Then LispError "cond: bad clause " & ToText(clause, True)
        If HeadName(clause) = "else" Then
            test = True
        Else
            test = Eval(Car(clause), env)
        End If
        If IsTrue(test) Then
            If IsNil(Cdr(clause)) Then
                EvalCond = test
            Else
                EvalCond = EvalBody(Cdr(clause), env)
            End If
            Exit Function
        End If
        clauses = Cdr(clauses)
    Loop
    EvalCond = Empty
End Function

' (and a b ...)：从左到右求值，遇到 #f 就停下返回 #f；全为真则返回最后一个值
Private Function EvalAnd(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim items As Variant, v As Variant
    v = True
    items = Cdr(x)
    Do While IsPair(items)
        v = Eval(Car(items), env)
        If Not IsTrue(v) Then Exit Do
        items = Cdr(items)
    Loop
    EvalAnd = v
End Function

' (or a b ...)：从左到右求值，遇到第一个真值就停下返回它；全为假返回 #f
Private Function EvalOr(ByVal x As Variant, ByVal env As LispEnv) As Variant
    Dim items As Variant, v As Variant
    v = False
    items = Cdr(x)
    Do While IsPair(items)
        v = Eval(Car(items), env)
        If IsTrue(v) Then Exit Do
        items = Cdr(items)
    Loop
    EvalOr = v
End Function

'==============================================================
' 第 5 部分  Apply —— 调用函数
'
' args 是已经求好值的参数，放在一个 VBA 数组里。
' 内置函数：交给 CallBuiltin。
' 用户函数：新建一个环境，它的外层(Parent)是函数“定义时”的环境，
'           把参数放进去，再在里面求值函数体。
'           正因为外层是“定义时”而不是“调用时”的环境，函数才能
'           记住它出生地的变量 —— 这就是闭包。
'==============================================================
Public Function Apply(ByVal f As Variant, ByVal args As Variant) As Variant
    Dim e As LispEnv
    Select Case TagOf(f)
    Case "#builtin"
        Apply = CallBuiltin(f(1), args)
    Case "#lambda"
        ' f = Array("#lambda", 参数, 函数体, 定义时的环境)
        Set e = New LispEnv
        Set e.Parent = f(3)
        BindParams e, f(1), args
        Apply = EvalBody(f(2), e)
    Case Else
        LispError "not a procedure: " & ToText(f, True)
    End Select
End Function

' 把实参放进新环境：参数表 (a b) 对应两个实参；
' (a b . rest) 时多出来的实参打包成列表放进 rest；单个符号 args 收集全部实参
Private Sub BindParams(ByVal e As LispEnv, ByVal params As Variant, ByVal args As Variant)
    Dim i As Long, n As Long, p As Variant
    n = UBound(args) + 1
    p = params
    Do While IsPair(p)
        If i >= n Then ArityError params, n
        e.Define Car(p), args(i)
        i = i + 1
        p = Cdr(p)
    Loop
    If IsSymbol(p) Then
        e.Define p, ArrayToList(args, i)
    ElseIf i < n Then
        ArityError params, n
    End If
End Sub

Private Sub ArityError(ByVal params As Variant, ByVal got As Long)
    Dim need As Long, p As Variant
    p = params
    Do While IsPair(p)
        need = need + 1
        p = Cdr(p)
    Loop
    If IsSymbol(p) Then
        LispError "procedure expects at least " & need & " argument(s), got " & got
    Else
        LispError "procedure expects " & need & " argument(s), got " & got
    End If
End Sub

'==============================================================
' 第 6 部分  内置函数
'
' 添加一个内置函数只要两步：
'   1. 在 BuiltinNames 里加上名字；
'   2. 在 CallBuiltin 里加一个 Case，算出结果放进 r。
' args 是已经求好值的参数（VBA 数组，从 0 开始），n 是参数个数。
'==============================================================
Private Function BuiltinNames() As String
    BuiltinNames = "+ - * / remainder abs min max sqrt" _
        & " = < > <= >= not" _
        & " cons car cdr set-car! set-cdr! list length append apply" _
        & " null? pair? list? number? symbol? string? boolean? procedure? eq? equal?" _
        & " string-append string-length number->string" _
        & " display newline error" _
        & " cell set-cell!"
End Function

Private Function CallBuiltin(ByVal nm As String, ByVal args As Variant) As Variant
    Dim n As Long, i As Long, acc As Double, r As Variant, s As String
    n = UBound(args) + 1
    Select Case nm

    '---------- 算术 ----------
    Case "+"
        acc = 0
        For i = 0 To n - 1
            acc = acc + Num(args(i))
        Next
        r = acc
    Case "*"
        acc = 1
        For i = 0 To n - 1
            acc = acc * Num(args(i))
        Next
        r = acc
    Case "-"
        NeedArgs nm, n, 1, -1
        If n = 1 Then
            r = -Num(args(0))
        Else
            acc = Num(args(0))
            For i = 1 To n - 1
                acc = acc - Num(args(i))
            Next
            r = acc
        End If
    Case "/"
        NeedArgs nm, n, 1, -1
        If n = 1 Then
            acc = 1
            i = 0
        Else
            acc = Num(args(0))
            i = 1
        End If
        Do While i < n
            If Num(args(i)) = 0 Then LispError "/: division by zero"
            acc = acc / Num(args(i))
            i = i + 1
        Loop
        r = acc
    Case "remainder"
        NeedArgs nm, n, 2, 2
        If Num(args(1)) = 0 Then LispError "remainder: division by zero"
        r = Num(args(0)) - Num(args(1)) * Fix(Num(args(0)) / Num(args(1)))
    Case "abs"
        NeedArgs nm, n, 1, 1
        r = Abs(Num(args(0)))
    Case "min", "max"
        NeedArgs nm, n, 1, -1
        acc = Num(args(0))
        For i = 1 To n - 1
            If nm = "min" And Num(args(i)) < acc Then acc = Num(args(i))
            If nm = "max" And Num(args(i)) > acc Then acc = Num(args(i))
        Next
        r = acc
    Case "sqrt"
        NeedArgs nm, n, 1, 1
        If Num(args(0)) < 0 Then LispError "sqrt: negative number"
        r = Sqr(Num(args(0)))

    '---------- 比较和逻辑 ----------
    Case "=", "<", ">", "<=", ">="
        NeedArgs nm, n, 2, -1
        r = True
        For i = 0 To n - 2
            If Not CompareNums(nm, Num(args(i)), Num(args(i + 1))) Then r = False
        Next
    Case "not"
        NeedArgs nm, n, 1, 1
        r = Not IsTrue(args(0))

    '---------- 序对和列表 ----------
    Case "cons"
        NeedArgs nm, n, 2, 2
        r = Cons(args(0), args(1))
    Case "car"
        NeedArgs nm, n, 1, 1
        r = Car(args(0))
    Case "cdr"
        NeedArgs nm, n, 1, 1
        r = Cdr(args(0))
    Case "set-car!"
        NeedArgs nm, n, 2, 2
        SetCar args(0), args(1)
        r = Empty
    Case "set-cdr!"
        NeedArgs nm, n, 2, 2
        SetCdr args(0), args(1)
        r = Empty
    Case "list"
        r = ArrayToList(args, 0)
    Case "length"
        NeedArgs nm, n, 1, 1
        r = CDbl(ListLength(AsList(nm, args(0))))
    Case "append"
        ' 前面的列表各复制一份，最后一个直接接在后面（和 Scheme 一样）
        r = NIL
        If n > 0 Then r = args(n - 1)
        For i = n - 2 To 0 Step -1
            r = AppendTwo(AsList(nm, args(i)), r)
        Next
    Case "apply"
        NeedArgs nm, n, 2, 2
        r = Apply(args(0), ListToArray(AsList(nm, args(1))))

    '---------- 类型判断和比较 ----------
    Case "null?"
        NeedArgs nm, n, 1, 1
        r = IsNil(args(0))
    Case "pair?"
        NeedArgs nm, n, 1, 1
        r = IsPair(args(0))
    Case "list?"
        NeedArgs nm, n, 1, 1
        r = (ListLength(args(0)) >= 0)
    Case "number?"
        NeedArgs nm, n, 1, 1
        r = (VarType(args(0)) = vbDouble)
    Case "symbol?"
        NeedArgs nm, n, 1, 1
        r = IsSymbol(args(0))
    Case "string?"
        NeedArgs nm, n, 1, 1
        r = IsString(args(0))
    Case "boolean?"
        NeedArgs nm, n, 1, 1
        r = (VarType(args(0)) = vbBoolean)
    Case "procedure?"
        NeedArgs nm, n, 1, 1
        r = (TagOf(args(0)) = "#lambda" Or TagOf(args(0)) = "#builtin")
    Case "eq?"
        NeedArgs nm, n, 2, 2
        r = IsSame(args(0), args(1))
    Case "equal?"
        NeedArgs nm, n, 2, 2
        r = IsEqual(args(0), args(1))

    '---------- 字符串 ----------
    Case "string-append"
        s = ""
        For i = 0 To n - 1
            s = s & StrArg(nm, args(i))
        Next
        r = MakeStr(s)
    Case "string-length"
        NeedArgs nm, n, 1, 1
        r = CDbl(Len(StrArg(nm, args(0))))
    Case "number->string"
        NeedArgs nm, n, 1, 1
        r = MakeStr(NumText(Num(args(0))))

    '---------- 输出和报错 ----------
    Case "display"
        NeedArgs nm, n, 1, 1
        gOut = gOut & ToText(args(0), False)
        r = Empty
    Case "newline"
        NeedArgs nm, n, 0, 0
        gOut = gOut & vbLf
        r = Empty
    Case "error"
        ' (error "消息" 其它值...)
        NeedArgs nm, n, 1, -1
        s = ToText(args(0), False)
        For i = 1 To n - 1
            s = s & " " & ToText(args(i), True)
        Next
        LispError s

    '---------- Excel ----------
    Case "cell"
        ' (cell "A1") 读当前工作表的单元格；空单元格读出 #f
        NeedArgs nm, n, 1, 1
        r = FromExcel(ActiveSheet.Range(StrArg(nm, args(0))).Value)
    Case "set-cell!"
        ' (set-cell! "B1" 值) 写单元格
        NeedArgs nm, n, 2, 2
        ActiveSheet.Range(StrArg(nm, args(0))).Value = ToExcel(args(1))
        r = Empty

    Case Else
        ' 在 BuiltinNames 里登记了名字，却忘了在这里写 Case
        LispError "builtin not implemented: " & nm
    End Select
    CallBuiltin = r
End Function

Private Function CompareNums(ByVal op As String, ByVal a As Double, ByVal b As Double) As Boolean
    Select Case op
    Case "=":  CompareNums = (a = b)
    Case "<":  CompareNums = (a < b)
    Case ">":  CompareNums = (a > b)
    Case "<=": CompareNums = (a <= b)
    Case ">=": CompareNums = (a >= b)
    End Select
End Function

' 把列表 a 复制一份，最后接上 b。AppendTwo((1 2), (3)) = (1 2 3)
Private Function AppendTwo(ByVal a As Variant, ByVal b As Variant) As Variant
    If IsNil(a) Then
        AppendTwo = b
    Else
        AppendTwo = Cons(Car(a), AppendTwo(Cdr(a), b))
    End If
End Function

' 单元格的值 -> Lisp 值
Private Function FromExcel(ByVal v As Variant) As Variant
    Select Case VarType(v)
    Case vbEmpty
        FromExcel = False
    Case vbBoolean
        FromExcel = v
    Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
        FromExcel = CDbl(v)
    Case vbError
        FromExcel = MakeStr("#ERROR")
    Case Else
        FromExcel = MakeStr(CStr(v))
    End Select
End Function

' Lisp 值 -> 单元格的值
Private Function ToExcel(ByVal v As Variant) As Variant
    If VarType(v) = vbDouble Or VarType(v) = vbBoolean Then
        ToExcel = v
    ElseIf IsEmpty(v) Then
        ToExcel = Empty
    Else
        ToExcel = ToText(v, False)
    End If
End Function

'==============================================================
' 第 7 部分  ToText —— 把值变回文字
' quoted = True 时字符串带引号显示（REPL 的写法），False 时不带（display 的写法）
'==============================================================
Public Function ToText(ByVal v As Variant, ByVal quoted As Boolean) As String
    Dim s As String
    If IsPair(v) Then
        ' 列表：顺着 cdr 走，把每个 car 写出来；最后如果不是 ()，写成“ . 尾巴”
        s = "("
        Do
            s = s & ToText(gCar(v), quoted)
            v = gCdr(v)
            If Not IsPair(v) Then Exit Do
            s = s & " "
        Loop
        If Not IsNil(v) Then s = s & " . " & ToText(v, quoted)
        ToText = s & ")"
        Exit Function
    End If
    Select Case TagOf(v)
    Case "#str"
        If quoted Then
            s = Replace(v(1), "\", "\\")
            s = Replace(s, """", "\""")
            s = Replace(s, vbLf, "\n")
            ToText = """" & s & """"
        Else
            ToText = v(1)
        End If
    Case "#lambda"
        ToText = "#<procedure>"
    Case "#builtin"
        ToText = "#<builtin " & v(1) & ">"
    Case Else
        If IsNil(v) Then
            ToText = "()"
        ElseIf VarType(v) = vbBoolean Then
            If v Then ToText = "#t" Else ToText = "#f"
        ElseIf VarType(v) = vbDouble Then
            ToText = NumText(v)
        ElseIf IsEmpty(v) Then
            ToText = ""
        Else
            ToText = CStr(v)        ' 符号
        End If
    End Select
End Function

' 数字 -> 文字。Str$ 总是用小数点“.”，不受系统区域设置影响
Private Function NumText(ByVal d As Double) As String
    Dim s As String
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    NumText = s
End Function

'==============================================================
' 第 8 部分  小工具函数
'==============================================================

' 出错：抛出一个 VBA 错误，最外层的 RunAndShow 会接住它
Private Sub LispError(ByVal msg As String)
    Err.Raise LISP_ERROR, "MiniLisp", msg
End Sub

' 带标签的值（字符串、函数）返回标签，例如 "#str"；其它值返回 ""
Private Function TagOf(ByVal v As Variant) As String
    If IsArray(v) Then
        If UBound(v) >= 0 Then
            If VarType(v(0)) = vbString Then
                If Left$(v(0), 1) = "#" Then TagOf = v(0)
            End If
        End If
    End If
End Function

Private Function IsSymbol(ByVal v As Variant) As Boolean
    IsSymbol = (VarType(v) = vbString)
End Function

Private Function IsString(ByVal v As Variant) As Boolean
    IsString = (TagOf(v) = "#str")
End Function

Private Function MakeStr(ByVal s As String) As Variant
    MakeStr = Array("#str", s)
End Function

' 只有 #f 是假，其它一切都是真
Private Function IsTrue(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTrue = v
    Else
        IsTrue = True
    End If
End Function

' 列表的长度；如果不是正常的列表（比如 (1 . 2) 或者 5），返回 -1
Private Function ListLength(ByVal lst As Variant) As Long
    Dim n As Long
    Do While IsPair(lst)
        n = n + 1
        lst = gCdr(lst)
    Loop
    If IsNil(lst) Then
        ListLength = n
    Else
        ListLength = -1
    End If
End Function

' 列表的第 i 个元素（从 0 开始数）。Nth((a b c), 1) = b
Private Function Nth(ByVal lst As Variant, ByVal i As Long) As Variant
    Nth = Car(Drop(lst, i))
End Function

' 去掉前 i 个元素后剩下的列表。Drop((a b c), 1) = (b c)
Private Function Drop(ByVal lst As Variant, ByVal i As Long) As Variant
    Dim k As Long
    For k = 1 To i
        lst = Cdr(lst)
    Next
    Drop = lst
End Function

' 如果列表的第一个元素是符号，返回它的名字；否则返回 ""
Private Function HeadName(ByVal x As Variant) As String
    If IsPair(x) Then
        If IsSymbol(gCar(x)) Then HeadName = gCar(x)
    End If
End Function

' VBA 数组从第 start 个元素起 -> Lisp 列表。从后往前 cons，最后得到的就是正序
Private Function ArrayToList(ByVal arr As Variant, ByVal start As Long) As Variant
    Dim i As Long, r As Variant
    r = NIL
    For i = UBound(arr) To start Step -1
        r = Cons(arr(i), r)
    Next
    ArrayToList = r
End Function

' Lisp 列表 -> VBA 数组
Private Function ListToArray(ByVal lst As Variant) As Variant
    Dim r() As Variant, i As Long, n As Long
    n = ListLength(lst)
    If n <= 0 Then
        ListToArray = Array()
        Exit Function
    End If
    ReDim r(0 To n - 1)
    For i = 0 To n - 1
        r(i) = Car(lst)
        lst = Cdr(lst)
    Next
    ListToArray = r
End Function

' eq?：序对比较“是不是同一个”（编号相同），其它值比较内容
Private Function IsSame(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) = vbLong Or VarType(b) = vbLong Then
        If VarType(a) = VarType(b) Then IsSame = (a = b)
    Else
        IsSame = IsEqual(a, b)
    End If
End Function

' equal?：两个值是否“长得一样”（列表逐个元素比较）
Private Function IsEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim i As Long
    If IsPair(a) And IsPair(b) Then
        If IsEqual(gCar(a), gCar(b)) Then IsEqual = IsEqual(gCdr(a), gCdr(b))
    ElseIf IsObject(a) Or IsObject(b) Then
        ' 函数里藏着的环境对象：是同一个才算相等
        If IsObject(a) And IsObject(b) Then IsEqual = (a Is b)
    ElseIf IsArray(a) Or IsArray(b) Then
        ' 字符串、函数：逐个元素比较
        If Not (IsArray(a) And IsArray(b)) Then Exit Function
        If UBound(a) <> UBound(b) Then Exit Function
        For i = 0 To UBound(a)
            If Not IsEqual(a(i), b(i)) Then Exit Function
        Next
        IsEqual = True
    ElseIf VarType(a) = VarType(b) Then
        IsEqual = (a = b)
    End If
End Function

'---------------- 参数检查 ----------------

' 参数个数检查：至少 lo 个，至多 hi 个（hi = -1 表示不限）
Private Sub NeedArgs(ByVal nm As String, ByVal n As Long, ByVal lo As Long, ByVal hi As Long)
    If n >= lo And (hi < 0 Or n <= hi) Then Exit Sub
    If lo = hi Then
        LispError nm & ": expects " & lo & " argument(s), got " & n
    ElseIf hi < 0 Then
        LispError nm & ": expects at least " & lo & " argument(s), got " & n
    Else
        LispError nm & ": expects " & lo & " to " & hi & " arguments, got " & n
    End If
End Sub

' 特殊形式的长度检查，不对就报语法错误
Private Sub CheckForm(ByVal x As Variant, ByVal lo As Long, ByVal hi As Long, ByVal usage As String)
    Dim n As Long
    n = ListLength(x)
    If n < lo Or (hi >= 0 And n > hi) Then LispError "bad syntax, expected " & usage
End Sub

Private Sub CheckSymbol(ByVal v As Variant, ByVal who As String)
    If Not IsSymbol(v) Then LispError who & ": expected a name, got " & ToText(v, True)
End Sub

Private Function Num(ByVal v As Variant) As Double
    If VarType(v) <> vbDouble Then LispError "expected a number, got " & ToText(v, True)
    Num = v
End Function

Private Function StrArg(ByVal who As String, ByVal v As Variant) As String
    If Not IsString(v) Then LispError who & ": expected a string, got " & ToText(v, True)
    StrArg = v(1)
End Function

Private Function AsList(ByVal who As String, ByVal v As Variant) As Variant
    If ListLength(v) < 0 Then LispError who & ": expected a list, got " & ToText(v, True)
    AsList = v
End Function

' 如果输出不是以换行结尾，先补一个换行
Private Sub FreshLine()
    If Len(gOut) > 0 Then
        If Right$(gOut, 1) <> vbLf Then gOut = gOut & vbLf
    End If
End Sub
