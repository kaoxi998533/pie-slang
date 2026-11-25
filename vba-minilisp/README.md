# MiniLisp：写给初学者的 Excel VBA Lisp 解释器

用一千多行 VBA（含大量中文注释）实现一个真正的 Lisp。目标不是“功能全”，而是**每一行都看得懂、改得动**：

- 只有两个文件：`MiniLisp.bas`（解释器）和 `LispEnv.cls`（变量表）。
- 有真正的序对（cons cell）：点对 `(1 . 2)`、共享结构、`set-car!` / `set-cdr!` 都支持。所有序对存放在两个数组里，这两个数组就是 Lisp 的“内存”。
- 所有 Lisp 值都是普通的 VBA 值（数字、字符串、编号），不需要和对象的 `Set` 打交道。
- 源码按“词法分析 → 语法分析 → 求值 → 打印”从上到下排列，可以像读教程一样顺着读。
- 加一个内置函数只要改两处；加一个新语法（特殊形式）只要写一个小函数。

> 同一仓库里的 [`vba-lisp`](../vba-lisp/) 是功能完整的版本（尾调用优化、宏、续延、异常、哈希表、约 290 个内置过程），适合拿来用；这里的 MiniLisp 适合拿来学和改。

```lisp
(define (fact n)
  (if (= n 0) 1 (* n (fact (- n 1)))))
(display (fact 10))                 ; 3628800

(define (make-counter)
  (define count 0)
  (lambda () (set! count (+ count 1)) count))
(define c (make-counter))
(c) (c)
(display (c))                       ; 3

(display (cons 1 2))                ; (1 . 2)
(display (map (lambda (x) (* x x)) '(1 2 3 4)))   ; (1 4 9 16)
(set-cell! "B1" (fact 5))           ; 往单元格 B1 写入 120
```

---

## 目录

1. [快速开始](#1-快速开始)
2. [语言速览](#2-语言速览)
3. [解释器是怎么工作的](#3-解释器是怎么工作的)
4. [动手扩展](#4-动手扩展)
5. [练习题](#5-练习题)
6. [已知限制](#6-已知限制)
7. [文件和测试](#7-文件和测试)

---

## 1. 快速开始

### 导入

1. 打开 Excel，按 `Alt + F11` 打开 VBA 编辑器。
2. 菜单 **文件 → 导入文件**，依次导入 `src/LispEnv.cls` 和 `src/MiniLisp.bas`。
3. 保存时选择 **启用宏的工作簿 (*.xlsm)**。

> 源文件是 **GBK 编码 + CRLF 换行**，这是中文版 Windows 上 VBA 编辑器导入时需要的格式，导入后中文注释能正常显示。
> 如果你的系统不是中文 Windows，注释可能变成乱码，但代码照样能运行（代码里的字符串全是英文）。

### 四种用法

**① 单元格公式**：在任意单元格输入

```
=LISP("(+ 1 2)")                          → 3
=LISP("(define (sq x) (* x x)) (sq 12)")  → 144
=LISP(A1)                                  → 运行 A1 里写的程序
```

返回最后一个表达式的值；如果它没有值（比如最后是 `display`），返回打印出来的文字。每次调用都是独立的新环境。

**② 工作表“笔记本”**：在 A 列每个单元格写一段代码，按 `Alt + F8` 运行宏 **RunSheet**，结果会写到同一行的 B 列。前面定义的函数后面可以接着用：

| | A | B |
|---|---|---|
| 1 | `(define (sq x) (* x x))` | |
| 2 | `(sq 9)` | `81` |
| 3 | `(map sq '(1 2 3))` | `(1 4 9)` |
| 4 | `(oops)` | `Error: unbound variable: oops` |

**③ 交互式 REPL**：按 `Alt + F8` 运行宏 **LispRepl**，会弹出输入框，输入一个表达式就显示结果，点“取消”退出。

**④ 立即窗口**：在 VBA 编辑器里按 `Ctrl + G`，输入

```
?LispRun("(display (* 6 7))")
42
?LispRun("(+ 1 2) (list 1 2)", True)
3
(1 2)
```

第二个参数为 `True` 时，像 REPL 一样打印每个表达式的值。

---

## 2. 语言速览

### 数据类型

| 写法 | 含义 |
|---|---|
| `42` `-7` `3.14` | 数字（统一是 Double） |
| `#t` `#f` | 真、假。**只有 `#f` 是假**，`0` 和 `()` 都算真 |
| `"hello\n"` | 字符串，支持 `\n` `\t` `\"` `\\` |
| `x` `list?` `set-cell!` | 符号（变量名） |
| `(1 . 2)` | 序对（点对） |
| `(1 2 3)` `()` | 列表、空表。`(1 2 3)` 就是 `(1 . (2 . (3 . ())))` |
| `'x` `'(1 2)` | `(quote x)` 的简写：不求值，原样返回 |
| `; ...` | 注释，到行尾 |

### 特殊形式（语法）

| 形式 | 说明 |
|---|---|
| `(define x 10)` | 定义变量 |
| `(define (f a b) 函数体...)` | 定义函数；`(define (f a . rest) ...)` 把多出来的参数收集成列表 `rest` |
| `(lambda (a b) 函数体...)` | 匿名函数；`(lambda args ...)` 把所有参数收集成列表 |
| `(if 条件 是则 否则)` | 条件，否则部分可省略 |
| `(cond (条件 表达式...) ... (else 表达式...))` | 多分支条件 |
| `(let ((a 1) (b 2)) 函数体...)` | 局部变量 |
| `(set! x 新值)` | 修改已存在的变量 |
| `(begin 表达式...)` | 依次执行，返回最后一个的值 |
| `(and ...)` `(or ...)` | 短路求值 |
| `(quote x)` | 不求值 |

### 内置函数

| 类别 | 函数 |
|---|---|
| 算术 | `+` `-` `*` `/` `remainder` `abs` `min` `max` `sqrt` |
| 比较、逻辑 | `=` `<` `>` `<=` `>=`（可以连写，如 `(< 1 2 3)`）`not` |
| 序对和列表 | `cons` `car` `cdr` `set-car!` `set-cdr!` `list` `length` `append` `apply` |
| 类型判断 | `null?` `pair?` `list?` `number?` `symbol?` `string?` `boolean?` `procedure?` |
| 相等 | `eq?`（序对必须是同一个）`equal?`（内容一样就行） |
| 字符串 | `string-append` `string-length` `number->string` |
| 输出、报错 | `display` `newline` `(error "消息" 值...)` |
| Excel | `(cell "A1")` 读单元格（空单元格是 `#f`）、`(set-cell! "B1" 值)` 写单元格 |
| 用 Lisp 写的 | `cadr` `map` `filter` `reverse`（在 `Prelude` 里，见 4.2） |

---

## 3. 解释器是怎么工作的

任何解释器都在重复同一件事：

```
 源代码          记号                  表达式              值      文字
"(+ 1 2)" ──► ( + 1 2 ) ──► 列表 (+ 1 2) ──► 3 ──► "3"
          Tokenize      ReadExpr        Eval      ToText
          词法分析       语法分析         求值       打印
```

这就是 **Read-Eval-Print Loop（REPL）**，对应 `EvalAll` 函数里的那个循环：

```vb
Do While pos <= toks.Count
    expr = ReadExpr(toks, pos)              ' Read
    v = Eval(expr, env)                     ' Eval
    If showValues And Not IsEmpty(v) Then   ' Print
        ...ToText(v, True)...
    End If
Loop
```

下面按顺序讲每一步。建议一边看一边打开 `MiniLisp.bas` 对照。

### 3.1 数据怎么表示（第 0 部分）

Lisp 的一大特点是“代码就是数据”：程序 `(+ 1 2)` 本身就是一个列表。所以只要定好“列表怎么存”，代码也就有了存法。

| Lisp | VBA 里存的 |
|---|---|
| `42` | `42#`（Double） |
| `#t` | `True` |
| `x` | `"x"`（String —— 字符串在 MiniLisp 里表示**符号**） |
| `"hi"` | `Array("#str", "hi")`（真正的字符串加个标签，和符号区分） |
| 序对、列表 | 一个 `Long`：序对的编号（见下面） |
| `()` | `Long` 0（变量 `NIL`） |
| 用户函数 | `Array("#lambda", 参数表, 函数体, 定义时的环境)` |
| 内置函数 | `Array("#builtin", "+")` |

#### 序对和“内存”

Lisp 的列表全部由**序对**（pair，也叫 cons cell）搭成。一个序对只有两格：`car` 和 `cdr`。`(cons 1 2)` 造出一个序对，`car` 是 1，`cdr` 是 2，写出来是 `(1 . 2)`。

**列表只是一种约定**：`cdr` 指向下一个序对，最后一个序对的 `cdr` 是空表 `()`。

```
(1 2 3)  =  (1 . (2 . (3 . ())))

   ┌───┬───┐   ┌───┬───┐   ┌───┬───┐
   │ 1 │ ●─┼──►│ 2 │ ●─┼──►│ 3 │ ()│
   └───┴───┘   └───┴───┘   └───┴───┘
```

MiniLisp 把所有序对放在两个大数组里，这两个数组就是 Lisp 的“内存”：

```vb
Private gCar() As Variant       ' 第 i 号序对的 car
Private gCdr() As Variant       ' 第 i 号序对的 cdr
```

一个序对在 VBA 里就是它的**编号**（`Long` 类型），就像 C 语言里的指针。编号 0 不用，留给空表 `()`。读入 `(1 2 3)` 之后内存里是：

| 编号 | 1 | 2 | 3 |
|---|---|---|---|
| `gCar` | 1 | 2 | 3 |
| `gCdr` | 2 | 3 | 0 |

于是 `(1 2 3)` 这个值就是 `Long 1`：从 1 号序对出发，顺着 `gCdr` 一直走到 0。三个基本操作都只有几行：

```vb
Private Function Cons(ByVal a As Variant, ByVal d As Variant) As Long
    gPairCount = gPairCount + 1          ' 拿一个新编号（数组满了会自动扩容）
    gCar(gPairCount) = a
    gCdr(gPairCount) = d
    Cons = gPairCount
End Function

' Car(p) 就是 gCar(p)，Cdr(p) 就是 gCdr(p)（再加上“p 必须是序对”的检查）
```

这样表示有几个直接的好处，正好是 Lisp 列表的特点：

- **`cons` 和 `cdr` 都是 O(1)**：`(cons 0 lst)` 只造一个新序对，`gCdr` 填上 `lst` 的编号；原来的列表一点都不用复制。
- **结构可以共享**：`(define a (cons 1 tail))` 和 `(define b (cons 0 tail))` 的尾巴是同一串序对。用 `set-car!` 改了 `tail`，`a` 和 `b` 都会看到变化。
- **`eq?` 就是比较编号**：两个列表内容一样但编号不同，`equal?` 为真，`eq?` 为假。

因为数字都是 `Double`，`Long` 只用来表示序对，两者不会混淆（`IsPair` 只要检查 `VarType(v) = vbLong`）。

### 3.2 Tokenize：把文字切成记号（第 2 部分）

```
"(define x (+ 1 2))"  ──►  (  define  x  (  +  1  2  )  )
```

规则很少：`(` `)` `'` 各自是一个记号；空白分隔记号；`;` 到行尾是注释；`"..."` 是一个字符串记号。实现就是一个逐字符扫描的循环，结果放进一个 `Collection`。

### 3.3 ReadExpr：把记号变成列表（第 3 部分）

`ReadExpr` 读出**一个**完整的表达式：

- 遇到 `(`：交给 `ReadListRest` 读剩下的部分。
- 遇到 `'`：读出下一个表达式 `x`，返回 `(quote x)`。
- 其它：交给 `ParseAtom` 判断是数字、`#t`/`#f`、字符串还是符号。

`ReadListRest` 几乎就是列表定义本身：

```vb
If tok = ")" Then                ' 读到 )：列表结束，就是 ()
    ReadListRest = NIL
ElseIf tok = "." Then            ' 读到点号：后面那一个就是最后的 cdr
    ReadListRest = ReadExpr(toks, pos)   ' （再检查后面是 )）
Else                             ' 否则：先读一个元素，再读剩下的，cons 起来
    first = ReadExpr(toks, pos)
    ReadListRest = Cons(first, ReadListRest(toks, pos))
End If
```

元素本身又可能是列表——`ReadExpr` 和 `ReadListRest` 互相递归，嵌套多少层都能处理。

### 3.4 Eval：求值，解释器的心脏（第 4 部分）

`Eval(x, env)` 计算表达式 `x` 在环境 `env` 下的值，只有三种情况（简化后的样子）：

```vb
If IsSymbol(x) Then              ' ① 符号：查变量
    Eval = LookupVar(x, env)
ElseIf Not IsPair(x) Then        ' ② 数字、字符串……：值就是自己
    Eval = x
Else                             ' ③ 列表：看第一个元素
    Select Case HeadName(x)
    Case "if":     Eval = EvalIf(x, env)
    Case "define": Eval = EvalDefine(x, env)
    ...
    Case Else:     Eval = EvalCall(x, env)   ' 普通函数调用
    End Select
End If
```

处理表达式时用几个小工具取出各个部分：`Car(x)` 是第一个元素，`Nth(x, 1)` 是第二个，`Drop(x, 2)` 是去掉前两个之后剩下的列表。比如 `(if 条件 A B)` 里条件是 `Nth(x, 1)`，A 是 `Nth(x, 2)`。

**特殊形式**（`if`、`define`、`lambda`……）和普通函数调用的区别在于：函数调用会先把所有参数都求值，而特殊形式有自己的规则。比如 `(if 条件 A B)` 只能求值 A、B 中的一个——如果 `if` 是普通函数，`(if (= n 0) 1 (fact (- n 1)))` 会永远递归下去。

普通函数调用 `(f a b)` 的规则：先求值 `f`、`a`、`b`，再交给 `Apply`：

```vb
f = Eval(Car(x), env)
args = EvalArgs(Cdr(x), env)     ' 参数的值放进一个 VBA 数组
EvalCall = Apply(f, args)
```

参数的值放在 VBA 数组里而不是 Lisp 列表里，这样每次调用函数都不用造新的序对，省内存。

### 3.5 环境和闭包（LispEnv.cls）

环境就是一张“名字 → 值”的表，外加一个 `Parent` 指向外层环境。查找变量时先找本层，找不到就去 `Parent` 找，直到全局环境（`LispEnv.Find`）。

用户函数被调用时（`Apply`），会新建一个环境放参数，它的 `Parent` 是**函数定义时**的环境：

```lisp
(define (make-adder n) (lambda (x) (+ x n)))
(define add5 (make-adder 5))
(add5 10)   ; => 15
```

```
全局环境:  make-adder, add5, +, ...
   ▲
   │ Parent
调用 (make-adder 5) 时的环境:  n = 5      ◄── lambda 创建时记住了这个环境
   ▲
   │ Parent
调用 (add5 10) 时的环境:  x = 10
```

求值 `(+ x n)` 时：`x` 在本层找到 10；`n` 本层没有，去 Parent 找到 5。`make-adder` 早就返回了，但它的环境还被 lambda 记着——这就是**闭包**。`make-counter` 的例子里 `set!` 修改的也正是这个被记住的环境里的变量。

### 3.6 Apply 和内置函数（第 5、6 部分）

```vb
Select Case TagOf(f)
Case "#builtin"                      ' 内置函数：查表执行
    Apply = CallBuiltin(f(1), args)
Case "#lambda"                       ' 用户函数：新环境 + 求值函数体
    Set e = New LispEnv
    Set e.Parent = f(3)
    BindParams e, f(1), args
    Apply = EvalBody(f(2), e)
End Select
```

`BindParams` 沿着参数表一个个放参数。参数表是 `(a b . rest)` 这样的点对时，多出来的实参用 `ArrayToList` 打包成列表放进 `rest`；参数表是单个符号时收集全部实参。

`CallBuiltin` 是一个大 `Select Case`，每个 `Case` 是一个内置函数。全局环境里登记的内置函数只是 `Array("#builtin", "car")` 这样的“名片”，真正的代码在 `CallBuiltin` 里。

### 3.7 ToText：把值变回文字（第 7 部分）

序对的打印规则：顺着 `cdr` 走，把每个 `car` 写出来；如果最后停在的不是 `()`，就写成 ` . 尾巴`。所以 `(cons 1 (cons 2 3))` 打印成 `(1 2 . 3)`。数字用 `Str$` 转换（这样小数点永远是 `.`，不受系统区域设置影响）。参数 `quoted` 决定字符串要不要加引号：REPL 显示 `"hi"`，`display` 显示 `hi`。

### 3.8 走一遍完整流程

```lisp
(define (square x) (* x x))
(square 5)
```

1. **Tokenize**：`( define ( square x ) ( * x x ) ) ( square 5 )`
2. **ReadExpr** 读出第一个表达式：列表 `(define (square x) (* x x))`，在内存里是几个序对。
3. **Eval**：第一个元素是 `define` → `EvalDefine`：`Nth(x, 1)` 是 `(square x)`，它是序对，所以这是定义函数的写法 → 造一个函数 `Array("#lambda", (x), ((* x x)), 全局环境)`，登记为 `square`。
4. **ReadExpr** 读出第二个表达式：列表 `(square 5)`。
5. **Eval**：`square` 不是特殊形式 → `EvalCall`：求值 `square` 得到上面的函数，求值参数得到数组 `[5]` → `Apply`：新环境 `{x = 5}`，在里面求值函数体 `(* x x)` → 又是一次 `EvalCall`：`*` 是内置函数，参数是 `[5, 5]` → `CallBuiltin("*", ...)` → `25`。

---

## 4. 动手扩展

下面的每个例子都在本项目的测试环境里实际跑通过，可以直接照抄。

### 4.1 加一个内置函数

以 `(string-upcase "abc")` → `"ABC"` 为例，只要两步：

**第一步**，在 `BuiltinNames` 里登记名字：

```vb
Private Function BuiltinNames() As String
    BuiltinNames = "+ - * / remainder abs min max sqrt" _
        ...
        & " cell set-cell!" _
        & " string-upcase"
End Function
```

**第二步**，在 `CallBuiltin` 的 `Select Case` 里加一个 `Case`：

```vb
    Case "string-upcase"
        NeedArgs nm, n, 1, 1                        ' 恰好 1 个参数
        r = MakeStr(UCase$(StrArg(nm, args(0))))    ' 结果放进 r
```

可用的小工具：

| 工具 | 作用 |
|---|---|
| `args(i)`、`n` | 第 i 个参数（已求值，从 0 开始）、参数个数 |
| `NeedArgs nm, n, 最少, 最多` | 检查参数个数，最多为 `-1` 表示不限 |
| `Num(v)` / `StrArg(nm, v)` / `AsList(nm, v)` | 取出数字 / 字符串内容 / 列表，类型不对就报错 |
| `Cons(a, d)` / `Car(p)` / `Cdr(p)` / `NIL` | 造序对 / 取 car / 取 cdr / 空表 |
| `Nth(lst, i)` / `Drop(lst, i)` / `ListLength(lst)` | 第 i 个元素 / 去掉前 i 个 / 长度（不是正常列表时为 -1） |
| `ArrayToList(arr, 0)` / `ListToArray(lst)` | VBA 数组和 Lisp 列表互相转换 |
| `MakeStr(s)` | 把 VBA 字符串变成 Lisp 字符串 |
| `IsTrue(v)` | Lisp 意义上的真假 |
| `Apply(f, args)` | 调用一个 Lisp 函数，`args` 是 VBA 数组（写高阶函数时用） |
| `LispError "消息"` | 报错 |

返回“没有值”时写 `r = Empty`；返回真假直接写 `r = True` / `r = False`；返回数字时注意要是 `Double`（比如 `r = CDbl(Len(s))`）。

### 4.2 用 Lisp 扩展 Lisp

很多函数根本不用写 VBA。`Prelude` 函数里的代码会在启动时运行，`map`、`filter`、`reverse` 就是这样用 Lisp 写的。往里面加一行就多一个函数：

```vb
Private Function Prelude() As String
    Dim p As String
    ...
    p = p & "(define (sum lst) (if (null? lst) 0 (+ (car lst) (sum (cdr lst)))))" & vbLf
    p = p & "(define (greet name) (string-append ""Hello, "" name))" & vbLf
    Prelude = p
End Function
```

注意：在 VBA 字符串里，Lisp 代码中的双引号要写成两个 `""`。

**什么时候用 VBA、什么时候用 Lisp？** 需要访问 VBA/Excel 功能（字符串处理、单元格）或者追求速度时用 VBA；能用已有函数组合出来的，用 Lisp 写更短、更不容易出错。

### 4.3 加一个特殊形式（新语法）

当新功能**不能先把所有参数求值**时，就需要特殊形式。以 `(when 条件 表达式...)` 为例（条件为真时依次执行表达式）：

**第一步**，在 `Eval` 的 `Select Case` 里加一行：

```vb
    Case "or":     Eval = EvalOr(x, env)
    Case "when":   Eval = EvalWhen(x, env)
```

**第二步**，写一个处理函数。`x` 是整个表达式，`Nth(x, 1)` 是条件，`Drop(x, 2)` 是后面的表达式：

```vb
' (when 条件 表达式...)：条件为真时依次求值表达式
Private Function EvalWhen(ByVal x As Variant, ByVal env As LispEnv) As Variant
    CheckForm x, 2, -1, "(when test body...)"
    If IsTrue(Eval(Nth(x, 1), env)) Then
        EvalWhen = EvalBody(Drop(x, 2), env)
    Else
        EvalWhen = Empty
    End If
End Function
```

再来一个循环 `(while 条件 表达式...)`：

```vb
    Case "while":  Eval = EvalWhile(x, env)
```

```vb
' (while 条件 表达式...)：条件为真就一直重复执行
Private Function EvalWhile(ByVal x As Variant, ByVal env As LispEnv) As Variant
    CheckForm x, 2, -1, "(while test body...)"
    Do While IsTrue(Eval(Nth(x, 1), env))
        EvalBody Drop(x, 2), env
    Loop
    EvalWhile = Empty
End Function
```

```lisp
(define i 0)
(while (< i 5) (display i) (set! i (+ i 1)))   ; 01234
```

> 小提示：函数的每条执行路径都要给返回值赋值，没有值就写 `= Empty`。在 Excel 里不写也没事，但 LibreOffice 等其它环境可能会出错。

### 4.4 和 Excel 打交道

把一片单元格读成列表 `(range->list "A1:A3")` → `(1 "two" 3)`。构造列表的诀窍：**从最后一个元素往前，一个个 `cons` 到前面**：

```vb
' 在 CallBuiltin 开头的 Dim 后面加一行：
    Dim rng As Object

' 在 Select Case 里：
    Case "range->list"
        NeedArgs nm, n, 1, 1
        Set rng = ActiveSheet.Range(StrArg(nm, args(0)))
        r = NIL
        For i = rng.Cells.Count To 1 Step -1         ' 从最后一格往前
            r = Cons(FromExcel(rng.Cells(i).Value), r)
        Next
```

`FromExcel` / `ToExcel` 负责单元格值和 Lisp 值之间的转换。别忘了在 `BuiltinNames` 里登记 `range->list`。

---

## 5. 练习题

按难度排列。★ 只需要照着第 4 节的套路做；★★ 需要理解 `Eval` / 环境 / 序对；★★★ 要动解释器的结构。

**★ 入门**

1. 用 Lisp（写进 `Prelude`）实现 `even?`、`odd?`、`list-ref`、`last`、`assoc`（在“点对列表” `((a . 1) (b . 2))` 里按键查找）。
2. 用 VBA 加内置函数 `quotient`、`modulo`、`substring`、`string->number`、`symbol->string`。
3. 让 `display` 可以接受多个参数，依次打印。
4. 加特殊形式 `unless`（条件为假时执行）。

**★★ 进阶**

5. 加 `let*`：后面的绑定能用前面的变量。（提示：每个绑定都新建一层环境，或者在同一个新环境里依次 `Define`，并且在新环境里求值初值。）
6. 加 `(for i 起点 终点 表达式...)` 循环，参考 `while`。
7. 加 `list->range`：把列表写到一列单元格里。再试试把二维表格读成“列表的列表”。
8. 更好的错误信息：出错时显示是在调用哪个函数时出的错。（提示：在 `Apply` 里加 `On Error`，把函数名加到错误信息前面再抛出去。）
9. 加 `eval` 函数：`(eval '(+ 1 2))` → `3`。在 `CallBuiltin` 里调用 `Eval` 就行——想想应该在哪个环境里求值。
10. 用 `set-cdr!` 造一个循环列表，再 `display` 它，会发生什么？（程序停不下来时按 `Esc` 或 `Ctrl + Break` 中断。）修改 `ToText`，让它遇到循环时停下来（提示：记录走过的序对编号）。

**★★★ 挑战**

11. **尾调用优化（TCO）**。现在 `(count-down 100000)` 这种很深的递归会让 VBA 报“堆栈空间不足”。思路：把 `Eval` 改成一个 `Do ... Loop` 循环，遇到 `if` 时不再递归调用 `Eval(Nth(x, 2), env)`，而是把 `x` 换成 `Nth(x, 2)` 然后继续循环；调用用户函数时，新建环境后把 `env` 换成新环境、`x` 换成函数体的最后一个表达式，继续循环（前面的表达式照常求值）。
12. **宏**。加 `(define-macro (名字 参数...) 函数体...)`：存成 `Array("#macro", ...)`；`Eval` 遇到调用宏时，把**未求值的**参数交给它，得到一段新代码，再 `Eval` 这段代码。然后用宏实现 `unless`、`when`。
13. **垃圾回收**。现在造过的序对只有在重新开始（`LispReset`）时才会全部回收。在每个顶层表达式求值完之后做一次“标记-整理”：从全局环境里所有变量的值出发，标记所有还能走到的序对，把它们搬到数组前部并更新编号。（需要给 `LispEnv` 加一个能遍历所有值的方法。）
14. **更快的变量查找**。把 `LispEnv` 里的数组换成 `Scripting.Dictionary`（仅 Windows），比较 `(fib 20)` 前后的速度。

---

## 6. 已知限制

这些都是为了简单而故意做的取舍，也正好是练习的方向：

- **没有垃圾回收**：程序运行期间造出的序对一直占着内存，直到下一次 `LispReset`（每次 `LispRun`、`RunSheet`、`LispRepl` 开始时自动调用）。`=LISP()` 每次调用结束都会把自己造的序对还回去。
- **没有尾调用优化**：递归太深（几千层）会耗尽 VBA 的堆栈。
- **循环结构**：用 `set-cdr!` 造出的循环列表，`display` 会陷入死循环（按 `Esc` 或 `Ctrl + Break` 中断），`equal?` 会一直递归直到堆栈耗尽。
- **数字都是 Double**：超过 15 位的整数会丢失精度；没有分数，`(/ 1 3)` 是 `0.333333333333333`。
- **字符串不可修改**：没有 `string-set!`。
- **变量查找是线性的**：每次查找要扫描整张表。
- **单元格公式 `=LISP(...)` 不能修改单元格**：这是 Excel 对工作表函数的限制，`set-cell!` 请在宏（`RunSheet`、`LispRepl`）里使用。`cell` / `set-cell!` 操作的是当前活动工作表。

---

## 7. 文件和测试

```
vba-minilisp/
├── src/
│   ├── MiniLisp.bas      解释器（内存、读入、求值、内置函数、Excel 入口）
│   └── LispEnv.cls       环境（变量表）
├── tests/
│   ├── run_tests.py      在 LibreOffice 里运行全部测试
│   ├── vba_lint.py       静态检查（VBA 编译错误）
│   └── cases/*.scm/.out  测试程序和期望输出
└── README.md
```

测试在 Linux 上用无界面的 LibreOffice（VBA 兼容模式）运行解释器，把每个 `cases/*.scm` 的输出和对应的 `.out` 比较，最后还会测试 `RunSheet`、`=LISP()`、`cell`、`set-cell!`：

```
/usr/bin/python3 tests/run_tests.py        # 需要 LibreOffice 和 python3-uno
python3 tests/vba_lint.py src/*.bas src/*.cls
```

`.scm` 文件第一行是 `;; echo` 时按 REPL 方式运行（打印每个表达式的值）；是 `;; each` 时每一行单独运行，适合测试错误信息。`t07_pairs.scm` 专门测试序对：点对、共享结构、`set-car!`/`set-cdr!`、`eq?` 和剩余参数。

改了解释器之后跑一遍测试，就知道有没有把原来的功能改坏——给自己的扩展也写几个测试吧。
