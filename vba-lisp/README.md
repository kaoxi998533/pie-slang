# VBA Lisp：用 Excel VBA 实现的 Scheme 子集（完整版）

这是一个功能比较完整的 Scheme 解释器，用纯 VBA 写成，可以在 Excel 里直接运行。

> 想学习“解释器是怎么写出来的”，请先看同仓库的 [`vba-minilisp`](../vba-minilisp/)：那是给初学者的精简版，两个文件、中文注释、附教程和练习。
> 这里的完整版代码量大约是它的 7 倍，目标是**好用**：尾调用优化、宏、续延、异常、记录类型、哈希表等都有。

```scheme
(define-record-type point (make-point x y) point? (x point-x) (y point-y))

(define-syntax swap!
  (syntax-rules ()
    ((_ a b) (let ((tmp a)) (set! a b) (set! b tmp)))))

(define (sum-to n)                 ; 尾递归：再深也不会爆栈
  (let loop ((i 0) (acc 0))
    (if (> i n) acc (loop (+ i 1) (+ acc i)))))
(display (sum-to 10000))           ; 50005000

(define rows (get-range "A1:B5"))  ; 读一片单元格
(set-cell! "D1" (apply + (map cadr (cdr rows))))
```

---

## 1. 安装

1. 打开 Excel，按 `Alt + F11` 打开 VBA 编辑器。
2. **文件 → 导入文件**，导入 `src/` 下全部 6 个文件：`LispObj.cls`、`LispCore.bas`、`LispReader.bas`、`LispEval.bas`、`LispBuiltins.bas`、`LispExcel.bas`。
3. 回到 Excel，按 `Alt + F8` 运行宏 **SetupLispSheet**：会建好一个 `Lisp` 工作表（A 列写代码、C 列显示输出、带“运行”按钮）和一个 `Data` 工作表，并放好示例程序。
4. 保存为 **启用宏的工作簿 (*.xlsm)**。

源文件是 ASCII 编码、CRLF 换行，任何语言版本的 Excel 都能直接导入。

## 2. 使用方法

| 方式 | 说明 |
|---|---|
| 宏 `RunLisp` | 运行 `Lisp` 表 A 列的整个程序，输出写到 C 列；顶层表达式的值显示为 `=> 值` |
| 宏 `RunSelectedLisp` | 运行当前选中单元格里的代码，用消息框显示结果 |
| 工作表函数 `=LISP(code, [参数...])` | 返回最后一个表达式的值；额外参数绑定到 `x`、`y`、`z` 和列表 `args`，例如 `=LISP("(* x y)", A1, B1)` |
| VBA 里调用 | `LispRunProgram(src, echo)` 返回程序输出（`echo = True` 时附带每个顶层表达式的值） |

## 3. 支持的语言特性

**数据类型**：整数（`Long`，超过 32 位时自动转为 `Decimal`，最多 28 位）、浮点数、布尔、字符（`#\a`）、字符串、符号、**真正的序对和点对**（`(1 . 2)`）、列表、向量（`#(1 2 3)`）、哈希表、记录、promise/流、多值、续延。

**特殊形式**：
`quote` `quasiquote` `unquote` `unquote-splicing` `if` `define` `set!` `lambda` `named-lambda` `begin`
`let`（含命名 let）`let*` `letrec` `letrec*` `cond`（含 `=>`）`case` `and` `or` `when` `unless` `do`
`delay` `delay-force` `make-promise-lazy` `cons-stream`
`define-syntax` `let-syntax` `letrec-syntax` `syntax-rules`（含省略号 `...`）`define-macro` `defmacro`
`guard` `assert` `define-record-type` `case-lambda` `let-values` `let*-values` `define-values`

**内置过程**（约 290 个，常用的有）：

| 类别 | 例子 |
|---|---|
| 数值 | `+ - * /` `quotient remainder modulo` `gcd lcm expt sqrt exact-integer-sqrt` `floor ceiling round truncate` `sin cos exp log` `number->string string->number` `random` |
| 序对和列表 | `cons car cdr set-car! set-cdr!` `list length append reverse list-tail list-ref` `memq member assq assoc` `map for-each filter reduce fold-left fold-right` `sort iota delete any every` |
| 字符和字符串 | `string-append substring string-length string-ref` `string-upcase string-split string-join string-trim string-contains` `char->integer char-upcase char-alphabetic?` |
| 向量和哈希表 | `make-vector vector-ref vector-set!` `make-hash-table hash-table-set! hash-table-ref hash-table-keys hash-table-walk` |
| 控制 | `apply` `call/cc` `dynamic-wind` `values call-with-values` `eval` `force` |
| 异常 | `error raise with-exception-handler error-object-message error-object-irritants` |
| 输出 | `display write newline write-string format` |

**尾调用优化**：`if`、`cond`、`case`、`and`、`or`、`when`、`unless`、`let` 系列、`begin` 的最后一个表达式以及过程体中的尾调用都不会增加 VBA 的调用栈，所以用尾递归写的循环可以跑任意多次。非尾递归的深度上限是 2000 层（超过时报 `maximum recursion depth exceeded`，而不是让 Excel 崩溃）。

## 4. 和 Excel 交互

可选的最后一个参数是工作表名；默认是 `Data` 表（不存在时自动创建），在 `=LISP()` 里默认是公式所在的工作表。空单元格读出来是 `()`。

| 过程 | 说明 |
|---|---|
| `(cell 行 列 [表])` / `(cell "B3" [表])` | 读一个单元格 |
| `(set-cell! 行 列 值 [表])` / `(set-cell! "B3" 值 [表])` | 写一个单元格 |
| `(get-range "A1:C10" [表])` | 读成“行的列表”；只有一行或一列时是平的列表 |
| `(set-range! "A1" 数据 [表])` | 数据可以是单个值、列表（写成一列）或列表的列表（写成表格） |
| `(clear-range! "A1:C10" [表])` | 清空区域 |
| `(sheet-names)` | 所有工作表名 |
| `(excel-eval "SUM(A1:A9)")` | 让 Excel 计算一个公式 |
| `(msgbox 文字)` | 弹出消息框 |

## 5. 代码结构

| 文件 | 内容 |
|---|---|
| `LispObj.cls` | 所有堆对象（序对、符号、字符串、过程、环境、向量、哈希表……）共用的类 |
| `LispCore.bas` | 对象创建、符号表、数值运算（`Long` → `Decimal` → `Double` 自动提升）、输出缓冲 |
| `LispReader.bas` | 词法和语法分析，记录行号用于报错 |
| `LispEval.bas` | 求值器（循环实现的尾调用优化）、特殊形式、`syntax-rules` 宏展开、续延和异常 |
| `LispBuiltins.bas` | 内置过程 |
| `LispExcel.bas` | 宏、`LISP()` 工作表函数和单元格操作 |

## 6. 已知限制

- `call/cc` 只支持**逃逸式**续延（向外跳出，例如提前返回、跳出循环），不能重新进入已经返回的续延。
- 没有分数（`(/ 1 3)` 是浮点数）、没有复数；整数最多 28 位（VBA `Decimal` 的上限）。
- `syntax-rules` 不是完全卫生（hygienic）的：宏里引入的临时变量会被重命名，但宏里引用的自由变量按使用处解析。
- 只有逃逸续延，所以 `dynamic-wind` 只处理向外跳出的情况。
- 没有文件读写和端口（port）。
- 运行速度比较慢：在 LibreOffice 里一万次循环要二十多秒。Excel 的 VBA 通常快不少（没有实测），但整体上仍适合脚本和教学用途，不适合大量计算。

## 7. 测试

`tests/run_tests.py` 用无界面的 LibreOffice（VBA 兼容模式）加载全部模块，运行 `tests/cases/*.scm` 并和 `.out` 文件比较：

```
/usr/bin/python3 tests/run_tests.py            # 需要 LibreOffice 和 python3-uno
python3 tests/vba_lint.py src/*.bas src/*.cls  # 静态检查 Excel 才会报的编译错误
```

`x01_bigint.scm` 用到 VBA 的 `Decimal` 类型，LibreOffice 不支持，只能在 Excel 里验证，所以默认不运行。其余 15 个测试在 LibreOffice 中全部通过。
