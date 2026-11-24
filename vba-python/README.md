# PyVBA：用 Excel VBA 实现的 Python 子集解释器

这是一个**纯 VBA** 编写的 Python 解释器：词法分析（含缩进 INDENT/DEDENT）、递归下降语法分析生成 AST、树遍历求值，
不依赖任何外部组件（不用 `Scripting.Dictionary`，因此 Windows 和 Mac 版 Excel 都能用）。

```python
# 写在 "Python" 工作表的 A 列里，点 "Run Python" 按钮运行
from dataclasses import dataclass

@dataclass
class Item:
    name: str
    price: float
    qty: int = 1

items = [Item("pen", 2.5, 4), Item("book", 12.0)]
total = sum(i.price * i.qty for i in items)
print(f"合计: {total:.2f}")                 # 合计: 22.00
set_cell("B1", total, sheet="Data")         # 写回 Excel
```

## 文件

| 文件 | 作用 |
| --- | --- |
| `src/PyNode.cls` | AST 节点类 |
| `src/PyObj.cls` | 运行时对象类（list / tuple / dict / set / 函数 / 类 / 实例 / 作用域……），内置有序哈希表 |
| `src/PyRuntime.bas` | 值模型与核心运算：数字、比较、`repr`、迭代、下标/切片、属性、运算符、异常 |
| `src/PyParser.bas` | 词法分析器 + 语法分析器（含 f-string） |
| `src/PyEval.bas` | 求值器：语句、表达式、函数调用、作用域、类、`try/except`，以及入口 `PyRunProgram` |
| `src/PyBuiltins.bas` | 内置函数、str/list/dict/set 方法、格式化、排序、标准库子集 |
| `src/PyExcel.bas` | Excel 集成：宏、工作表函数 `PY()`、`cell()`/`set_cell()` 等 |
| `tests/` | 测试：在 LibreOffice 的 VBA 兼容模式里运行解释器，并与真正的 CPython 输出逐行对比 |

## 安装

1. 打开（或新建）一个工作簿，按 `Alt+F11` 打开 VBA 编辑器。
2. 菜单 **文件 → 导入文件…**，依次导入 `src` 目录下的 7 个文件（两个 `.cls` 和五个 `.bas`）。
   请直接用仓库里的文件导入（行尾必须是 CRLF，仓库已通过 `.gitattributes` 保证；
   不要复制粘贴 `.cls` 文件的内容，否则类模块头部会丢失）。
3. 把工作簿另存为 **启用宏的工作簿（`.xlsm`）**。

> 如果想用宏一次导入，可在“信任中心”勾选“信任对 VBA 工程对象模型的访问”，然后运行：
>
> ```vba
> Sub ImportPyVBA()
>     Dim f As Variant, dir_ As String
>     dir_ = "C:\path\to\vba-python\src\"   ' 改成你的路径
>     For Each f In Array("PyNode.cls", "PyObj.cls", "PyRuntime.bas", "PyParser.bas", _
>                         "PyEval.bas", "PyBuiltins.bas", "PyExcel.bas")
>         ThisWorkbook.VBProject.VBComponents.Import dir_ & f
>     Next
> End Sub
> ```

## 使用

### 1. 在工作表里写程序

运行宏 **`SetupPythonSheets`**，会创建：

* `Python` 表：A 列写代码（一行一个单元格，缩进用空格），C 列显示输出，右侧有 **Run Python** 按钮；
  表里预置了一段示例程序；
* `Data` 表：一些示例数据。

点击按钮（或运行宏 **`RunPython`**）即可执行。未捕获的异常会像 Python 一样打印 traceback，行号就是 A 列的行号。
运行时按 `Ctrl+Break` 可以中断（抛出 `KeyboardInterrupt`）。

小提示：A 列已设为“文本”格式，避免 Excel 把 `1/2` 之类的内容改成日期；
以 `@`（装饰器）开头的行如果被 Excel 当成公式，请在前面加一个英文单引号 `'`。
另一个宏 **`RunSelectedPython`** 会运行当前选中单元格里的代码，并用消息框显示输出（适合写在一个单元格里、用 `Alt+Enter` 换行的小段代码）。

### 2. 工作表函数 `=PY(代码, 参数...)`

```
=PY("x * 2 + y", A1, B1)                 → 参数依次绑定为 x, y, z，全部参数在 args 列表里
=PY("sum(v*v for v in x)", A1:A10)       → 单行/单列区域变成列表，多行多列变成“列表的列表”
=PY("sorted(x, reverse=True)", A1:A5)    → 返回列表时溢出为一列（列表的列表则是二维区域）
=PY(D1)                                  → D1 里可以是多行程序；若定义了变量 result 就返回它，否则返回 print 的内容
```

出错时返回 `#ValueError: ...` 这样的文本。

### 3. 在 Python 里访问 Excel

| 函数 | 说明 |
| --- | --- |
| `cell(row, col)` / `cell("B3")` | 读一个单元格（整数会变成 `int`，空单元格是 `None`） |
| `set_cell(row, col, value)` / `set_cell("B3", value)` | 写一个单元格 |
| `get_range("A1:C10")` | 读区域：单行或单列返回列表，否则返回行的列表 |
| `set_range("A1", data)` | 写区域：`data` 可以是值、列表（写成一列）或列表的列表 |
| `clear_range("A1:C10")` | 清空内容 |
| `sheet_names()` | 所有工作表名 |
| `excel_eval("SUM(A1:A9)")` | 用 Excel 计算一个公式 |
| `msgbox(text)`、`input(prompt)` | 消息框 / 输入框 |

以上函数都接受 `sheet="名称"`（或序号）。不指定时默认操作 `Data` 表（不存在会自动创建，
这样程序不会误写到自己的代码表上）；在 `PY()` 中默认是公式所在的工作表。
也可以 `import excel` 后用 `excel.cell(...)` 的形式调用。

### 4. 在 VBA 里调用

```vba
Debug.Print PyRunProgram("print(sum(range(10)))")   ' 返回程序打印的全部文本（出错时附带 traceback）
```

## 支持的 Python 子集

**语句**：赋值（链式、元组解包、`a, *rest = ...`、下标/切片/属性赋值）、增量赋值 `+= -= *= /= //= %= **= &= |= ^= <<= >>=`、
带类型注解的赋值、`if/elif/else`、`while/else`、`for/else`、`break`、`continue`、`pass`、`def`（默认参数、`*args`、`**kwargs`、仅限关键字参数）、
`return`、`lambda`、`global`、`nonlocal`、`class`（单继承/多继承、`super()`）、装饰器、`try/except/else/finally`、`raise`（含 `raise ... from ...`）、
`assert`、`del`、`import` / `from ... import ...`、海象运算符 `:=`、`yield` / `yield from`、分号分隔多条语句。

**表达式**：算术、比较链 `a < b < c`、`and/or/not`、`in` / `not in`、`is` / `is not`、位运算、条件表达式、
列表/字典/集合推导式和生成器表达式、`*` / `**` 解包、切片（含步长）、f-string（格式说明、`!r`、嵌套 `{width}`、`{x=}`）、
`%` 格式化、`str.format`。

**数据类型**：`int`（超出 32 位自动用 VBA 的 Decimal，最多 28 位）、`float`、`bool`、`str`（Unicode，可用中文变量名）、
`None`、`list`、`tuple`、`dict`（保持插入顺序）、`set`、`range`，以及用户自定义类。

**类**：实例属性、类属性、`@property`（含 setter）、`@staticmethod`、`@classmethod`、`__init__ __str__ __repr__ __eq__ __lt__ __le__ __gt__ __ge__
__hash__ __len__ __bool__ __getitem__ __setitem__ __delitem__ __contains__ __iter__ __next__ __call__ __add__ __radd__ __sub__ __mul__ __truediv__ __neg__ __abs__ __format__ ...`

**内置函数**：`print len range int float str bool list dict tuple set frozenset type isinstance issubclass abs min max sum sorted reversed enumerate
zip map filter any all round pow divmod repr ascii chr ord hex bin oct input id hash callable getattr setattr hasattr delattr format iter next
staticmethod classmethod property super globals dir vars exit`，以及完整的异常层次（`Exception`、`ValueError`、`KeyError`……可继承）。

**字符串方法**：`upper lower strip split rsplit join replace find rfind index count startswith endswith isdigit isalpha isalnum isspace
title capitalize center ljust rjust zfill format splitlines partition removeprefix ...`；
**列表**：`append extend insert pop remove index count sort(key=, reverse=) reverse clear copy`；
**字典**：`keys values items get pop popitem setdefault update clear copy`；**集合**：`add remove discard union intersection difference ...`

**标准库子集**：`math`、`random`、`time`、`sys`（`setrecursionlimit` 等）、`string`、
`collections`（`Counter`、`defaultdict`、`OrderedDict`、`deque`）、`functools`（`reduce`、`lru_cache`、`cache`）、
`itertools`（`permutations`、`combinations`、`combinations_with_replacement`、`product`、`chain`、`accumulate`）、
`dataclasses`（`dataclass`、`field`、`asdict`）、`typing`（`List`、`Dict`、`Optional`……仅用于注解）、`__future__`。

## 与 CPython 的差异（限制）

* **生成器是“立即求值”的**：调用生成器函数时会先把所有 `yield` 的值算完，再返回一个迭代器。
  普通的有限生成器结果完全一样，但无限生成器（`while True: yield ...`）不受支持（超过 100 万个值会报错），
  `send()`、`yield` 表达式（`x = yield`）也不支持。`map`、`filter`、`zip`、`enumerate` 直接返回列表。
* 整数最大 28 位十进制（VBA Decimal），超出时抛 `OverflowError`；位运算只支持 32 位整数。
* 浮点数按 VBA 的精度打印最多 15 位有效数字（`0.1 + 0.2` 显示为 `0.3`，CPython 显示 `0.30000000000000004`）；
  没有真正的 `nan`，`inf` 用最大的 double 近似。
* 集合按插入顺序迭代（CPython 的集合顺序本来就无保证，测试时请用 `sorted()`）。
* 不支持：`with` 语句、`async`/`await`、`match`、文件读写、`bytes`、复数、`eval`/`exec`、自定义元类、
  `@property` 以外的描述符协议、`__slots__` 的限制作用。
* 作用域规则做了简化：函数里先读全局变量再给同名变量赋值不会报 `UnboundLocalError`。
* 默认递归深度上限 300（VBA 调用栈有限），可用 `sys.setrecursionlimit()` 调大到 5000；VBA 栈溢出也会转换成 `RecursionError`。
* 速度：树遍历解释器跑在 VBA 上，大约比 CPython 慢 100 倍以上，适合教学、自动化和中小规模的数据处理。

## 设计概要

* **词法分析**：逐字符扫描，用缩进栈生成 INDENT/DEDENT，括号内忽略换行；字符串支持各种前缀、转义和三引号；f-string 在语法分析阶段再拆分、递归解析其中的表达式。
* **语法分析**：标准的递归下降，优先级从 `lambda`、条件表达式、`or/and/not`、比较链、`| ^ & << >> + - * / // % **` 逐级下降，生成 `PyNode` 树；名字的哈希值在解析时预先算好。
* **值的表示**：`None`=Empty，`bool`=Boolean，`int`=Long（溢出时自动转 Decimal），`float`=Double，`str`=String，其余都是 `PyObj`。
* **字典与作用域**：`PyObj` 内置开放寻址的有序哈希表；Python 字典的键先转成规范化的键字符串（因此 `1 == 1.0 == True` 是同一个键）。
* **控制流**：`break/continue/return` 用语句返回的状态码传递；Python 异常用 `Err.Raise` 一个固定的错误号在 VBA 调用栈上传播，
  `try` 语句用 `On Error GoTo` 捕获，VBA 运行时错误（除零、溢出、类型不匹配……）也会被转换成对应的 Python 异常。
* **内存**：VBA 是引用计数，闭包和类会形成循环引用，每次运行结束时解释器会主动断开这些引用。

## 测试

`tests/run_tests.py` 用无界面的 LibreOffice（VBA 兼容模式）加载全部模块，运行 `tests/cases/` 下的每个 `.py` 文件，
并与 CPython 的输出逐行比较（对未捕获异常只比较最后一行 `ExcType: message`）：

```bash
/usr/bin/python3 tests/run_tests.py          # 需要 LibreOffice 和 python3-uno
python3 tests/vba_lint.py src/*.cls src/*.bas
```

`vba_lint.py` 是一个静态检查器，专门查 LibreOffice 不会报、但 Excel 的 VBA 编译器会报的错误
（未声明的变量、ByRef 参数类型不匹配、参数个数错误、带括号调用多参数 Sub 等）。
`tests/cases/x01_bigint.py`（超过 32 位的整数）依赖 VBA 的 Decimal 类型，LibreOffice 没有实现，因此只能在 Excel 里验证。
