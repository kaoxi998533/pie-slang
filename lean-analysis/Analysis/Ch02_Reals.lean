import Mathlib

/-!
# 第 2 章　实数：不等式与绝对值

数学分析的舞台是实数集 $\mathbb{R}$。从代数的角度看，$\mathbb{R}$ 是一个
**有序域**：可以做加减乘除，并且有一个与运算相容的大小关系 $\le$。
（实数区别于有理数的关键性质——**完备性**——留到下一章。）

分析中的证明，绝大部分篇幅都是**不等式的估计**。而估计"距离"最基本的工具
是**绝对值** $|x|$：$|x - a|$ 表示 $x$ 与 $a$ 在数轴上的距离。

在 Lean 中，绝对值写作 `|x|`（竖线直接用键盘输入）。

## 2.1 有序域中的不等式

上一章的 `linarith`、`nlinarith`、`positivity` 是处理不等式的主力。
对于非线性的不等式，经典技巧是"配方"：$(a - b)^2 \ge 0$。
`nlinarith` 接受一个提示列表，把这些事实加入推理。
-/

namespace Ch02

example (a b : ℝ) : a * b ≤ (a ^ 2 + b ^ 2) / 2 := by
  nlinarith [sq_nonneg (a - b)]

/-!
有时需要用到乘法的单调性。Mathlib 中的命名规则相当有规律：

* `mul_pos : 0 < a → 0 < b → 0 < a * b`
* `mul_le_mul_of_nonneg_left : b ≤ c → 0 ≤ a → a * b ≤ a * c`
* `div_le_div_of_nonneg_left`、`div_lt_iff₀`……

记不住名字不要紧：在 Lean 中输入 `exact?`，它会自动搜索能直接证明目标的定理。

## 2.2 绝对值

> **定义 2.1**　$|x| = \begin{cases} x, & x \ge 0 \\ -x, & x < 0 \end{cases}$

Mathlib 中关于绝对值的常用定理：

| 定理 | 内容 |
|------|------|
| `abs_nonneg x` | $0 \le \lvert x \rvert$ |
| `abs_of_nonneg` | $0 \le x \Rightarrow \lvert x \rvert = x$ |
| `abs_of_neg` | $x < 0 \Rightarrow \lvert x \rvert = -x$ |
| `le_abs_self x` | $x \le \lvert x \rvert$ |
| `neg_abs_le x` | $-\lvert x \rvert \le x$ |
| `abs_mul x y` | $\lvert xy \rvert = \lvert x \rvert \lvert y \rvert$ |
| `abs_sub_comm x y` | $\lvert x - y \rvert = \lvert y - x \rvert$ |
| `abs_lt` | $\lvert x \rvert < b \iff -b < x \land x < b$ |
| `abs_le` | $\lvert x \rvert \le b \iff -b \le x \land x \le b$ |

其中 `abs_lt` 与 `abs_le` 最为重要：它们把绝对值不等式转化为两个普通不等式，
之后就可以交给 `linarith`。

> **定理 2.2（三角不等式）**　$|a + b| \le |a| + |b|$。

**证明**：由 `abs_le`，只需证明 $-(|a| + |b|) \le a + b \le |a| + |b|$，
而这由 $-|a| \le a \le |a|$ 与 $-|b| \le b \le |b|$ 相加即得。
-/

theorem triangle (a b : ℝ) : |a + b| ≤ |a| + |b| := by
  rw [abs_le]
  constructor
  · linarith [neg_abs_le a, neg_abs_le b]
  · linarith [le_abs_self a, le_abs_self b]

/-!
`linarith [h₁, h₂]` 的写法表示"把 `h₁`、`h₂` 也加入推理"，效果和先用 `have` 引入一样。

Mathlib 中三角不等式的名字是 `abs_add_le a b`。之后的章节我们直接使用它。

## 2.3 "任意小"的论证

分析中最重要的思想之一：**要证明 $x \le y$，只需证明对任意 $\varepsilon > 0$ 有
$x < y + \varepsilon$**。直观上，如果 $x$ 比 $y$ 大，那么取 $\varepsilon = x - y$ 就矛盾了。
-/

theorem le_of_forall_lt_add {x y : ℝ} (h : ∀ ε > 0, x < y + ε) : x ≤ y := by
  by_contra hxy
  push Not at hxy
  have h1 := h (x - y) (by linarith)
  linarith

/-!
这个证明展示了分析中的典型模式：**反证 + 选取一个特殊的 $\varepsilon$**。

## 练习

### 练习 2.1　距离的三角不等式
$|a - c| \le |a - b| + |b - c|$：从 $a$ 到 $c$ 的距离不超过经过 $b$ 绕路的距离。

提示：利用 `triangle (a - b) (b - c)`。你可能需要 `calc` 和
`congr 1; ring` 来说明 $|a - c| = |(a - b) + (b - c)|$。
-/

theorem ex2_1 (a b c : ℝ) : |a - c| ≤ |a - b| + |b - c| := by
  sorry

/-!
### 练习 2.2　反向三角不等式
$|a| - |b| \le |a - b|$。

提示：$a = (a - b) + b$。
-/

theorem ex2_2 (a b : ℝ) : |a| - |b| ≤ |a - b| := by
  sorry

/-!
### 练习 2.3　邻域的刻画
$|x - a| < \varepsilon$ 当且仅当 $a - \varepsilon < x < a + \varepsilon$。

提示：`rw [abs_lt]`，然后 `constructor` 分别证明两个方向。
-/

theorem ex2_3 (x a ε : ℝ) : |x - a| < ε ↔ a - ε < x ∧ x < a + ε := by
  sorry

/-!
### 练习 2.4　任意小的数等于零
若对任意 $\varepsilon > 0$ 都有 $|x| < \varepsilon$，则 $x = 0$。
这个引理将用来证明**极限的唯一性**。

提示：反证，若 $x \ne 0$，则 $|x| > 0$（`abs_pos.mpr`），取 $\varepsilon = |x|$。
-/

theorem ex2_4 (x : ℝ) (h : ∀ ε > 0, |x| < ε) : x = 0 := by
  sorry

/-!
### 练习 2.5　绝对值的乘法估计
若 $|x| < 1$，则 $|x \cdot y| \le |y|$。

提示：`rw [abs_mul]`，然后用 `nlinarith` 并提供 `abs_nonneg y`。
-/

theorem ex2_5 (x y : ℝ) (hx : |x| < 1) : |x * y| ≤ |y| := by
  sorry

/-!
### 练习 2.6 ★　伯努利不等式
对 $x \ge -1$ 与自然数 $n$，$(1 + x)^n \ge 1 + n x$。

这是我们第一次使用**数学归纳法**。`induction n with` 产生两个情形：
`zero`（$n = 0$）与 `succ k ih`（由 $n = k$ 推 $n = k + 1$，`ih` 是归纳假设）。

提示：在 `succ` 情形，先 `push_cast` 把 `((k + 1 : ℕ) : ℝ)` 变为 `(k : ℝ) + 1`，
再 `rw [pow_succ]`。关键事实是 $(1+x)^k(1+x) \ge (1 + kx)(1 + x)$，
它来自 `mul_le_mul_of_nonneg_right ih (_ : 0 ≤ 1 + x)`。
最后还需要 $k x^2 \ge 0$。
-/

theorem ex2_6 (x : ℝ) (hx : -1 ≤ x) (n : ℕ) : 1 + n * x ≤ (1 + x) ^ n := by
  induction n with
  | zero => sorry
  | succ k ih => sorry

/-!
## 本章小结

* 实数是有序域，`linarith` / `nlinarith` / `positivity` 处理大部分不等式；
* 绝对值不等式用 `abs_lt`、`abs_le` 拆成两个普通不等式；
* 三角不等式 `abs_add_le` 是一切估计的基础；
* "对任意 $\varepsilon > 0$ 成立"的论证，常配合反证法与特殊的 $\varepsilon$。
-/

end Ch02
