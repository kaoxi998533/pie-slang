import Mathlib

/-!
# 第 8 章　闭区间上连续函数的性质

上一章研究的是函数在**一点**附近的性质。本章研究连续函数在**整个闭区间**上的
整体性质。这些定理在直观上"显而易见"，但它们的证明都离不开实数的完备性：
在 $\mathbb{Q}$ 上，函数 $x^2 - 2$ 在 $[1, 2]$ 上连续，两端异号，却没有零点！

本章我们学习**如何使用 Mathlib 中的定理**。这是形式化数学中非常重要的一项技能：
读懂定理的陈述，检查它需要的条件，再把它应用到具体问题上。

## 8.1 记号

* `Set.Icc a b` 是闭区间 $[a, b]$，`Set.Ioo a b` 是开区间 $(a, b)$；
* `ContinuousOn f s` 表示 $f$ 在集合 $s$ 上连续；
* `f '' s` 表示像集 $f(s) = \{f(x) \mid x \in s\}$；
* `x ∈ Set.Icc a b` 展开后就是 `a ≤ x ∧ x ≤ b`（定理 `Set.mem_Icc`）。

对于多项式这类由基本运算组合出来的函数，连续性可以用 `fun_prop` 自动证明。
-/

namespace Ch08

open Set

example : ContinuousOn (fun x : ℝ ↦ x ^ 3 + x - 1) (Icc 0 1) := by
  fun_prop

/-!
## 8.2 介值定理

> **定理 8.1（介值定理）**　设 $f$ 在 $[a, b]$ 上连续，则 $f$ 取到 $f(a)$ 与 $f(b)$
> 之间的每一个值。

Mathlib 中的陈述是：

```text
intermediate_value_Icc  : a ≤ b → ContinuousOn f (Icc a b) → Icc (f a) (f b) ⊆ f '' Icc a b
intermediate_value_Icc' : a ≤ b → ContinuousOn f (Icc a b) → Icc (f b) (f a) ⊆ f '' Icc a b
```

第一个适用于 $f(a) \le f(b)$ 的情形，第二个适用于 $f(a) \ge f(b)$ 的情形。

> **例 8.2**　存在 $x \in [1, 2]$ 使得 $x^2 = 2$。也就是说，$\sqrt{2}$ 存在！

**证明**：$f(x) = x^2$ 在 $[1,2]$ 上连续，$f(1) = 1 \le 2 \le 4 = f(2)$。
-/

example : ∃ x ∈ Icc (1 : ℝ) 2, x ^ 2 = 2 := by
  have hcont : ContinuousOn (fun x : ℝ ↦ x ^ 2) (Icc 1 2) := by fun_prop
  have hivt := intermediate_value_Icc (by norm_num : (1 : ℝ) ≤ 2) hcont
  -- hivt : Icc (1 ^ 2) (2 ^ 2) ⊆ (fun x ↦ x ^ 2) '' Icc 1 2
  have h2 : (2 : ℝ) ∈ Icc ((1 : ℝ) ^ 2) ((2 : ℝ) ^ 2) := by
    rw [mem_Icc]
    norm_num
  obtain ⟨x, hx, hx2⟩ := hivt h2
  exact ⟨x, hx, hx2⟩

/-!
`hivt h2` 的意思是：`hivt` 是一个包含关系 $A \subseteq B$，作用在 `h2 : 2 ∈ A` 上
就得到 `2 ∈ B`，即"存在 $x \in [1,2]$ 使 $x^2 = 2$"。

## 8.3 最值定理

> **定理 8.3（最值定理）**　闭区间上的连续函数一定能取到最大值和最小值。

在 Mathlib 中，这个定理是对**紧集**陈述的：`isCompact_Icc` 说闭区间是紧的，
`IsCompact.exists_isMaxOn` 说紧集上的连续函数有最大值点。
`IsMaxOn f s c` 的意思是 $c$ 是 $f$ 在 $s$ 上的最大值点，
`isMaxOn_iff` 把它展开为 `∀ x ∈ s, f x ≤ f c`。
-/

theorem exists_max {f : ℝ → ℝ} {a b : ℝ} (hab : a ≤ b) (hf : ContinuousOn f (Icc a b)) :
    ∃ c ∈ Icc a b, ∀ x ∈ Icc a b, f x ≤ f c := by
  obtain ⟨c, hc, hmax⟩ := isCompact_Icc.exists_isMaxOn (nonempty_Icc.mpr hab) hf
  exact ⟨c, hc, isMaxOn_iff.mp hmax⟩

/-!
## 练习

### 练习 8.1　方程的根
证明方程 $x^3 + x - 1 = 0$ 在 $[0, 1]$ 中有根。

提示：模仿例 8.2。
-/

theorem ex8_1 : ∃ x ∈ Icc (0 : ℝ) 1, x ^ 3 + x - 1 = 0 := by
  sorry

/-!
### 练习 8.2　最小值
仿照 `exists_max` 证明最小值的存在性。

提示：`IsCompact.exists_isMinOn` 与 `isMinOn_iff`。
-/

theorem exists_min {f : ℝ → ℝ} {a b : ℝ} (hab : a ≤ b) (hf : ContinuousOn f (Icc a b)) :
    ∃ c ∈ Icc a b, ∀ x ∈ Icc a b, f c ≤ f x := by
  sorry

/-!
### 练习 8.3　闭区间上的连续函数有界
提示：用 `exists_max` 和 `exists_min` 得到最大值 $f(c_1)$ 和最小值 $f(c_2)$，
取 $M = |f(c_1)| + |f(c_2)|$。用 `abs_le` 拆开绝对值，再配合 `le_abs_self`、`neg_abs_le`。
-/

theorem ex8_3 {f : ℝ → ℝ} {a b : ℝ} (hab : a ≤ b) (hf : ContinuousOn f (Icc a b)) :
    ∃ M, ∀ x ∈ Icc a b, |f x| ≤ M := by
  sorry

/-!
### 练习 8.4 ★　不动点定理
若 $f : [0, 1] \to [0, 1]$ 连续，则存在 $x \in [0,1]$ 使 $f(x) = x$。

提示：考虑 $g(x) = f(x) - x$，则 $g(0) \ge 0 \ge g(1)$。
对 $g$ 使用 `intermediate_value_Icc'`，目标值为 $0$。
`ContinuousOn.sub hf continuousOn_id` 给出 $g$ 的连续性（或者试试 `fun_prop`）。
-/

theorem ex8_4 {f : ℝ → ℝ} (hf : ContinuousOn f (Icc 0 1))
    (hmaps : ∀ x ∈ Icc (0 : ℝ) 1, f x ∈ Icc (0 : ℝ) 1) :
    ∃ x ∈ Icc (0 : ℝ) 1, f x = x := by
  sorry

/-!
### 练习 8.5 ★　开区间上可能取不到最大值
最值定理中"闭区间"的条件不能去掉：$f(x) = x$ 在 $(0, 1)$ 上没有最大值。

提示：若 $c \in (0, 1)$ 是最大值点，考虑 $\frac{c + 1}{2}$。`Set.mem_Ioo` 展开开区间的成员关系。
-/

theorem ex8_5 : ¬ ∃ c ∈ Ioo (0 : ℝ) 1, ∀ x ∈ Ioo (0 : ℝ) 1, x ≤ c := by
  sorry

/-!
## 本章小结

* 介值定理：`intermediate_value_Icc`（及其反向版本 `intermediate_value_Icc'`）；
* 最值定理：`IsCompact.exists_isMaxOn` / `exists_isMinOn` 配合 `isCompact_Icc`；
* 连续性的自动证明：`fun_prop`；
* 使用库中定理时，关键是逐条核对它的**前提条件**。

**延伸阅读**：闭区间上的连续函数还是**一致连续**的（Cantor 定理），
在 Mathlib 中为 `IsCompact.uniformContinuousOn_of_continuous`。
-/

end Ch08
