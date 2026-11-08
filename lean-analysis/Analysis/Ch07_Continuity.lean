import Mathlib

/-!
# 第 7 章　函数的连续性

直观上，连续函数的图像是"一笔画成、没有断开"的曲线。
柯西和魏尔斯特拉斯把这个直观变成了精确的 $\varepsilon$-$\delta$ 语言。

## 7.1 $\varepsilon$-$\delta$ 定义

> **定义 7.1**　函数 $f : \mathbb{R} \to \mathbb{R}$ 在点 $a$ **连续**，如果
> $$\forall \varepsilon > 0,\ \exists \delta > 0,\ \forall x,\ |x - a| < \delta \Rightarrow |f(x) - f(a)| < \varepsilon.$$

与数列极限的游戏类似：对手给出输出端的误差 $\varepsilon$，
你要找到输入端的误差 $\delta$，使得 $x$ 离 $a$ 不超过 $\delta$ 时，
$f(x)$ 离 $f(a)$ 不超过 $\varepsilon$。

<!-- widget:eps-delta -->

（网页版在这里有一个交互演示：选择函数并拖动 $\varepsilon$。图中横向色带是
$f(a) \pm \varepsilon$，纵向色带是 $a \pm \delta$。对符号函数，当 $\varepsilon \le 2$ 时
无论 $\delta$ 多小都会有一段图像跑出横向色带——这就是"不连续"。）
-/

namespace Ch07

/-- 函数 `f` 在点 `a` 处连续。 -/
def ContAt (f : ℝ → ℝ) (a : ℝ) : Prop :=
  ∀ ε > 0, ∃ δ > 0, ∀ x, |x - a| < δ → |f x - f a| < ε

/-!
## 7.2 例子：线性函数

> **例 7.2**　$f(x) = 3x + 1$ 在每一点连续。

**草稿**：$|f(x) - f(a)| = 3|x - a|$。要使它 $< \varepsilon$，只需 $|x - a| < \frac{\varepsilon}{3}$。
所以取 $\delta = \frac{\varepsilon}{3}$。

`use ε / 3, by positivity` 同时给出 $\delta$ 和 $\delta > 0$ 的证明。
-/

example (a : ℝ) : ContAt (fun x ↦ 3 * x + 1) a := by
  intro ε hε
  use ε / 3, by positivity
  intro x hx
  calc |3 * x + 1 - (3 * a + 1)| = 3 * |x - a| := by
        rw [show 3 * x + 1 - (3 * a + 1) = 3 * (x - a) by ring, abs_mul]
        norm_num
    _ < 3 * (ε / 3) := by linarith
    _ = ε := by ring

/-!
## 7.3 例子：不连续的函数

要证明 $f$ 在 $a$ **不**连续，需要写出定义的否定：
$$\exists \varepsilon > 0,\ \forall \delta > 0,\ \exists x,\ |x - a| < \delta \text{ 且 } |f(x) - f(a)| \ge \varepsilon.$$

> **例 7.3**　符号函数 $f(x) = \begin{cases} -1, & x < 0 \\ 1, & x \ge 0 \end{cases}$ 在 $0$ 处不连续。

**证明**：取 $\varepsilon = 1$。无论 $\delta$ 多小，$x = -\frac{\delta}{2}$ 都满足 $|x| < \delta$，
但 $|f(x) - f(0)| = |-1 - 1| = 2 \ge 1$。
-/

example : ¬ ContAt (fun x ↦ if x < 0 then (-1 : ℝ) else 1) 0 := by
  intro h
  obtain ⟨δ, hδ, hδ'⟩ := h 1 one_pos
  have hx : -δ / 2 < 0 := by linarith
  have h1 := hδ' (-δ / 2) (by rw [sub_zero, abs_lt]; constructor <;> linarith)
  simp [hx] at h1
  norm_num at h1

/-!
## 7.4 连续性与数列极限

连续性也可以用数列来刻画（**海涅定理**）：$f$ 在 $a$ 连续当且仅当
对每个收敛到 $a$ 的数列 $x_n$，都有 $f(x_n) \to f(a)$。练习 7.6 证明其中一个方向。
为此我们需要数列极限的定义。
-/

/-- 数列 `a` 收敛到 `L`。 -/
def TendsTo (a : ℕ → ℝ) (L : ℝ) : Prop :=
  ∀ ε > 0, ∃ N : ℕ, ∀ n ≥ N, |a n - L| < ε

/-!
## 练习

### 练习 7.1　常值函数连续
提示：任何 $\delta > 0$ 都可以，比如 `use 1, one_pos`。
-/

theorem contAt_const (c a : ℝ) : ContAt (fun _ ↦ c) a := by
  sorry

/-!
### 练习 7.2　恒等函数连续
提示：取 $\delta = \varepsilon$。
-/

theorem contAt_id (a : ℝ) : ContAt (fun x ↦ x) a := by
  sorry

/-!
### 练习 7.3　连续函数之和连续
提示：$\frac{\varepsilon}{2}$ 技巧，取 $\delta = \min(\delta_1, \delta_2)$。
`lt_min : a < b → a < c → a < min b c`，
`min_le_left a b : min a b ≤ a`，`min_le_right a b : min a b ≤ b`。
-/

theorem contAt_add {f g : ℝ → ℝ} {a : ℝ} (hf : ContAt f a) (hg : ContAt g a) :
    ContAt (fun x ↦ f x + g x) a := by
  sorry

/-!
### 练习 7.4　复合函数连续
若 $f$ 在 $a$ 连续，$g$ 在 $f(a)$ 连续，则 $g \circ f$ 在 $a$ 连续。

提示：先对 $g$ 用 $\varepsilon$ 得到 $\eta$，再对 $f$ 用 $\eta$ 得到 $\delta$。
-/

theorem contAt_comp {f g : ℝ → ℝ} {a : ℝ} (hf : ContAt f a) (hg : ContAt g (f a)) :
    ContAt (fun x ↦ g (f x)) a := by
  sorry

/-!
### 练习 7.5 ★　平方函数连续
**草稿**：$|x^2 - a^2| = |x - a| \cdot |x + a|$。若 $|x - a| < 1$，
则 $|x + a| \le |x - a| + 2|a| < 2|a| + 1$。所以取
$$\delta = \min\left(1, \frac{\varepsilon}{2|a| + 1}\right).$$
这是"先限制 $\delta \le 1$，控制住其余因子"的经典技巧。

提示：`lt_of_lt_of_le hx (min_le_left _ _)` 得到 `|x - a| < 1`。
最后一步 `ε / (2 * |a| + 1) * (2 * |a| + 1) = ε` 可以用 `field_simp`。
-/

theorem contAt_sq (a : ℝ) : ContAt (fun x ↦ x ^ 2) a := by
  sorry

/-!
### 练习 7.6　连续性的序列刻画
若 $f$ 在 $a$ 连续且 $x_n \to a$，则 $f(x_n) \to f(a)$。

提示：对 $f$ 用 $\varepsilon$ 得到 $\delta$，再对数列用 $\delta$ 得到 $N$。
-/

theorem contAt_seq {f : ℝ → ℝ} {a : ℝ} {x : ℕ → ℝ} (hf : ContAt f a) (hx : TendsTo x a) :
    TendsTo (fun n ↦ f (x n)) (f a) := by
  sorry

/-!
### 练习 7.7 ★　数乘
若 $f$ 在 $a$ 连续，则 $c f$ 在 $a$ 连续。

提示：对 $f$ 使用 $\frac{\varepsilon}{|c| + 1}$。
-/

theorem contAt_const_mul {f : ℝ → ℝ} {a : ℝ} (c : ℝ) (hf : ContAt f a) :
    ContAt (fun x ↦ c * f x) a := by
  sorry

/-!
## 7.5 与 Mathlib 接轨

Mathlib 中 $f$ 在 $a$ 连续写作 `ContinuousAt f a`，在每一点连续写作 `Continuous f`。
`Metric.continuousAt_iff` 把它翻译成 $\varepsilon$-$\delta$ 语言（用 `dist x a` 表示 $|x - a|$）。
-/

theorem contAt_iff_continuousAt (f : ℝ → ℝ) (a : ℝ) : ContAt f a ↔ ContinuousAt f a := by
  rw [Metric.continuousAt_iff]
  simp only [Real.dist_eq, ContAt]

/-!
Mathlib 还提供了自动证明连续性的策略 `fun_prop`：对由连续函数通过加、乘、复合等
构造出来的函数，它会自动完成证明。下一章我们会大量使用它。
-/

example : Continuous (fun x : ℝ ↦ x ^ 3 - 2 * x + 1) := by
  fun_prop

/-!
## 本章小结

* $\varepsilon$-$\delta$：根据输出误差 $\varepsilon$ 构造输入误差 $\delta$；
* 多个条件同时成立时，取 $\delta = \min(\delta_1, \delta_2)$；
* 非线性函数：先限制 $\delta \le 1$ 控制其余因子；
* 连续 ⇔ 保持数列极限（海涅定理）。
-/

end Ch07
