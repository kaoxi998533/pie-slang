import Mathlib

/-!
# 第 6 章　单调收敛定理与柯西收敛准则

前两章中，我们总是**先知道极限 $L$ 是多少**，再去验证 $a_n \to L$。
但很多时候极限的值根本写不出来，例如
$$a_n = 1 + \frac{1}{1!} + \frac{1}{2!} + \cdots + \frac{1}{n!}.$$
我们需要一些**不依赖于极限值**就能判断收敛的准则。本章的两个准则——
单调收敛定理与柯西准则——都是实数完备性的直接推论。

## 6.0 预备
-/

namespace Ch06

/-- 数列 `a` 收敛到 `L`。 -/
def TendsTo (a : ℕ → ℝ) (L : ℝ) : Prop :=
  ∀ ε > 0, ∃ N : ℕ, ∀ n ≥ N, |a n - L| < ε

theorem tendsTo_neg {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ -a n) (-L) := by
  intro ε hε
  obtain ⟨N, hN⟩ := h ε hε
  use N
  intro n hn
  have h1 := hN n hn
  show |-a n - -L| < ε
  rw [show -a n - -L = -(a n - L) by ring, abs_neg]
  exact h1

/-!
## 6.1 单调数列

> **定义 6.1**　数列 $(a_n)$ 称为**单调递增**的，如果对所有 $n$，$a_n \le a_{n+1}$。

定义中只比较相邻两项。由归纳法立刻得到：$m \le n$ 时 $a_m \le a_n$。
-/

/-- 数列 `a` 单调递增。 -/
def Increasing (a : ℕ → ℝ) : Prop :=
  ∀ n, a n ≤ a (n + 1)

theorem Increasing.mono {a : ℕ → ℝ} (h : Increasing a) {m n : ℕ} (hmn : m ≤ n) :
    a m ≤ a n := by
  induction hmn with
  | refl => exact le_rfl
  | step _ ih => exact le_trans ih (h _)

/-!
这里的 `induction hmn` 是对"$m \le n$ 的证明"做归纳：要么 $n = m$（`refl`），
要么 $n = k + 1$ 且 $m \le k$（`step`，`ih : a m ≤ a k`）。

## 6.2 单调收敛定理

> **定理 6.2（单调收敛定理）**　单调递增且有上界的数列必收敛，且极限就是
> $\sup\{a_n\}$。

**证明**：令 $S = \{a_n \mid n \in \mathbb{N}\}$，它非空且有上界，由确界原理有 $s = \sup S$。
任给 $\varepsilon > 0$，$s - \varepsilon$ 不是上界，故存在 $N$ 使 $a_N > s - \varepsilon$。
由单调性，$n \ge N$ 时
$$s - \varepsilon < a_N \le a_n \le s,$$
从而 $|a_n - s| < \varepsilon$。

我们直接使用 Mathlib 中的上确界 `sSup`：
`le_csSup` 说 `sSup S` 是上界，`exists_lt_of_lt_csSup` 说比 `sSup S` 小的数不是上界。
-/

theorem monotone_convergence {a : ℕ → ℝ} (h : Increasing a) (hb : ∃ M, ∀ n, a n ≤ M) :
    ∃ L, TendsTo a L := by
  obtain ⟨M, hM⟩ := hb
  have hne : (Set.range a).Nonempty := ⟨a 0, 0, rfl⟩
  have hbdd : BddAbove (Set.range a) := by
    use M
    rintro _ ⟨n, rfl⟩
    exact hM n
  refine ⟨sSup (Set.range a), ?_⟩
  intro ε hε
  have hlt : sSup (Set.range a) - ε < sSup (Set.range a) := by linarith
  obtain ⟨_, ⟨N, rfl⟩, hN⟩ := exists_lt_of_lt_csSup hne hlt
  use N
  intro n hn
  have h1 : a N ≤ a n := h.mono hn
  have h2 : a n ≤ sSup (Set.range a) := le_csSup hbdd ⟨n, rfl⟩
  rw [abs_lt]
  constructor <;> linarith

/-!
`constructor <;> linarith` 的意思是：`constructor` 产生的**每一个**目标都用 `linarith` 解决。

## 6.3 柯西数列

> **定义 6.3**　数列 $(a_n)$ 称为**柯西数列**，如果
> $$\forall \varepsilon > 0,\ \exists N,\ \forall m, n \ge N,\ |a_m - a_n| < \varepsilon.$$

柯西条件只涉及数列**自身**的项：充分靠后的项彼此任意接近。
它不需要事先知道极限是什么。
-/

/-- 数列 `a` 是柯西数列。 -/
def IsCauchy (a : ℕ → ℝ) : Prop :=
  ∀ ε > 0, ∃ N : ℕ, ∀ m ≥ N, ∀ n ≥ N, |a m - a n| < ε

/-!
> **定理 6.4（柯西收敛准则）**　实数列收敛当且仅当它是柯西数列。

"收敛 ⇒ 柯西"是练习 6.3。"柯西 ⇒ 收敛"正是实数完备性的另一种表述
（在 $\mathbb{Q}$ 中它不成立：$\sqrt 2$ 的十进制近似 $1, 1.4, 1.41, 1.414, \ldots$ 是柯西数列，
但在 $\mathbb{Q}$ 中没有极限）。它的证明较长（需要先证明柯西数列有界，再用
Bolzano–Weierstrass 定理取收敛子列），这里我们借用 Mathlib：
`Metric.cauchySeq_iff` 把 Mathlib 的 `CauchySeq` 翻译为 $\varepsilon$-$N$ 形式，
`cauchySeq_tendsto_of_complete` 说完备空间中的柯西列收敛。
-/

theorem cauchy_complete {a : ℕ → ℝ} (h : IsCauchy a) : ∃ L, TendsTo a L := by
  have hc : CauchySeq a := by
    rw [Metric.cauchySeq_iff]
    intro ε hε
    obtain ⟨N, hN⟩ := h ε hε
    exact ⟨N, fun m hm n hn ↦ by rw [Real.dist_eq]; exact hN m hm n hn⟩
  obtain ⟨L, hL⟩ := cauchySeq_tendsto_of_complete hc
  refine ⟨L, ?_⟩
  rw [Metric.tendsto_atTop] at hL
  intro ε hε
  obtain ⟨N, hN⟩ := hL ε hε
  exact ⟨N, fun n hn ↦ by rw [← Real.dist_eq]; exact hN n hn⟩

/-!
## 练习

### 练习 6.1　一个单调数列
证明 $a_n = 1 - \frac{1}{n + 1}$ 单调递增。

提示：`intro n`，用 `show` 写出目标，再 `push_cast`。
`gcongr` 可以把 `1 - 1 / x ≤ 1 - 1 / y` 化为 `x ≤ y`。
-/

theorem ex6_1 : Increasing (fun n ↦ 1 - 1 / ((n : ℝ) + 1)) := by
  sorry

/-!
### 练习 6.2 ★　单调递减数列
单调递减且有下界的数列必收敛。

提示：对 $-a_n$ 使用单调收敛定理，再用 `tendsTo_neg`。
最后得到的是 `TendsTo (fun n ↦ -(-a n)) (-L)`，用 `simpa using` 化简。
-/

theorem ex6_2 {a : ℕ → ℝ} (h : ∀ n, a (n + 1) ≤ a n) (hb : ∃ m, ∀ n, m ≤ a n) :
    ∃ L, TendsTo a L := by
  sorry

/-!
### 练习 6.3　收敛数列是柯西数列
提示：$\frac{\varepsilon}{2}$ 技巧。$|a_m - a_n| \le |a_m - L| + |L - a_n|$。
-/

theorem ex6_3 {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) : IsCauchy a := by
  sorry

/-!
### 练习 6.4 ★　柯西数列有界
提示：取 $\varepsilon = 1$ 得到 $N$，则 $n \ge N$ 时 $|a_n| \le |a_N| + 1$。
前 $N$ 项可以用有限和处理（参考第 5 章定理 5.3 的证明），
也可以利用"柯西 ⇒ 收敛 ⇒ 有界"。这里给出的有界形式比第 5 章简单：不要求 $C > 0$。
-/

theorem ex6_4 {a : ℕ → ℝ} (h : IsCauchy a) : ∃ C, ∀ n, |a n| ≤ C := by
  sorry

/-!
### 练习 6.5 ★★　有界单调数列的极限即上确界
若递增数列 $a_n \to L$，则对所有 $n$，$a_n \le L$。

提示：反证。若 $a_k > L$，取 $\varepsilon = a_k - L$，对足够大的 $n \ge k$，
$a_n \ge a_k$，与 $|a_n - L| < a_k - L$ 矛盾。取 $n = \max(N, k)$。
-/

theorem ex6_5 {a : ℕ → ℝ} {L : ℝ} (h : Increasing a) (hL : TendsTo a L) :
    ∀ n, a n ≤ L := by
  sorry

/-!
## 本章小结

* 单调收敛定理：单调有界 ⇒ 收敛，极限是上（下）确界；
* 柯西准则：收敛 ⇔ 柯西，判断收敛无需知道极限；
* 这两个定理都等价于实数的完备性，是区分 $\mathbb{R}$ 与 $\mathbb{Q}$ 的关键。

**延伸阅读**：Bolzano–Weierstrass 定理（有界数列必有收敛子列）
在 Mathlib 中是 `tendsto_subseq_of_bounded`。
-/

end Ch06
