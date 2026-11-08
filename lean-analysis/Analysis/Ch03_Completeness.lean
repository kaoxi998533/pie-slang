import Mathlib

/-!
# 第 3 章　完备性：上确界与阿基米德性质

有理数 $\mathbb{Q}$ 也是有序域，为什么分析要建立在实数上？因为 $\mathbb{Q}$ 有"空隙"：
集合 $\{x \in \mathbb{Q} \mid x^2 < 2\}$ 在 $\mathbb{Q}$ 中有上界，却没有**最小的**上界——
那个位置本该是 $\sqrt{2}$。实数填补了所有这样的空隙，这就是**完备性**。

完备性是数学分析的基石：单调收敛定理、柯西收敛准则、介值定理、最值定理……
全都依赖于它。

## 3.1 上界与上确界
-/

namespace Ch03

/-- `M` 是集合 `S` 的一个上界。 -/
def IsUB (S : Set ℝ) (M : ℝ) : Prop :=
  ∀ x ∈ S, x ≤ M

/-- `s` 是集合 `S` 的上确界（最小上界）。 -/
def IsSup (S : Set ℝ) (s : ℝ) : Prop :=
  IsUB S s ∧ ∀ M, IsUB S M → s ≤ M

/-!
> **定义 3.1**　$s$ 称为 $S$ 的**上确界**，记作 $s = \sup S$，如果
> (1) $s$ 是 $S$ 的上界；(2) $S$ 的任意上界 $M$ 都满足 $s \le M$。

`Set ℝ` 是实数集合的类型，`x ∈ S` 表示 $x$ 属于 $S$。
集合 `{x : ℝ | x < 1}` 的写法与数学中完全一样。

## 3.2 完备性公理

> **公理 3.2（确界原理）**　$\mathbb{R}$ 的每个**非空**、**有上界**的子集都有上确界。

在 Mathlib 中，实数是由柯西序列严格构造出来的，确界原理是一个**定理**。
它的上确界记作 `sSup S`。下面我们把 Mathlib 的版本翻译成本章的定义——
从此以后，我们只通过 `completeness` 来使用实数的完备性。
-/

theorem completeness (S : Set ℝ) (hne : S.Nonempty) (hbdd : ∃ M, IsUB S M) :
    ∃ s, IsSup S s := by
  obtain ⟨M, hM⟩ := hbdd
  have hb : BddAbove S := ⟨M, hM⟩
  refine ⟨sSup S, ?_, ?_⟩
  · intro x hx
    exact le_csSup hb hx
  · intro M' hM'
    exact csSup_le hne hM'

/-!
`refine ⟨sSup S, ?_, ?_⟩` 的意思是：给出 $s = $ `sSup S`，
而 `IsSup` 的两个条件留作两个新目标（`?_` 表示"待证"）。

## 3.3 上确界的 $\varepsilon$ 刻画

上确界的第二个条件"是最小的上界"有一个等价的、更好用的形式：

> **定理 3.3**　若 $s = \sup S$，则对任意 $\varepsilon > 0$，存在 $x \in S$ 使得 $x > s - \varepsilon$。

也就是说，$s - \varepsilon$ 不再是上界，$S$ 中的元素可以任意接近 $s$。

**证明**：反证。若对所有 $x \in S$ 都有 $x \le s - \varepsilon$，则 $s - \varepsilon$ 是上界，
于是 $s \le s - \varepsilon$，与 $\varepsilon > 0$ 矛盾。
-/

theorem IsSup.approx {S : Set ℝ} {s : ℝ} (hs : IsSup S s) :
    ∀ ε > 0, ∃ x ∈ S, s - ε < x := by
  intro ε hε
  by_contra h
  push Not at h
  have h1 : s ≤ s - ε := hs.2 (s - ε) h
  linarith

/-!
注意定理名字 `IsSup.approx`：这样命名之后，若 `hs : IsSup S s`，
就可以用 `hs.approx` 调用它。

## 3.4 阿基米德性质

> **定理 3.4（阿基米德性质）**　对任意实数 $x$，存在自然数 $n > x$。

这看起来"显然"，但它其实依赖于完备性（存在不满足阿基米德性质的有序域！）。

**证明**：反证。若所有自然数都 $\le x$，则 $\mathbb{N}$ 有上界，设其上确界为 $s$。
由定理 3.3，存在自然数 $n > s - 1$，于是 $n + 1 > s$。但 $n + 1$ 也是自然数，
应当 $\le s$，矛盾。

`(n : ℝ)` 表示把自然数 `n` 看作实数（称为**类型转换**，coercion）。
`Set.range (fun n : ℕ ↦ (n : ℝ))` 就是集合 $\{0, 1, 2, \ldots\} \subseteq \mathbb{R}$。
`push_cast` 负责化简类型转换，例如把 `((n + 1 : ℕ) : ℝ)` 变成 `(n : ℝ) + 1`。
-/

theorem archimedean (x : ℝ) : ∃ n : ℕ, x < n := by
  by_contra h
  push Not at h
  -- h : ∀ (n : ℕ), (n : ℝ) ≤ x
  have hne : (Set.range (fun n : ℕ ↦ (n : ℝ))).Nonempty := ⟨_, 0, rfl⟩
  have hbdd : ∃ M, IsUB (Set.range (fun n : ℕ ↦ (n : ℝ))) M := by
    use x
    rintro _ ⟨n, rfl⟩
    exact h n
  obtain ⟨s, hs⟩ := completeness _ hne hbdd
  obtain ⟨_, ⟨n, rfl⟩, hn⟩ := hs.approx 1 one_pos
  have h1 : ((n + 1 : ℕ) : ℝ) ≤ s := hs.1 _ ⟨n + 1, rfl⟩
  push_cast at h1
  linarith

/-!
`rintro _ ⟨n, rfl⟩` 是 `intro` 与 `obtain` 的组合：引入 `range` 中的一个元素，
它必然形如 `(n : ℝ)`，`rfl` 直接把它代入。

Mathlib 中同样的结论叫做 `exists_nat_gt`。

## 练习

### 练习 3.1　上确界的唯一性
提示：`IsSup` 是一个合取，`h₁.1` 是"上界"，`h₁.2` 是"最小"。用 `le_antisymm`。
-/

theorem ex3_1 {S : Set ℝ} {s t : ℝ} (h₁ : IsSup S s) (h₂ : IsSup S t) : s = t := by
  sorry

/-!
### 练习 3.2　最大元就是上确界
若 $s \in S$ 且 $s$ 是 $S$ 的上界，则 $s = \sup S$。
-/

theorem ex3_2 {S : Set ℝ} {s : ℝ} (hs : s ∈ S) (hub : IsUB S s) : IsSup S s := by
  sorry

/-!
### 练习 3.3　开区间的上确界
$\sup \{x \mid x < 1\} = 1$。注意 $1$ 并不属于这个集合！

提示：第二部分用反证法。若某个上界 $M < 1$，考虑 $\frac{M + 1}{2}$。
`x ∈ {x : ℝ | x < 1}` 与 `x < 1` 是一回事，可以用 `show` 转换。
-/

theorem ex3_3 : IsSup {x : ℝ | x < 1} 1 := by
  sorry

/-!
### 练习 3.4　$\frac{1}{n+1}$ 可以任意小
对任意 $\varepsilon > 0$，存在自然数 $n$ 使 $\frac{1}{n + 1} < \varepsilon$。
这是证明数列极限时最常用的工具之一。

提示：对 $\frac{1}{\varepsilon}$ 用阿基米德性质。`div_lt_iff₀ (hc : 0 < c) : a / c < b ↔ a < b * c`
可以帮你去分母。
-/

theorem ex3_4 (ε : ℝ) (hε : 0 < ε) : ∃ n : ℕ, 1 / ((n : ℝ) + 1) < ε := by
  sorry

/-!
### 练习 3.5 ★　下确界
仿照上确界，定义下界和下确界，并证明：若 $s = \sup S$，则 $-s$ 是集合
$-S = \{-x \mid x \in S\}$ 的下确界。这里我们用 `{y | -y ∈ S}` 来表示 $-S$。
-/

/-- `m` 是集合 `S` 的一个下界。 -/
def IsLB (S : Set ℝ) (m : ℝ) : Prop :=
  ∀ x ∈ S, m ≤ x

/-- `s` 是集合 `S` 的下确界（最大下界）。 -/
def IsInf (S : Set ℝ) (s : ℝ) : Prop :=
  IsLB S s ∧ ∀ m, IsLB S m → m ≤ s

theorem ex3_5 {S : Set ℝ} {s : ℝ} (h : IsSup S s) : IsInf {y | -y ∈ S} (-s) := by
  sorry

/-!
## 本章小结

* 确界原理：非空有上界的实数集有上确界（`completeness`，Mathlib 中为 `sSup`）；
* 上确界的 $\varepsilon$ 刻画：$\sup S - \varepsilon$ 不是上界；
* 阿基米德性质：自然数无上界，从而 $\frac{1}{n+1}$ 可以任意小。

这些就是下一章证明数列极限所需的全部"弹药"。
-/

end Ch03
