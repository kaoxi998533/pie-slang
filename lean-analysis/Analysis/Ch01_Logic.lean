import Mathlib

/-!
# 第 1 章　逻辑：命题、量词与证明

数列极限的定义是

$$\forall \varepsilon > 0,\ \exists N,\ \forall n \ge N,\ |a_n - L| < \varepsilon.$$

这一句话里有三个量词、一个蕴含和一个不等式。要熟练地证明分析中的定理，
首先要熟练地**拆解**和**构造**这些逻辑结构。本章介绍每一种逻辑联结词在
Lean 中对应的策略。

在 Lean 中，命题的类型是 `Prop`。`P : Prop` 表示"$P$ 是一个命题"，
`h : P` 表示"$h$ 是 $P$ 的一个证明"。

## 1.1 蕴含 `→`

* **证明** $P \to Q$：用 `intro h` 假设 $P$ 成立（得到 `h : P`），然后证明 $Q$。
* **使用** `h : P → Q`：如果目标是 $Q$，`apply h` 会把目标变成 $P$；
  也可以直接写 `h hp` 得到 $Q$ 的证明（这里 `hp : P`）。
-/

namespace Ch01

example (P Q : Prop) (hP : P) (hPQ : P → Q) : Q := by
  apply hPQ
  exact hP

example (P Q : Prop) (hQ : Q) : P → Q := by
  intro _hP
  exact hQ

/-!
## 1.2 合取 `∧`（并且）

* **证明** $P \land Q$：`constructor` 把目标拆成两个，分别证明 $P$ 与 $Q$。
  两个子目标用 `·` 分开书写（输入 `\.`）。
* **使用** `h : P ∧ Q`：`obtain ⟨hP, hQ⟩ := h` 把它拆成两个假设。
  也可以用 `h.1` 和 `h.2`。（尖括号 `⟨ ⟩` 输入 `\<` 和 `\>`。）
-/

example (x : ℝ) (h : 0 < x ∧ x < 1) : x ^ 2 < 1 := by
  obtain ⟨h₀, h₁⟩ := h
  nlinarith

example (x : ℝ) (hx : 2 < x) : 0 < x ∧ 1 < x := by
  constructor
  · linarith
  · linarith

/-!
## 1.3 析取 `∨`（或者）

* **证明** $P \lor Q$：用 `left` 选择证明 $P$，或用 `right` 选择证明 $Q$。
* **使用** `h : P ∨ Q`：`rcases h with hP | hQ` 进行**分类讨论**，
  得到两个目标，分别在 $P$ 成立和 $Q$ 成立的情况下证明。
-/

example (x : ℝ) (h : x < -1 ∨ 1 < x) : 1 < x ^ 2 := by
  rcases h with h | h
  · nlinarith
  · nlinarith

/-!
`le_or_gt a b : a ≤ b ∨ b < a` 是一个非常有用的定理，它让我们可以对任意两个实数
进行分类讨论。
-/

example (x : ℝ) : 0 ≤ x ^ 2 := by
  rcases le_or_gt 0 x with h | h
  · positivity
  · nlinarith

/-!
## 1.4 全称量词 `∀`（对任意）

* **证明** $\forall x, P(x)$：`intro x`，然后对这个"任意的" $x$ 证明 $P(x)$。
* **使用** `h : ∀ x, P x`：`h a` 就是 $P(a)$ 的证明。也可以 `specialize h a`。

注意：$\forall \varepsilon > 0, P(\varepsilon)$ 是 $\forall \varepsilon, \varepsilon > 0 \to P(\varepsilon)$
的缩写，因此要用 `intro ε hε` 引入**两个**东西。
-/

example (f : ℝ → ℝ) (h : ∀ x, f x ≤ 3) : f 1 + f 2 ≤ 6 := by
  have h1 := h 1
  have h2 := h 2
  linarith

example : ∀ ε > (0 : ℝ), ε / 2 < ε := by
  intro ε hε
  linarith

/-!
## 1.5 存在量词 `∃`（存在）

* **证明** $\exists x, P(x)$：用 `use a` 给出一个具体的 $a$，然后证明 $P(a)$。
  **构造一个合适的对象**——这是分析证明中最需要创造力的一步。
* **使用** `h : ∃ x, P x`：`obtain ⟨x, hx⟩ := h` 取出这样一个 $x$ 以及 `hx : P x`。
-/

example : ∃ x : ℝ, x ^ 2 = 4 := by
  use 2
  norm_num

example (f : ℝ → ℝ) (h : ∃ x, f x > 5) : ∃ x, f x > 4 := by
  obtain ⟨x, hx⟩ := h
  use x
  linarith

/-!
## 1.6 否定 `¬` 与 `push Not`

在 Lean 中 $\lnot P$ 定义为 $P \to \text{False}$。因此：

* **证明** $\lnot P$：`intro h` 假设 $P$，推出矛盾（目标是 `False`）。
* `by_contra h`：**反证法**，假设目标的否定，推出 `False`。
* `push Not`：把否定号"推进"量词内部，例如
  $\lnot (\forall x, P x)$ 变为 $\exists x, \lnot P x$，
  $\lnot (a \le b)$ 变为 $b < a$。

写出一个命题的**否定**，是分析中最基本也最容易出错的技能。
`push Not` 可以帮你检查自己写得对不对。
（在较旧版本的 Mathlib 中，这个策略叫做 `push_neg`。）
-/

example (f : ℝ → ℝ) : ¬ (∀ x, f x ≤ 0) ↔ ∃ x, 0 < f x := by
  push Not
  rfl

/-!
## 1.7 分析中的第一个定义：有上界的函数

我们说函数 $f$ **有上界**，如果存在 $M$ 使得对一切 $x$ 都有 $f(x) \le M$。

下面用 `def` 给出这个定义。之后 Lean 会把 `BddAboveFn f` 理解为
`∃ M, ∀ x, f x ≤ M`，所以可以对它使用 `obtain`、`use` 等策略。
-/

/-- 函数 `f` 有上界。 -/
def BddAboveFn (f : ℝ → ℝ) : Prop :=
  ∃ M, ∀ x, f x ≤ M

/-!
> **定理 1.1**　若 $f$ 与 $g$ 都有上界，则 $f + g$ 也有上界。

**证明思路**：设 $f \le M$，$g \le N$，则 $f + g \le M + N$。
注意 Lean 证明的结构与手写证明完全对应：取出 $M$、取出 $N$、给出 $M + N$、
对任意 $x$ 验证不等式。
-/

theorem bddAboveFn_add {f g : ℝ → ℝ} (hf : BddAboveFn f) (hg : BddAboveFn g) :
    BddAboveFn (fun x ↦ f x + g x) := by
  obtain ⟨M, hM⟩ := hf
  obtain ⟨N, hN⟩ := hg
  use M + N
  intro x
  have h1 := hM x
  have h2 := hN x
  show f x + g x ≤ M + N
  linarith

/-!
上面的 `show` 把目标 `(fun x ↦ f x + g x) x ≤ M + N` 写成更易读的等价形式
`f x + g x ≤ M + N`（两者只差一步代入，Lean 能自动识别）。

> **定理 1.2**　恒等函数 $x \mapsto x$ 没有上界。

**证明**：假设有上界 $M$，取 $x = M + 1$，得到 $M + 1 \le M$，矛盾。
-/

theorem not_bddAboveFn_id : ¬ BddAboveFn (fun x ↦ x) := by
  intro h
  obtain ⟨M, hM⟩ := h
  have h1 : M + 1 ≤ M := hM (M + 1)
  linarith

/-!
## 练习

### 练习 1.1　蕴含的传递性
-/

theorem ex1_1 (P Q R : Prop) (hPQ : P → Q) (hQR : Q → R) : P → R := by
  sorry

/-!
### 练习 1.2　合取的交换律
-/

theorem ex1_2 (P Q : Prop) : P ∧ Q → Q ∧ P := by
  sorry

/-!
### 练习 1.3　析取的交换律
-/

theorem ex1_3 (P Q : Prop) : P ∨ Q → Q ∨ P := by
  sorry

/-!
### 练习 1.4　实数没有最大元
对任意实数 $x$，都存在比它更大的实数。
-/

theorem ex1_4 : ∀ x : ℝ, ∃ y : ℝ, x < y := by
  sorry

/-!
### 练习 1.5　量词的顺序
"存在一个 $x$ 对所有 $y$ 都成立"蕴含"对每个 $y$ 都存在 $x$ 成立"。
（反过来对吗？想一想 $P(x, y) := x > y$。）
-/

theorem ex1_5 (P : ℝ → ℝ → Prop) (h : ∃ x, ∀ y, P x y) : ∀ y, ∃ x, P x y := by
  sorry

/-!
### 练习 1.6　有上界的数乘
若 $f$ 有上界且 $c \ge 0$，则 $c f$ 也有上界。

提示：Mathlib 中的 `mul_le_mul_of_nonneg_left : b ≤ c → 0 ≤ a → a * b ≤ a * c`。
-/

theorem ex1_6 {f : ℝ → ℝ} {c : ℝ} (hc : 0 ≤ c) (hf : BddAboveFn f) :
    BddAboveFn (fun x ↦ c * f x) := by
  sorry

/-!
### 练习 1.7　无上界的刻画
用 `push Not` 写出"没有上界"的含义。先用 `unfold BddAboveFn` 展开定义。
-/

theorem ex1_7 (f : ℝ → ℝ) : ¬ BddAboveFn f ↔ ∀ M, ∃ x, M < f x := by
  sorry

/-!
### 练习 1.8　平方函数没有上界
提示：对给定的 $M$，取 $x = |M| + 1$ 试试，并利用 `le_abs_self M : M ≤ |M|`。
-/

theorem ex1_8 : ¬ BddAboveFn (fun x ↦ x ^ 2) := by
  sorry

/-!
## 本章小结

| 逻辑结构 | 证明它 | 使用它 |
|---------|--------|--------|
| $P \to Q$ | `intro h` | `apply h`、`h hp` |
| $P \land Q$ | `constructor` | `obtain ⟨h₁, h₂⟩ := h` |
| $P \lor Q$ | `left` / `right` | `rcases h with h₁ \| h₂` |
| $\forall x, P x$ | `intro x` | `h a` |
| $\exists x, P x$ | `use a` | `obtain ⟨x, hx⟩ := h` |
| $\lnot P$ | `intro h`（推出 `False`） | `exact absurd hp h` |
| 反证法 | `by_contra h` | |
| 否定化简 | `push Not` | `push Not at h` |
-/

end Ch01
