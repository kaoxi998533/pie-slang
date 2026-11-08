import Mathlib

/-!
# 第 4 章　数列极限

终于来到了分析的核心概念：**极限**。

一个实数数列 $a_0, a_1, a_2, \ldots$ 就是一个函数 $a : \mathbb{N} \to \mathbb{R}$，
在 Lean 中它的类型是 `ℕ → ℝ`，第 $n$ 项写作 `a n`。
（注意：本书的数列从第 $0$ 项开始。）

## 4.1 $\varepsilon$-$N$ 定义

> **定义 4.1**　称数列 $(a_n)$ **收敛**到 $L$，记作 $a_n \to L$，如果
> $$\forall \varepsilon > 0,\ \exists N \in \mathbb{N},\ \forall n \ge N,\ |a_n - L| < \varepsilon.$$

可以把这个定义理解为一个**游戏**：

1. 对手先给出一个误差范围 $\varepsilon > 0$（可能非常小）；
2. 你需要找到一个 $N$；
3. 然后对手可以任选 $n \ge N$，而你必须保证 $|a_n - L| < \varepsilon$。

如果无论对手给出什么 $\varepsilon$，你总能赢，那么 $a_n \to L$。
证明的关键永远是第 2 步：**根据 $\varepsilon$ 构造 $N$**。

<!-- widget:eps-n -->

（网页版在这里有一个交互演示：拖动 $\varepsilon$，观察需要的 $N$ 如何变化。
$\varepsilon$ 越小，$N$ 就要越大——但只要总能找到，数列就收敛。）
-/

namespace Ch04

/-- 数列 `a` 收敛到 `L`。 -/
def TendsTo (a : ℕ → ℝ) (L : ℝ) : Prop :=
  ∀ ε > 0, ∃ N : ℕ, ∀ n ≥ N, |a n - L| < ε

/-!
`TendsTo a L` 的证明总是以 `intro ε hε` 开始（接受对手的 $\varepsilon$），
然后 `use N`（给出你的 $N$），再 `intro n hn`（接受对手的 $n$），最后估计不等式。

## 4.2 第一个例子：$\frac{1}{n+1} \to 0$

> **例 4.2**　$\lim\limits_{n \to \infty} \frac{1}{n + 1} = 0$。

**草稿**（找 $N$）：我们需要 $\frac{1}{n+1} < \varepsilon$，即 $n + 1 > \frac{1}{\varepsilon}$。
由阿基米德性质，取自然数 $N > \frac{1}{\varepsilon}$，则 $n \ge N$ 时 $n + 1 > N > \frac{1}{\varepsilon}$。

**证明**：下面每一步都与草稿对应。`exists_nat_gt` 是 Mathlib 中的阿基米德性质，
`exact_mod_cast` 把自然数的不等式 `N ≤ n` 转化为实数的不等式。
-/

theorem tendsTo_one_div : TendsTo (fun n ↦ 1 / ((n : ℝ) + 1)) 0 := by
  intro ε hε
  obtain ⟨N, hN⟩ := exists_nat_gt (1 / ε)
  use N
  intro n hn
  have hn' : (N : ℝ) ≤ n := by exact_mod_cast hn
  have hpos : (0 : ℝ) < n + 1 := by positivity
  show |1 / ((n : ℝ) + 1) - 0| < ε
  rw [sub_zero, abs_of_pos (by positivity), div_lt_iff₀ hpos]
  -- 目标：1 < ε * (n + 1)
  have h1 : 1 / ε < n + 1 := by linarith
  rw [div_lt_iff₀ hε] at h1
  linarith

/-!
## 4.3 极限的唯一性

> **定理 4.3**　收敛数列的极限是唯一的：若 $a_n \to L$ 且 $a_n \to M$，则 $L = M$。

**证明**：反证，设 $L \ne M$，令 $\varepsilon = \frac{|L - M|}{2} > 0$。
取 $N_1$、$N_2$ 分别对应两个极限，令 $n = \max(N_1, N_2)$，则
$$|L - M| \le |L - a_n| + |a_n - M| < \varepsilon + \varepsilon = |L - M|,$$
矛盾。直观上：$a_n$ 不可能同时离两个不同的点都很近。
-/

theorem tendsTo_unique {a : ℕ → ℝ} {L M : ℝ} (hL : TendsTo a L) (hM : TendsTo a M) :
    L = M := by
  by_contra hne
  have hpos : 0 < |L - M| := abs_pos.mpr (sub_ne_zero.mpr hne)
  obtain ⟨N₁, hN₁⟩ := hL (|L - M| / 2) (by positivity)
  obtain ⟨N₂, hN₂⟩ := hM (|L - M| / 2) (by positivity)
  let n := max N₁ N₂
  have h1 : |a n - L| < |L - M| / 2 := hN₁ n (le_max_left _ _)
  have h2 : |a n - M| < |L - M| / 2 := hN₂ n (le_max_right _ _)
  have h3 : |L - M| ≤ |a n - L| + |a n - M| := by
    calc |L - M| = |(L - a n) + (a n - M)| := by congr 1; ring
      _ ≤ |L - a n| + |a n - M| := abs_add_le _ _
      _ = |a n - L| + |a n - M| := by rw [abs_sub_comm L]
  linarith

/-!
这里的 `max N₁ N₂` 是一个反复出现的技巧：**同时满足多个"充分大"的条件**，
只需取各个 $N$ 的最大值。`le_max_left a b : a ≤ max a b`。

## 练习

### 练习 4.1　常数列
常数列 $c, c, c, \ldots$ 收敛到 $c$。任何 $N$ 都可以！

提示：`intro ε hε`，`use 0`，`intro n _`，然后 `simp`（它知道 $|c - c| = 0$），
可能还需要 `exact hε` 或 `simpa using hε`。
-/

theorem tendsTo_const (c : ℝ) : TendsTo (fun _ ↦ c) c := by
  sorry

/-!
### 练习 4.2　加上常数
若 $a_n \to L$，则 $a_n + c \to L + c$。

提示：用 `obtain ⟨N, hN⟩ := h ε hε` 取出 $N$。
注意 $a_n + c - (L + c) = a_n - L$。
-/

theorem tendsTo_add_const {a : ℕ → ℝ} {L : ℝ} (c : ℝ) (h : TendsTo a L) :
    TendsTo (fun n ↦ a n + c) (L + c) := by
  sorry

/-!
### 练习 4.3　取负
若 $a_n \to L$，则 $-a_n \to -L$。

提示：$|-x| = |x|$ 在 Mathlib 中叫 `abs_neg`。
-/

theorem tendsTo_neg {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ -a n) (-L) := by
  sorry

/-!
### 练习 4.4　平移不改变极限
若 $a_n \to L$，则对任意固定的 $k$，$a_{n+k} \to L$。（去掉有限项不影响收敛性。）

提示：`omega` 策略可以自动证明自然数之间的线性不等式，如 `n + k ≥ N`。
-/

theorem tendsTo_shift {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) (k : ℕ) :
    TendsTo (fun n ↦ a (n + k)) L := by
  sorry

/-!
### 练习 4.5　综合运用
利用例 4.2 和练习 4.2，证明 $\frac{1}{n+1} + 3 \to 3$。

提示：`have h := tendsTo_add_const 3 tendsTo_one_div` 得到的极限是 `0 + 3`，
用 `simpa using h` 化简。
-/

theorem ex4_5 : TendsTo (fun n ↦ 1 / ((n : ℝ) + 1) + 3) 3 := by
  sorry

/-!
### 练习 4.6　极限为零的等价刻画
$a_n \to L$ 当且仅当 $a_n - L \to 0$。

提示：`constructor` 分两个方向。每个方向里都是"同一个 $N$ 就行"。
-/

theorem tendsTo_iff_sub {a : ℕ → ℝ} {L : ℝ} :
    TendsTo a L ↔ TendsTo (fun n ↦ a n - L) 0 := by
  sorry

/-!
### 练习 4.7 ★　绝对值的极限
若 $a_n \to L$，则 $|a_n| \to |L|$。反过来对吗？（想一想 $a_n = (-1)^n$。）

提示：Mathlib 中的 `abs_abs_sub_abs_le_abs_sub x y : |(|x| - |y|)| ≤ |x - y|`。
-/

theorem tendsTo_abs {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ |a n|) |L| := by
  sorry

/-!
### 练习 4.8 ★　发散数列
数列 $0, 1, 0, 1, \ldots$ 不收敛到 $0$。

提示：取 $\varepsilon = \frac12$。对手给出 $N$ 之后，考虑奇数项 $n = 2N + 1$，
它 $\ge N$ 且值为 $1$。`Nat.odd_iff`、`Nat.add_mod` 等可能有用，
不过最简单的是用 `show` 和 `simp` 计算 `(2 * N + 1) % 2 = 1`。
-/

theorem ex4_8 : ¬ TendsTo (fun n ↦ if n % 2 = 0 then 0 else 1) 0 := by
  sorry

/-!
## 4.4 与 Mathlib 接轨

Mathlib 用更一般的**滤子**（filter）语言描述极限：
`Filter.Tendsto a Filter.atTop (nhds L)`，读作"当 $n$ 趋于无穷时 $a_n$ 趋于 $L$"。
下面的定理说明它与我们的定义完全一致。之后的章节中，
当我们需要 Mathlib 里更深的定理时，就通过它来"翻译"。
-/

theorem tendsTo_iff_tendsto (a : ℕ → ℝ) (L : ℝ) :
    TendsTo a L ↔ Filter.Tendsto a Filter.atTop (nhds L) := by
  rw [Metric.tendsto_atTop]
  simp only [Real.dist_eq, TendsTo]

/-!
## 本章小结

* 极限的定义是一个 $\varepsilon$-$N$ 游戏：拿到 $\varepsilon$，构造 $N$；
* 先在草稿上倒推出 $N$ 应满足的条件，再写正式证明；
* 多个条件同时成立：取 $N = \max(N_1, N_2)$；
* 证明"任意接近"时，阿基米德性质是找 $N$ 的标准工具。
-/

end Ch04
