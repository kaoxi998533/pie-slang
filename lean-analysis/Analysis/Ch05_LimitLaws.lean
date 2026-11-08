import Mathlib

/-!
# 第 5 章　极限的运算法则与夹逼定理

逐个用 $\varepsilon$-$N$ 定义验证极限是很累的。本章证明一组**运算法则**：
和、差、数乘、乘积的极限等于极限的和、差、数乘、乘积。有了它们，
复杂数列的极限就可以由简单数列"组装"出来。

## 5.0 预备：上一章的定义与结论

每一章都是一个独立的 Lean 文件，所以我们先把需要的定义和定理复制过来。
（它们的证明就是上一章练习的参考答案。）
-/

namespace Ch05

/-- 数列 `a` 收敛到 `L`。 -/
def TendsTo (a : ℕ → ℝ) (L : ℝ) : Prop :=
  ∀ ε > 0, ∃ N : ℕ, ∀ n ≥ N, |a n - L| < ε

theorem tendsTo_const (c : ℝ) : TendsTo (fun _ ↦ c) c := by
  intro ε hε
  use 0
  intro n _
  simpa using hε

theorem tendsTo_one_div : TendsTo (fun n ↦ 1 / ((n : ℝ) + 1)) 0 := by
  intro ε hε
  obtain ⟨N, hN⟩ := exists_nat_gt (1 / ε)
  use N
  intro n hn
  have hn' : (N : ℝ) ≤ n := by exact_mod_cast hn
  have hpos : (0 : ℝ) < n + 1 := by positivity
  show |1 / ((n : ℝ) + 1) - 0| < ε
  rw [sub_zero, abs_of_pos (by positivity), div_lt_iff₀ hpos]
  have h1 : 1 / ε < n + 1 := by linarith
  rw [div_lt_iff₀ hε] at h1
  linarith

/-!
## 5.1 和的极限：$\frac{\varepsilon}{2}$ 技巧

> **定理 5.1**　若 $a_n \to L$，$b_n \to M$，则 $a_n + b_n \to L + M$。

**草稿**：由三角不等式
$$|(a_n + b_n) - (L + M)| \le |a_n - L| + |b_n - M|.$$
只要让右边两项**各自**小于 $\frac{\varepsilon}{2}$ 即可。所以对 $a$ 和 $b$ 分别使用
$\frac{\varepsilon}{2}$ 得到 $N_1, N_2$，再取 $N = \max(N_1, N_2)$。

这就是著名的 **$\frac{\varepsilon}{2}$ 技巧**：对手给你 $\varepsilon$，
你把它"分给"两个子问题。
-/

theorem tendsTo_add {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M) :
    TendsTo (fun n ↦ a n + b n) (L + M) := by
  intro ε hε
  obtain ⟨N₁, hN₁⟩ := ha (ε / 2) (by positivity)
  obtain ⟨N₂, hN₂⟩ := hb (ε / 2) (by positivity)
  use max N₁ N₂
  intro n hn
  have h1 := hN₁ n (le_of_max_le_left hn)
  have h2 := hN₂ n (le_of_max_le_right hn)
  calc |a n + b n - (L + M)| = |(a n - L) + (b n - M)| := by congr 1; ring
    _ ≤ |a n - L| + |b n - M| := abs_add_le _ _
    _ < ε / 2 + ε / 2 := by linarith
    _ = ε := by ring

/-!
## 5.2 收敛数列有界

> **定义 5.2**　数列 $(a_n)$ **有界**，如果存在 $C > 0$ 使得对所有 $n$，$|a_n| \le C$。

> **定理 5.3**　收敛数列必有界。

**证明**：取 $\varepsilon = 1$ 得到 $N$。当 $n \ge N$ 时 $|a_n| \le |a_n - L| + |L| < 1 + |L|$。
前面只有有限项 $a_0, \ldots, a_{N-1}$，它们的绝对值之和也是有限的。令
$$C = |a_0| + |a_1| + \cdots + |a_{N-1}| + |L| + 1,$$
则所有项都满足 $|a_n| \le C$。

在 Lean 中，有限和 $\sum_{i < N} |a_i|$ 写作 `∑ i ∈ Finset.range N, |a i|`。
`Finset.single_le_sum` 说的是：非负项的和不小于其中任何一项。
-/

/-- 数列 `a` 有界。 -/
def SeqBounded (a : ℕ → ℝ) : Prop :=
  ∃ C, 0 < C ∧ ∀ n, |a n| ≤ C

theorem bounded_of_tendsTo {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) : SeqBounded a := by
  obtain ⟨N, hN⟩ := h 1 one_pos
  have hsum : 0 ≤ ∑ i ∈ Finset.range N, |a i| :=
    Finset.sum_nonneg (fun i _ ↦ abs_nonneg (a i))
  refine ⟨(∑ i ∈ Finset.range N, |a i|) + |L| + 1, by positivity, ?_⟩
  intro n
  rcases lt_or_ge n N with hn | hn
  · -- 情形 n < N：|a n| 是有限和中的一项
    have : |a n| ≤ ∑ i ∈ Finset.range N, |a i| :=
      Finset.single_le_sum (f := fun i ↦ |a i|) (fun i _ ↦ abs_nonneg (a i))
        (Finset.mem_range.mpr hn)
    linarith [abs_nonneg L]
  · -- 情形 n ≥ N：|a n| ≤ |a n - L| + |L| < 1 + |L|
    have h1 := hN n hn
    have h2 : |a n| ≤ |a n - L| + |L| := by
      calc |a n| = |(a n - L) + L| := by congr 1; ring
        _ ≤ |a n - L| + |L| := abs_add_le _ _
    linarith

/-!
## 5.3 乘积的极限

> **定理 5.4**　若 $a_n \to L$，$b_n \to M$，则 $a_n b_n \to LM$。

**草稿**：关键是"加一项减一项"：
$$a_n b_n - LM = a_n(b_n - M) + M(a_n - L).$$
由定理 5.3，$|a_n| \le C$。于是
$$|a_n b_n - LM| \le C|b_n - M| + |M||a_n - L|.$$
让 $|b_n - M| < \frac{\varepsilon}{2C}$，$|a_n - L| < \frac{\varepsilon}{2(|M|+1)}$
（用 $|M| + 1$ 而不是 $|M|$，是为了避免 $M = 0$ 时除以零）。
-/

theorem tendsTo_mul {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M) :
    TendsTo (fun n ↦ a n * b n) (L * M) := by
  obtain ⟨C, hC, hCa⟩ := bounded_of_tendsTo ha
  intro ε hε
  obtain ⟨N₁, hN₁⟩ := ha (ε / (2 * (|M| + 1))) (by positivity)
  obtain ⟨N₂, hN₂⟩ := hb (ε / (2 * C)) (by positivity)
  use max N₁ N₂
  intro n hn
  have h1 := hN₁ n (le_of_max_le_left hn)
  have h2 := hN₂ n (le_of_max_le_right hn)
  have hM : |M| ≤ |M| + 1 := by linarith
  calc |a n * b n - L * M| = |a n * (b n - M) + M * (a n - L)| := by congr 1; ring
    _ ≤ |a n * (b n - M)| + |M * (a n - L)| := abs_add_le _ _
    _ = |a n| * |b n - M| + |M| * |a n - L| := by rw [abs_mul, abs_mul]
    _ ≤ C * |b n - M| + (|M| + 1) * |a n - L| := by gcongr; exact hCa n
    _ < C * (ε / (2 * C)) + (|M| + 1) * (ε / (2 * (|M| + 1))) := by gcongr
    _ = ε := by field_simp; ring

/-!
`gcongr`（generalized congruence）会把形如 $x_1 y_1 + x_2 y_2 \le x_1' y_1' + x_2' y_2'$
的目标拆成各个因子的比较 $x_i \le x_i'$、$y_i \le y_i'$，并自动处理非负性条件。
`field_simp` 负责通分化简分式。

## 练习

### 练习 5.1　数乘的极限
若 $a_n \to L$，则 $c\, a_n \to cL$。

提示：可以直接用 `tendsTo_mul` 和 `tendsTo_const`！
注意 `tendsTo_mul (tendsTo_const c) h` 得到的数列是 `fun n ↦ (fun _ ↦ c) n * a n`，
和目标只差一步化简，可以用 `exact` 直接给出（Lean 会自动识别），或用 `simpa using`。
-/

theorem tendsTo_const_mul {a : ℕ → ℝ} {L : ℝ} (c : ℝ) (h : TendsTo a L) :
    TendsTo (fun n ↦ c * a n) (c * L) := by
  sorry

/-!
### 练习 5.2　差的极限
若 $a_n \to L$，$b_n \to M$，则 $a_n - b_n \to L - M$。

提示：可以模仿定理 5.1 的证明，也可以利用 $a_n - b_n = a_n + (-1) b_n$。
-/

theorem tendsTo_sub {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M) :
    TendsTo (fun n ↦ a n - b n) (L - M) := by
  sorry

/-!
### 练习 5.3　夹逼定理
若 $a_n \le b_n \le c_n$，且 $a_n \to L$，$c_n \to L$，则 $b_n \to L$。

提示：用 `abs_lt` 把 `|a n - L| < ε` 等全部拆开，剩下的交给 `linarith`。
`rw [abs_lt] at h1 h2 ⊢` 可以同时改写多个假设和目标（`⊢` 表示目标）。
-/

theorem squeeze {a b c : ℕ → ℝ} {L : ℝ} (ha : TendsTo a L) (hc : TendsTo c L)
    (hab : ∀ n, a n ≤ b n) (hbc : ∀ n, b n ≤ c n) : TendsTo b L := by
  sorry

/-!
### 练习 5.4　平方的极限
若 $a_n \to L$，则 $a_n^2 \to L^2$。

提示：用 `tendsTo_mul`，然后用 `simpa [sq] using ...`（`sq x : x ^ 2 = x * x`）。
-/

theorem tendsTo_sq {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ a n ^ 2) (L ^ 2) := by
  sorry

/-!
### 练习 5.5 ★　保号性与保序性
若 $a_n \to L$，$b_n \to M$，且对所有 $n$ 有 $a_n \le b_n$，则 $L \le M$。

提示：反证。若 $M < L$，取 $\varepsilon = \frac{L - M}{2}$，在足够大的 $n$ 处
会得到 $b_n < \frac{L+M}{2} < a_n$。
-/

theorem le_of_tendsTo {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M)
    (h : ∀ n, a n ≤ b n) : L ≤ M := by
  sorry

/-!
### 练习 5.6 ★　计算一个具体的极限
证明 $\displaystyle \lim_{n \to \infty} \frac{2n + 1}{n + 1} = 2$。

提示：$\frac{2n+1}{n+1} = 2 - \frac{1}{n+1}$。先用运算法则得到
`h : TendsTo (fun n ↦ 2 - 1 / ((n : ℝ) + 1)) (2 - 0)`，然后用 `convert h using 1` 把目标
化为两个等式：数列相等（用 `funext n`，再 `field_simp`、`ring`）和极限相等（`norm_num`）。
-/

theorem ex5_6 : TendsTo (fun n ↦ (2 * (n : ℝ) + 1) / ((n : ℝ) + 1)) 2 := by
  sorry

/-!
## 本章小结

* **$\frac{\varepsilon}{2}$ 技巧**：把误差分给几个部分；
* **加一项减一项**：$a_n b_n - LM = a_n(b_n - M) + M(a_n - L)$；
* 收敛数列有界，这是证明乘积法则的关键一步；
* 夹逼定理与保序性让我们可以通过不等式来研究极限。
-/

end Ch05
