import Mathlib

/-!
# 第 9 章　导数与中值定理

导数描述函数的**瞬时变化率**。经典的定义是差商的极限：
$$f'(a) = \lim_{x \to a} \frac{f(x) - f(a)}{x - a}.$$

## 9.1 不用除法的定义

差商定义中有除法，并且要求 $x \ne a$，在形式化时有些麻烦。
我们采用一个等价的、更适合估计的形式：$f'(a)$ 是使得
$$f(x) = f(a) + f'(a)(x - a) + \text{（比 } |x - a| \text{ 更小的误差）}$$
成立的那个数。也就是说，$f$ 在 $a$ 附近可以用直线 $f(a) + f'(a)(x-a)$ 来**线性逼近**。

> **定义 9.1**　称 $f$ 在 $a$ 处可导且导数为 $f'$，如果
> $$\forall \varepsilon > 0,\ \exists \delta > 0,\ \forall x,\ |x - a| < \delta \Rightarrow
> |f(x) - f(a) - f'\cdot(x - a)| \le \varepsilon |x - a|.$$

对 $x \ne a$，两边除以 $|x - a|$，就得到 $\left|\frac{f(x) - f(a)}{x - a} - f'\right| \le \varepsilon$，
这正是差商极限的定义。
-/

namespace Ch09

/-- `f` 在 `a` 处可导，导数为 `f'`。 -/
def HasDeriv (f : ℝ → ℝ) (f' a : ℝ) : Prop :=
  ∀ ε > 0, ∃ δ > 0, ∀ x, |x - a| < δ → |f x - f a - f' * (x - a)| ≤ ε * |x - a|

/-!
## 9.2 例子：$(x^2)' = 2x$

> **例 9.2**　$f(x) = x^2$ 在 $a$ 处的导数为 $2a$。

**草稿**：误差项为
$$x^2 - a^2 - 2a(x - a) = (x - a)^2,$$
所以 $|{\text{误差}}| = |x - a| \cdot |x - a|$。只要 $|x - a| < \varepsilon$，
误差就 $\le \varepsilon |x - a|$。取 $\delta = \varepsilon$。
-/

example (a : ℝ) : HasDeriv (fun x ↦ x ^ 2) (2 * a) a := by
  intro ε hε
  refine ⟨ε, hε, fun x hx ↦ ?_⟩
  show |x ^ 2 - a ^ 2 - 2 * a * (x - a)| ≤ ε * |x - a|
  have h : x ^ 2 - a ^ 2 - 2 * a * (x - a) = (x - a) * (x - a) := by ring
  rw [h, abs_mul]
  exact mul_le_mul_of_nonneg_right hx.le (abs_nonneg _)

/-!
## 9.3 可导必连续

为了陈述"可导必连续"，我们再次引入连续的定义。
-/

/-- 函数 `f` 在点 `a` 处连续。 -/
def ContAt (f : ℝ → ℝ) (a : ℝ) : Prop :=
  ∀ ε > 0, ∃ δ > 0, ∀ x, |x - a| < δ → |f x - f a| < ε

/-!
## 练习（一）：用定义计算导数

### 练习 9.1　常数的导数为零
提示：取 $\delta = 1$。`simp` 可以化简误差项，剩下的目标可以用 `positivity`。
-/

theorem hasDeriv_const (c a : ℝ) : HasDeriv (fun _ ↦ c) 0 a := by
  sorry

/-!
### 练习 9.2　线性函数的导数
$(mx + b)' = m$。误差项恒等于零！
-/

theorem hasDeriv_linear (m b a : ℝ) : HasDeriv (fun x ↦ m * x + b) m a := by
  sorry

/-!
### 练习 9.3 ★　和的求导法则
$(f + g)' = f' + g'$。

提示：$\frac{\varepsilon}{2}$ 技巧与 $\delta = \min(\delta_1, \delta_2)$。
误差项可以拆成 $f$ 的误差加上 $g$ 的误差。
-/

theorem hasDeriv_add {f g : ℝ → ℝ} {f' g' a : ℝ} (hf : HasDeriv f f' a) (hg : HasDeriv g g' a) :
    HasDeriv (fun x ↦ f x + g x) (f' + g') a := by
  sorry

/-!
### 练习 9.4 ★★　可导必连续
**草稿**：取定义中 $\varepsilon = 1$ 得到 $\delta_0$。当 $|x - a| < \delta_0$ 时
$$|f(x) - f(a)| \le |f(x) - f(a) - f'(x - a)| + |f'||x - a| \le (1 + |f'|)|x - a|.$$
因此取 $\delta = \min\left(\delta_0, \frac{\varepsilon}{1 + |f'|}\right)$。

提示：最后一步需要 $(1 + |f'|) \cdot |x - a| < (1 + |f'|) \cdot \frac{\varepsilon}{1 + |f'|} = \varepsilon$，
`mul_lt_mul_of_pos_left` 和 `field_simp` 会有帮助。
-/

theorem contAt_of_hasDeriv {f : ℝ → ℝ} {f' a : ℝ} (hf : HasDeriv f f' a) : ContAt f a := by
  sorry

/-!
## 9.4 Mathlib 中的导数

Mathlib 中 `HasDerivAt f f' a` 表示 $f$ 在 $a$ 处的导数为 $f'$，
`deriv f a` 是导数的值，`Differentiable ℝ f` 表示 $f$ 处处可导。
Mathlib 提供了完整的求导法则，例如 `HasDerivAt.add`、`HasDerivAt.mul`、`hasDerivAt_pow`。

组合使用这些法则，就可以求出具体函数的导数。`hasDerivAt_pow 2 x` 给出
$(x^2)' = 2x^{2-1}$，`simpa` 把它化简成 $2x$；`HasDerivAt.const_mul` 处理常数倍；
`h1.add h2` 是和的求导法则；最后 `.deriv` 从 `HasDerivAt` 得到 `deriv` 的值。
-/

example (x : ℝ) : deriv (fun x : ℝ ↦ x ^ 2 + 3 * x) x = 2 * x + 3 := by
  have h1 : HasDerivAt (fun x : ℝ ↦ x ^ 2) (2 * x) x := by
    simpa using hasDerivAt_pow 2 x
  have h2 : HasDerivAt (fun x : ℝ ↦ 3 * x) 3 x := by
    simpa using (hasDerivAt_id x).const_mul 3
  exact (h1.add h2).deriv

/-!
## 9.5 微分中值定理

> **定理 9.3（拉格朗日中值定理）**　若 $f$ 在 $[a, b]$ 上连续、在 $(a, b)$ 内可导，
> 则存在 $c \in (a, b)$ 使得
> $$f'(c) = \frac{f(b) - f(a)}{b - a}.$$

几何意义：曲线上总有一点的切线平行于两端点的连线。Mathlib 中的陈述为

```text
exists_deriv_eq_slope (f : ℝ → ℝ) (hab : a < b) (hfc : ContinuousOn f (Icc a b))
    (hfd : DifferentiableOn ℝ f (Ioo a b)) : ∃ c ∈ Ioo a b, deriv f c = (f b - f a) / (b - a)
```

中值定理是联系**导数**（局部信息）与**函数值**（整体信息）的桥梁。

> **例 9.4**　若 $f' \equiv 0$，则 $f$ 是常数。

**证明**：对任意 $x < y$，由中值定理存在 $c$ 使 $0 = f'(c) = \frac{f(y) - f(x)}{y - x}$，
所以 $f(y) = f(x)$。
-/

theorem const_of_deriv_zero {f : ℝ → ℝ} (hf : Differentiable ℝ f) (h : ∀ x, deriv f x = 0)
    {x y : ℝ} (hxy : x < y) : f x = f y := by
  obtain ⟨c, _, hc⟩ := exists_deriv_eq_slope f hxy hf.continuous.continuousOn
    hf.differentiableOn
  rw [h c, eq_comm, div_eq_zero_iff] at hc
  rcases hc with hc | hc
  · linarith
  · linarith

/-!
`div_eq_zero_iff : a / b = 0 ↔ a = 0 ∨ b = 0`。第二种情形 $y - x = 0$ 与 $x < y$ 矛盾。

## 练习（二）：中值定理的应用

### 练习 9.5 ★　导数非负则函数递增
若 $f$ 处处可导且 $f' \ge 0$，则 $a < b$ 时 $f(a) \le f(b)$。

提示：模仿例 9.4。得到 `hc : 0 ≤ (f b - f a) / (b - a)` 后，
`div_nonneg_iff` 或 `le_div_iff₀ (hc : 0 < c) : a ≤ b / c ↔ a * c ≤ b` 可以去分母。
-/

theorem ex9_5 {f : ℝ → ℝ} (hf : Differentiable ℝ f) (h : ∀ x, 0 ≤ deriv f x)
    {a b : ℝ} (hab : a < b) : f a ≤ f b := by
  sorry

/-!
### 练习 9.6 ★★　导数有界则函数利普希茨
若 $|f'| \le M$，则 $a < b$ 时 $|f(b) - f(a)| \le M (b - a)$。

提示：由中值定理，$\frac{f(b) - f(a)}{b - a} = f'(c)$，于是
$f(b) - f(a) = f'(c) (b - a)$（`div_eq_iff`），再用 `abs_mul`。
-/

theorem ex9_6 {f : ℝ → ℝ} {M : ℝ} (hf : Differentiable ℝ f) (h : ∀ x, |deriv f x| ≤ M)
    {a b : ℝ} (hab : a < b) : |f b - f a| ≤ M * (b - a) := by
  sorry

/-!
## 本章小结

* 导数的线性逼近定义：误差 $\le \varepsilon |x - a|$；
* 可导必连续，求导法则可以用 $\varepsilon$-$\delta$ 严格证明；
* 拉格朗日中值定理 `exists_deriv_eq_slope` 把导数的信息转化为函数值的信息。

## 全书结语

恭喜你完成了本书！你已经亲手用 Lean 严格证明了数学分析入门的核心定理：
极限的唯一性与运算法则、夹逼定理、单调收敛定理、连续函数的运算、
介值定理与最值定理的应用、可导必连续以及中值定理的应用。

接下来可以继续探索：

* [Mathematics in Lean](https://leanprover-community.github.io/mathematics_in_lean/)：Lean 官方的数学教程；
* [Mathlib 文档](https://leanprover-community.github.io/mathlib4_docs/)：查找定理；
* 尝试用本书的方法形式化积分、级数与一致收敛。
-/

end Ch09
