import Analysis.Ch09_Derivative

/-! # 第 9 章　练习参考答案 -/

namespace Ch09.Solutions

theorem hasDeriv_const (c a : ℝ) : HasDeriv (fun _ ↦ c) 0 a := by
  intro ε hε
  refine ⟨1, one_pos, fun x _ ↦ ?_⟩
  simp
  positivity

theorem hasDeriv_linear (m b a : ℝ) : HasDeriv (fun x ↦ m * x + b) m a := by
  intro ε hε
  refine ⟨1, one_pos, fun x _ ↦ ?_⟩
  show |m * x + b - (m * a + b) - m * (x - a)| ≤ ε * |x - a|
  rw [show m * x + b - (m * a + b) - m * (x - a) = 0 by ring, abs_zero]
  positivity

theorem hasDeriv_add {f g : ℝ → ℝ} {f' g' a : ℝ} (hf : HasDeriv f f' a) (hg : HasDeriv g g' a) :
    HasDeriv (fun x ↦ f x + g x) (f' + g') a := by
  intro ε hε
  obtain ⟨δ₁, hδ₁, h₁⟩ := hf (ε / 2) (by positivity)
  obtain ⟨δ₂, hδ₂, h₂⟩ := hg (ε / 2) (by positivity)
  refine ⟨min δ₁ δ₂, lt_min hδ₁ hδ₂, fun x hx ↦ ?_⟩
  have hx₁ := h₁ x (lt_of_lt_of_le hx (min_le_left _ _))
  have hx₂ := h₂ x (lt_of_lt_of_le hx (min_le_right _ _))
  calc |f x + g x - (f a + g a) - (f' + g') * (x - a)|
      = |(f x - f a - f' * (x - a)) + (g x - g a - g' * (x - a))| := by congr 1; ring
    _ ≤ |f x - f a - f' * (x - a)| + |g x - g a - g' * (x - a)| := abs_add_le _ _
    _ ≤ ε / 2 * |x - a| + ε / 2 * |x - a| := add_le_add hx₁ hx₂
    _ = ε * |x - a| := by ring

theorem contAt_of_hasDeriv {f : ℝ → ℝ} {f' a : ℝ} (hf : HasDeriv f f' a) : ContAt f a := by
  intro ε hε
  obtain ⟨δ₀, hδ₀, h₀⟩ := hf 1 one_pos
  have hK : 0 < 1 + |f'| := by positivity
  refine ⟨min δ₀ (ε / (1 + |f'|)), lt_min hδ₀ (by positivity), fun x hx ↦ ?_⟩
  have hx₀ : |x - a| < δ₀ := lt_of_lt_of_le hx (min_le_left _ _)
  have hx₁ : |x - a| < ε / (1 + |f'|) := lt_of_lt_of_le hx (min_le_right _ _)
  have h1 := h₀ x hx₀
  calc |f x - f a| = |(f x - f a - f' * (x - a)) + f' * (x - a)| := by congr 1; ring
    _ ≤ |f x - f a - f' * (x - a)| + |f' * (x - a)| := abs_add_le _ _
    _ ≤ 1 * |x - a| + |f'| * |x - a| := by rw [abs_mul]; linarith
    _ = (1 + |f'|) * |x - a| := by ring
    _ < (1 + |f'|) * (ε / (1 + |f'|)) := mul_lt_mul_of_pos_left hx₁ hK
    _ = ε := by field_simp

theorem ex9_5 {f : ℝ → ℝ} (hf : Differentiable ℝ f) (h : ∀ x, 0 ≤ deriv f x)
    {a b : ℝ} (hab : a < b) : f a ≤ f b := by
  obtain ⟨c, _, hc⟩ := exists_deriv_eq_slope f hab hf.continuous.continuousOn
    hf.differentiableOn
  have h1 := h c
  rw [hc, le_div_iff₀ (by linarith)] at h1
  linarith

theorem ex9_6 {f : ℝ → ℝ} {M : ℝ} (hf : Differentiable ℝ f) (h : ∀ x, |deriv f x| ≤ M)
    {a b : ℝ} (hab : a < b) : |f b - f a| ≤ M * (b - a) := by
  obtain ⟨c, _, hc⟩ := exists_deriv_eq_slope f hab hf.continuous.continuousOn
    hf.differentiableOn
  have hba : 0 < b - a := by linarith
  have h1 : f b - f a = deriv f c * (b - a) := by
    rw [hc]
    field_simp
  rw [h1, abs_mul, abs_of_pos hba]
  exact mul_le_mul_of_nonneg_right (h c) hba.le

end Ch09.Solutions
