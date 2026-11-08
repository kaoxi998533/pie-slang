import Analysis.Ch07_Continuity

/-! # 第 7 章　练习参考答案 -/

namespace Ch07.Solutions

theorem contAt_const (c a : ℝ) : ContAt (fun _ ↦ c) a := by
  intro ε hε
  use 1, one_pos
  intro x _
  simpa using hε

theorem contAt_id (a : ℝ) : ContAt (fun x ↦ x) a := by
  intro ε hε
  use ε, hε
  intro x hx
  exact hx

theorem contAt_add {f g : ℝ → ℝ} {a : ℝ} (hf : ContAt f a) (hg : ContAt g a) :
    ContAt (fun x ↦ f x + g x) a := by
  intro ε hε
  obtain ⟨δ₁, hδ₁, h₁⟩ := hf (ε / 2) (by positivity)
  obtain ⟨δ₂, hδ₂, h₂⟩ := hg (ε / 2) (by positivity)
  refine ⟨min δ₁ δ₂, lt_min hδ₁ hδ₂, fun x hx ↦ ?_⟩
  have hx₁ := h₁ x (lt_of_lt_of_le hx (min_le_left _ _))
  have hx₂ := h₂ x (lt_of_lt_of_le hx (min_le_right _ _))
  calc |f x + g x - (f a + g a)| = |(f x - f a) + (g x - g a)| := by congr 1; ring
    _ ≤ |f x - f a| + |g x - g a| := abs_add_le _ _
    _ < ε / 2 + ε / 2 := by linarith
    _ = ε := by ring

theorem contAt_comp {f g : ℝ → ℝ} {a : ℝ} (hf : ContAt f a) (hg : ContAt g (f a)) :
    ContAt (fun x ↦ g (f x)) a := by
  intro ε hε
  obtain ⟨η, hη, hg'⟩ := hg ε hε
  obtain ⟨δ, hδ, hf'⟩ := hf η hη
  refine ⟨δ, hδ, fun x hx ↦ ?_⟩
  exact hg' (f x) (hf' x hx)

theorem contAt_sq (a : ℝ) : ContAt (fun x ↦ x ^ 2) a := by
  intro ε hε
  have hpos : 0 < 2 * |a| + 1 := by positivity
  refine ⟨min 1 (ε / (2 * |a| + 1)), lt_min one_pos (by positivity), fun x hx ↦ ?_⟩
  have h1 : |x - a| < 1 := lt_of_lt_of_le hx (min_le_left _ _)
  have h2 : |x - a| < ε / (2 * |a| + 1) := lt_of_lt_of_le hx (min_le_right _ _)
  have h3 : |x + a| ≤ 2 * |a| + 1 := by
    calc |x + a| = |(x - a) + 2 * a| := by congr 1; ring
      _ ≤ |x - a| + |2 * a| := abs_add_le _ _
      _ = |x - a| + 2 * |a| := by rw [abs_mul]; norm_num
      _ ≤ 2 * |a| + 1 := by linarith
  calc |x ^ 2 - a ^ 2| = |x - a| * |x + a| := by rw [← abs_mul]; congr 1; ring
    _ ≤ |x - a| * (2 * |a| + 1) := by gcongr
    _ < ε / (2 * |a| + 1) * (2 * |a| + 1) := by gcongr
    _ = ε := by field_simp

theorem contAt_seq {f : ℝ → ℝ} {a : ℝ} {x : ℕ → ℝ} (hf : ContAt f a) (hx : TendsTo x a) :
    TendsTo (fun n ↦ f (x n)) (f a) := by
  intro ε hε
  obtain ⟨δ, hδ, hf'⟩ := hf ε hε
  obtain ⟨N, hN⟩ := hx δ hδ
  use N
  intro n hn
  exact hf' (x n) (hN n hn)

theorem contAt_const_mul {f : ℝ → ℝ} {a : ℝ} (c : ℝ) (hf : ContAt f a) :
    ContAt (fun x ↦ c * f x) a := by
  intro ε hε
  have hc : 0 < |c| + 1 := by positivity
  obtain ⟨δ, hδ, hf'⟩ := hf (ε / (|c| + 1)) (by positivity)
  refine ⟨δ, hδ, fun x hx ↦ ?_⟩
  have h1 := hf' x hx
  calc |c * f x - c * f a| = |c| * |f x - f a| := by rw [← mul_sub, abs_mul]
    _ ≤ (|c| + 1) * |f x - f a| := by gcongr; linarith
    _ < (|c| + 1) * (ε / (|c| + 1)) := by gcongr
    _ = ε := by field_simp

end Ch07.Solutions
