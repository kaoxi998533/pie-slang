import Analysis.Ch02_Reals

/-! # 第 2 章　练习参考答案 -/

namespace Ch02.Solutions

theorem ex2_1 (a b c : ℝ) : |a - c| ≤ |a - b| + |b - c| := by
  calc |a - c| = |(a - b) + (b - c)| := by congr 1; ring
    _ ≤ |a - b| + |b - c| := triangle (a - b) (b - c)

theorem ex2_2 (a b : ℝ) : |a| - |b| ≤ |a - b| := by
  have h : |a| ≤ |a - b| + |b| := by
    calc |a| = |(a - b) + b| := by congr 1; ring
      _ ≤ |a - b| + |b| := triangle (a - b) b
  linarith

theorem ex2_3 (x a ε : ℝ) : |x - a| < ε ↔ a - ε < x ∧ x < a + ε := by
  rw [abs_lt]
  constructor
  · intro h
    obtain ⟨h₁, h₂⟩ := h
    constructor
    · linarith
    · linarith
  · intro h
    obtain ⟨h₁, h₂⟩ := h
    constructor
    · linarith
    · linarith

theorem ex2_4 (x : ℝ) (h : ∀ ε > 0, |x| < ε) : x = 0 := by
  by_contra hx
  have hpos : 0 < |x| := abs_pos.mpr hx
  have h1 := h |x| hpos
  linarith

theorem ex2_5 (x y : ℝ) (hx : |x| < 1) : |x * y| ≤ |y| := by
  rw [abs_mul]
  nlinarith [abs_nonneg x, abs_nonneg y]

theorem ex2_6 (x : ℝ) (hx : -1 ≤ x) (n : ℕ) : 1 + n * x ≤ (1 + x) ^ n := by
  induction n with
  | zero => simp
  | succ k ih =>
    push_cast
    rw [pow_succ]
    have h1 : 0 ≤ 1 + x := by linarith
    have h2 : (1 + k * x) * (1 + x) ≤ (1 + x) ^ k * (1 + x) :=
      mul_le_mul_of_nonneg_right ih h1
    have h3 : 0 ≤ (k : ℝ) * x ^ 2 := by positivity
    nlinarith [h2, h3]

end Ch02.Solutions
