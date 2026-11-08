import Analysis.Ch00_Welcome

/-! # 第 0 章　练习参考答案 -/

namespace Ch00.Solutions

theorem ex0_1 (x : ℝ) : (x + 1) ^ 2 - 1 = x * (x + 2) := by
  ring

theorem ex0_2 (a b : ℝ) (h₁ : a = 2 * b) (h₂ : b = 3) : a = 6 := by
  rw [h₁, h₂]
  norm_num

theorem ex0_3 (x y : ℝ) (h₁ : 2 * x + y = 5) (h₂ : x - y = 1) : x = 2 := by
  linarith

theorem ex0_4 (ε : ℝ) : ε / 2 + ε / 2 = ε := by
  ring

theorem ex0_5 (x y : ℝ) (hx : x < 1) (hy : y < 2) : x + y < 4 := by
  calc x + y < 1 + 2 := by linarith
    _ < 4 := by norm_num

theorem ex0_6 (a b : ℝ) : 2 * a * b ≤ a ^ 2 + b ^ 2 := by
  have h : 0 ≤ (a - b) ^ 2 := sq_nonneg (a - b)
  nlinarith [h]

end Ch00.Solutions
