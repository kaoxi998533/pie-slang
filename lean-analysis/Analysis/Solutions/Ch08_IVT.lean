import Analysis.Ch08_IVT

/-! # 第 8 章　练习参考答案 -/

namespace Ch08.Solutions

open Set

theorem ex8_1 : ∃ x ∈ Icc (0 : ℝ) 1, x ^ 3 + x - 1 = 0 := by
  have hcont : ContinuousOn (fun x : ℝ ↦ x ^ 3 + x - 1) (Icc 0 1) := by fun_prop
  have hivt := intermediate_value_Icc (by norm_num : (0 : ℝ) ≤ 1) hcont
  have h0 : (0 : ℝ) ∈ Icc ((0 : ℝ) ^ 3 + 0 - 1) ((1 : ℝ) ^ 3 + 1 - 1) := by
    rw [mem_Icc]
    norm_num
  obtain ⟨x, hx, hx0⟩ := hivt h0
  exact ⟨x, hx, hx0⟩

theorem exists_min {f : ℝ → ℝ} {a b : ℝ} (hab : a ≤ b) (hf : ContinuousOn f (Icc a b)) :
    ∃ c ∈ Icc a b, ∀ x ∈ Icc a b, f c ≤ f x := by
  obtain ⟨c, hc, hmin⟩ := isCompact_Icc.exists_isMinOn (nonempty_Icc.mpr hab) hf
  exact ⟨c, hc, isMinOn_iff.mp hmin⟩

theorem ex8_3 {f : ℝ → ℝ} {a b : ℝ} (hab : a ≤ b) (hf : ContinuousOn f (Icc a b)) :
    ∃ M, ∀ x ∈ Icc a b, |f x| ≤ M := by
  obtain ⟨c₁, _, hc₁⟩ := exists_max hab hf
  obtain ⟨c₂, _, hc₂⟩ := exists_min hab hf
  use |f c₁| + |f c₂|
  intro x hx
  have h1 := hc₁ x hx
  have h2 := hc₂ x hx
  rw [abs_le]
  constructor
  · linarith [neg_abs_le (f c₂), abs_nonneg (f c₁)]
  · linarith [le_abs_self (f c₁), abs_nonneg (f c₂)]

theorem ex8_4 {f : ℝ → ℝ} (hf : ContinuousOn f (Icc 0 1))
    (hmaps : ∀ x ∈ Icc (0 : ℝ) 1, f x ∈ Icc (0 : ℝ) 1) :
    ∃ x ∈ Icc (0 : ℝ) 1, f x = x := by
  have hg : ContinuousOn (fun x ↦ f x - x) (Icc 0 1) := hf.sub continuousOn_id
  have hivt := intermediate_value_Icc' (by norm_num : (0 : ℝ) ≤ 1) hg
  have h0 := hmaps 0 (by rw [mem_Icc]; norm_num)
  have h1 := hmaps 1 (by rw [mem_Icc]; norm_num)
  rw [mem_Icc] at h0 h1
  have hmem : (0 : ℝ) ∈ Icc (f 1 - 1) (f 0 - 0) := by
    rw [mem_Icc]
    constructor
    · linarith
    · linarith
  obtain ⟨x, hx, hx0⟩ := hivt hmem
  refine ⟨x, hx, ?_⟩
  have : f x - x = 0 := hx0
  linarith

theorem ex8_5 : ¬ ∃ c ∈ Ioo (0 : ℝ) 1, ∀ x ∈ Ioo (0 : ℝ) 1, x ≤ c := by
  rintro ⟨c, hc, hmax⟩
  rw [mem_Ioo] at hc
  have hmem : (c + 1) / 2 ∈ Ioo (0 : ℝ) 1 := by
    rw [mem_Ioo]
    constructor
    · linarith
    · linarith
  have h1 := hmax ((c + 1) / 2) hmem
  linarith

end Ch08.Solutions
