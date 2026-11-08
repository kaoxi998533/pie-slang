import Analysis.Ch03_Completeness

/-! # 第 3 章　练习参考答案 -/

namespace Ch03.Solutions

theorem ex3_1 {S : Set ℝ} {s t : ℝ} (h₁ : IsSup S s) (h₂ : IsSup S t) : s = t := by
  apply le_antisymm
  · exact h₁.2 t h₂.1
  · exact h₂.2 s h₁.1

theorem ex3_2 {S : Set ℝ} {s : ℝ} (hs : s ∈ S) (hub : IsUB S s) : IsSup S s := by
  constructor
  · exact hub
  · intro M hM
    exact hM s hs

theorem ex3_3 : IsSup {x : ℝ | x < 1} 1 := by
  constructor
  · intro x hx
    have hx' : x < 1 := hx
    linarith
  · intro M hM
    by_contra h
    push Not at h
    -- h : M < 1，考虑 (M + 1) / 2
    have hmem : (M + 1) / 2 ∈ {x : ℝ | x < 1} := by
      show (M + 1) / 2 < 1
      linarith
    have h1 := hM ((M + 1) / 2) hmem
    linarith

theorem ex3_4 (ε : ℝ) (hε : 0 < ε) : ∃ n : ℕ, 1 / ((n : ℝ) + 1) < ε := by
  obtain ⟨n, hn⟩ := archimedean (1 / ε)
  use n
  have hpos : (0 : ℝ) < n + 1 := by positivity
  rw [div_lt_iff₀ hpos]
  rw [div_lt_iff₀ hε] at hn
  linarith

theorem ex3_5 {S : Set ℝ} {s : ℝ} (h : IsSup S s) : IsInf {y | -y ∈ S} (-s) := by
  constructor
  · intro y hy
    have h1 : -y ≤ s := h.1 (-y) hy
    linarith
  · intro m hm
    -- -m 是 S 的上界
    have hub : IsUB S (-m) := by
      intro x hx
      have hx' : -(-x) ∈ S := by simpa using hx
      have h1 : m ≤ -x := hm (-x) hx'
      linarith
    have h2 : s ≤ -m := h.2 (-m) hub
    linarith

end Ch03.Solutions
