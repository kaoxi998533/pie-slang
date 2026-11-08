import Analysis.Ch06_Monotone

/-! # 第 6 章　练习参考答案 -/

namespace Ch06.Solutions

theorem ex6_1 : Increasing (fun n ↦ 1 - 1 / ((n : ℝ) + 1)) := by
  intro n
  show 1 - 1 / ((n : ℝ) + 1) ≤ 1 - 1 / (((n + 1 : ℕ) : ℝ) + 1)
  push_cast
  gcongr
  linarith

theorem ex6_2 {a : ℕ → ℝ} (h : ∀ n, a (n + 1) ≤ a n) (hb : ∃ m, ∀ n, m ≤ a n) :
    ∃ L, TendsTo a L := by
  obtain ⟨m, hm⟩ := hb
  have hinc : Increasing (fun n ↦ -a n) := by
    intro n
    show -a n ≤ -a (n + 1)
    linarith [h n]
  have hbdd : ∃ M, ∀ n, (fun n ↦ -a n) n ≤ M := by
    use -m
    intro n
    show -a n ≤ -m
    linarith [hm n]
  obtain ⟨L, hL⟩ := monotone_convergence hinc hbdd
  use -L
  simpa using tendsTo_neg hL

theorem ex6_3 {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) : IsCauchy a := by
  intro ε hε
  obtain ⟨N, hN⟩ := h (ε / 2) (by positivity)
  use N
  intro m hm n hn
  have h1 := hN m hm
  have h2 := hN n hn
  calc |a m - a n| = |(a m - L) + (L - a n)| := by congr 1; ring
    _ ≤ |a m - L| + |L - a n| := abs_add_le _ _
    _ = |a m - L| + |a n - L| := by rw [abs_sub_comm L]
    _ < ε / 2 + ε / 2 := by linarith
    _ = ε := by ring

theorem ex6_4 {a : ℕ → ℝ} (h : IsCauchy a) : ∃ C, ∀ n, |a n| ≤ C := by
  obtain ⟨N, hN⟩ := h 1 one_pos
  use (∑ i ∈ Finset.range N, |a i|) + |a N| + 1
  intro n
  have hsum : 0 ≤ ∑ i ∈ Finset.range N, |a i| :=
    Finset.sum_nonneg (fun i _ ↦ abs_nonneg (a i))
  rcases lt_or_ge n N with hn | hn
  · have : |a n| ≤ ∑ i ∈ Finset.range N, |a i| :=
      Finset.single_le_sum (f := fun i ↦ |a i|) (fun i _ ↦ abs_nonneg (a i))
        (Finset.mem_range.mpr hn)
    linarith [abs_nonneg (a N)]
  · have h1 := hN n hn N le_rfl
    have h2 : |a n| ≤ |a n - a N| + |a N| := by
      calc |a n| = |(a n - a N) + a N| := by congr 1; ring
        _ ≤ |a n - a N| + |a N| := abs_add_le _ _
    linarith

theorem ex6_5 {a : ℕ → ℝ} {L : ℝ} (h : Increasing a) (hL : TendsTo a L) :
    ∀ n, a n ≤ L := by
  intro k
  by_contra hk
  push Not at hk
  obtain ⟨N, hN⟩ := hL (a k - L) (by linarith)
  have h1 := hN (max N k) (le_max_left _ _)
  have h2 : a k ≤ a (max N k) := h.mono (le_max_right _ _)
  rw [abs_lt] at h1
  linarith [h1.2]

end Ch06.Solutions
