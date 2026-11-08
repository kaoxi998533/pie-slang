import Analysis.Ch04_SeqLimit

/-! # 第 4 章　练习参考答案 -/

namespace Ch04.Solutions

theorem tendsTo_const (c : ℝ) : TendsTo (fun _ ↦ c) c := by
  intro ε hε
  use 0
  intro n _
  simpa using hε

theorem tendsTo_add_const {a : ℕ → ℝ} {L : ℝ} (c : ℝ) (h : TendsTo a L) :
    TendsTo (fun n ↦ a n + c) (L + c) := by
  intro ε hε
  obtain ⟨N, hN⟩ := h ε hε
  use N
  intro n hn
  show |a n + c - (L + c)| < ε
  rw [show a n + c - (L + c) = a n - L by ring]
  exact hN n hn

theorem tendsTo_neg {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ -a n) (-L) := by
  intro ε hε
  obtain ⟨N, hN⟩ := h ε hε
  use N
  intro n hn
  show |-a n - -L| < ε
  rw [show -a n - -L = -(a n - L) by ring, abs_neg]
  exact hN n hn

theorem tendsTo_shift {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) (k : ℕ) :
    TendsTo (fun n ↦ a (n + k)) L := by
  intro ε hε
  obtain ⟨N, hN⟩ := h ε hε
  use N
  intro n hn
  exact hN (n + k) (by omega)

theorem ex4_5 : TendsTo (fun n ↦ 1 / ((n : ℝ) + 1) + 3) 3 := by
  have h := tendsTo_add_const 3 tendsTo_one_div
  simpa using h

theorem tendsTo_iff_sub {a : ℕ → ℝ} {L : ℝ} :
    TendsTo a L ↔ TendsTo (fun n ↦ a n - L) 0 := by
  constructor
  · intro h ε hε
    obtain ⟨N, hN⟩ := h ε hε
    use N
    intro n hn
    simpa using hN n hn
  · intro h ε hε
    obtain ⟨N, hN⟩ := h ε hε
    use N
    intro n hn
    simpa using hN n hn

theorem tendsTo_abs {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ |a n|) |L| := by
  intro ε hε
  obtain ⟨N, hN⟩ := h ε hε
  use N
  intro n hn
  show |(|a n| - |L|)| < ε
  exact lt_of_le_of_lt (abs_abs_sub_abs_le_abs_sub (a n) L) (hN n hn)

theorem ex4_8 : ¬ TendsTo (fun n ↦ if n % 2 = 0 then 0 else 1) 0 := by
  intro h
  obtain ⟨N, hN⟩ := h (1 / 2) (by norm_num)
  have h1 := hN (2 * N + 1) (by omega)
  have h2 : (2 * N + 1) % 2 = 1 := by omega
  simp only [h2] at h1
  norm_num at h1

end Ch04.Solutions
