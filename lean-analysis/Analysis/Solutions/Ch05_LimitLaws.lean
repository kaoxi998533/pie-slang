import Analysis.Ch05_LimitLaws

/-! # 第 5 章　练习参考答案 -/

namespace Ch05.Solutions

theorem tendsTo_const_mul {a : ℕ → ℝ} {L : ℝ} (c : ℝ) (h : TendsTo a L) :
    TendsTo (fun n ↦ c * a n) (c * L) := by
  exact tendsTo_mul (tendsTo_const c) h

theorem tendsTo_sub {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M) :
    TendsTo (fun n ↦ a n - b n) (L - M) := by
  have h := tendsTo_add ha (tendsTo_const_mul (-1) hb)
  simpa [sub_eq_add_neg] using h

theorem squeeze {a b c : ℕ → ℝ} {L : ℝ} (ha : TendsTo a L) (hc : TendsTo c L)
    (hab : ∀ n, a n ≤ b n) (hbc : ∀ n, b n ≤ c n) : TendsTo b L := by
  intro ε hε
  obtain ⟨N₁, hN₁⟩ := ha ε hε
  obtain ⟨N₂, hN₂⟩ := hc ε hε
  use max N₁ N₂
  intro n hn
  have h1 := hN₁ n (le_of_max_le_left hn)
  have h2 := hN₂ n (le_of_max_le_right hn)
  rw [abs_lt] at h1 h2 ⊢
  have h3 := hab n
  have h4 := hbc n
  constructor
  · linarith
  · linarith

theorem tendsTo_sq {a : ℕ → ℝ} {L : ℝ} (h : TendsTo a L) :
    TendsTo (fun n ↦ a n ^ 2) (L ^ 2) := by
  simpa [sq] using tendsTo_mul h h

theorem le_of_tendsTo {a b : ℕ → ℝ} {L M : ℝ} (ha : TendsTo a L) (hb : TendsTo b M)
    (h : ∀ n, a n ≤ b n) : L ≤ M := by
  by_contra hLM
  push Not at hLM
  obtain ⟨N₁, hN₁⟩ := ha ((L - M) / 2) (by linarith)
  obtain ⟨N₂, hN₂⟩ := hb ((L - M) / 2) (by linarith)
  have h1 := hN₁ (max N₁ N₂) (le_max_left _ _)
  have h2 := hN₂ (max N₁ N₂) (le_max_right _ _)
  rw [abs_lt] at h1 h2
  have h3 := h (max N₁ N₂)
  linarith [h1.1, h2.2]

theorem ex5_6 : TendsTo (fun n ↦ (2 * (n : ℝ) + 1) / ((n : ℝ) + 1)) 2 := by
  have h := tendsTo_sub (tendsTo_const 2) tendsTo_one_div
  convert h using 1
  · funext n
    have hpos : (n : ℝ) + 1 ≠ 0 := by positivity
    field_simp
    ring
  · norm_num

end Ch05.Solutions
