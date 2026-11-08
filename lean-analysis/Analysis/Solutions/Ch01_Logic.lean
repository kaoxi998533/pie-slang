import Analysis.Ch01_Logic

/-! # 第 1 章　练习参考答案 -/

namespace Ch01.Solutions

theorem ex1_1 (P Q R : Prop) (hPQ : P → Q) (hQR : Q → R) : P → R := by
  intro hP
  apply hQR
  apply hPQ
  exact hP

theorem ex1_2 (P Q : Prop) : P ∧ Q → Q ∧ P := by
  intro h
  obtain ⟨hP, hQ⟩ := h
  constructor
  · exact hQ
  · exact hP

theorem ex1_3 (P Q : Prop) : P ∨ Q → Q ∨ P := by
  intro h
  rcases h with hP | hQ
  · right
    exact hP
  · left
    exact hQ

theorem ex1_4 : ∀ x : ℝ, ∃ y : ℝ, x < y := by
  intro x
  use x + 1
  linarith

theorem ex1_5 (P : ℝ → ℝ → Prop) (h : ∃ x, ∀ y, P x y) : ∀ y, ∃ x, P x y := by
  intro y
  obtain ⟨x, hx⟩ := h
  use x
  exact hx y

theorem ex1_6 {f : ℝ → ℝ} {c : ℝ} (hc : 0 ≤ c) (hf : BddAboveFn f) :
    BddAboveFn (fun x ↦ c * f x) := by
  obtain ⟨M, hM⟩ := hf
  use c * M
  intro x
  show c * f x ≤ c * M
  exact mul_le_mul_of_nonneg_left (hM x) hc

theorem ex1_7 (f : ℝ → ℝ) : ¬ BddAboveFn f ↔ ∀ M, ∃ x, M < f x := by
  unfold BddAboveFn
  push Not
  rfl

theorem ex1_8 : ¬ BddAboveFn (fun x ↦ x ^ 2) := by
  intro h
  obtain ⟨M, hM⟩ := h
  have h1 : (|M| + 1) ^ 2 ≤ M := hM (|M| + 1)
  have h2 : M ≤ |M| := le_abs_self M
  nlinarith [abs_nonneg M]

end Ch01.Solutions
