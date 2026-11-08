# Lean 数学分析入门

一本**可以运行**的数学分析入门教科书。书中每一个定义、每一条定理、每一道练习都是
[Lean 4](https://lean-lang.org) 代码，由计算机逐行检查。读者通过把练习中的 `sorry`
替换成自己的证明来学习：Lean 会实时显示"还需要证明什么"，并在证明出错时立即指出。

内容对应一门标准的"数学分析（一）"课程，从逻辑与实数出发，一直讲到介值定理和拉格朗日中值定理。

## 目录

| 章 | 文件 | 内容 | 练习 |
|----|------|------|-----:|
| 0 | [`Ch00_Welcome`](Analysis/Ch00_Welcome.lean) | 如何使用本书；`norm_num`、`ring`、`linarith`、`calc` | 6 |
| 1 | [`Ch01_Logic`](Analysis/Ch01_Logic.lean) | 命题与量词：`intro`、`obtain`、`use`、`push Not` | 8 |
| 2 | [`Ch02_Reals`](Analysis/Ch02_Reals.lean) | 实数的不等式、绝对值、三角不等式、伯努利不等式 | 6 |
| 3 | [`Ch03_Completeness`](Analysis/Ch03_Completeness.lean) | 确界原理、上确界的 ε 刻画、阿基米德性质 | 5 |
| 4 | [`Ch04_SeqLimit`](Analysis/Ch04_SeqLimit.lean) | 数列极限的 ε-N 定义、极限的唯一性 | 8 |
| 5 | [`Ch05_LimitLaws`](Analysis/Ch05_LimitLaws.lean) | 极限的四则运算、收敛数列有界、夹逼定理、保序性 | 6 |
| 6 | [`Ch06_Monotone`](Analysis/Ch06_Monotone.lean) | 单调收敛定理、柯西数列与柯西收敛准则 | 5 |
| 7 | [`Ch07_Continuity`](Analysis/Ch07_Continuity.lean) | ε-δ 连续性、连续函数的运算、海涅定理 | 7 |
| 8 | [`Ch08_IVT`](Analysis/Ch08_IVT.lean) | 介值定理、最值定理、不动点定理 | 5 |
| 9 | [`Ch09_Derivative`](Analysis/Ch09_Derivative.lean) | 导数的线性逼近定义、可导必连续、中值定理 | 6 |

共 62 道练习，参考答案在 [`Analysis/Solutions/`](Analysis/Solutions/)。

## 两种学习方式

### 1. 在浏览器中（无需安装）

用浏览器打开 [`book/index.html`](book/index.html)（网页版教材）。它包含全部讲解、例题与练习，并提供：

* **在浏览器中做这道题**：在 [Lean 4 Web](https://live.lean-lang.org) 中打开本章到该练习为止的代码，直接在网页里写证明；
* **查看参考答案**：每道练习的完整证明；
* **ε-N 与 ε-δ 交互演示**：拖动 ε，观察 N 或 δ 如何变化；
* **学习进度**：标记已完成的练习（只保存在你自己的浏览器中）。

> 提示：直接双击打开 `book/index.html` 即可，不需要服务器。首次在 Lean 4 Web 中加载 Mathlib 需要等待十几秒。

### 2. 在本地（推荐长期学习）

1. 按照 [Lean 官方说明](https://lean-lang.org/install/) 安装 VS Code 与 Lean 4 扩展（会自动安装 `elan`）。
2. 获取本书并下载预编译的 Mathlib（约 5 GB，只需一次）：

   ```bash
   git clone https://github.com/kaoxi998533/pie-slang.git
   cd pie-slang/lean-analysis
   lake exe cache get
   ```

3. 用 VS Code 打开 `lean-analysis` 目录，从 `Analysis/Ch00_Welcome.lean` 开始阅读。
   把光标放在证明的每一行上，右侧的 Lean Infoview 会显示当前目标。

## 每章的结构

每章都是一个独立的 Lean 文件（只依赖 `import Mathlib`），因此可以整章复制到 Lean 4 Web 中运行。

* `/-! ... -/` 注释中是讲解，用 Markdown 书写，公式用 `$...$`；
* **例题**是完整的证明，可以逐行观察；
* **练习**以 `sorry` 结尾，提示写在题目前的讲解里；
* 后面章节需要的定义会在该章开头的"预备"一节中重新给出。

## 维护者说明

```bash
lake build                     # 检查全部章节与参考答案（练习中的 sorry 只产生警告）
node scripts/build-book.mjs    # 由章节文件重新生成网页版 book/index.html
```

`build-book.mjs` 会核对每道练习在章节文件和答案文件中的**陈述完全一致**，不一致时报错。
修改章节内容后请重新生成网页版并一同提交。

版本：Lean `v4.35.0`，Mathlib `v4.35.0`。在这个版本中，否定化简策略叫做 `push Not`
（较早的 Mathlib 中叫做 `push_neg`）。

## 延伸阅读

* [Mathematics in Lean](https://leanprover-community.github.io/mathematics_in_lean/)
* [Mathlib 文档](https://leanprover-community.github.io/mathlib4_docs/)
* [Theorem Proving in Lean 4](https://lean-lang.org/theorem_proving_in_lean4/)
