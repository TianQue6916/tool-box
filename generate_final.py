
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os

output_dir = r"c:\Users\27063\Desktop\AI   文件夹"
img_dir = os.path.join(output_dir, "summary_images")

doc = Document()

# ============ Style Setup ============
style = doc.styles['Normal']
font = style.font
font.name = 'Microsoft YaHei'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
style.paragraph_format.line_spacing = 1.35
style.paragraph_format.space_after = Pt(4)

for i in range(1, 5):
    hs = doc.styles[f'Heading {i}']
    hs.font.name = 'Microsoft YaHei'
    hs.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    hs.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)

# ============ Helper Functions ============
def bp(doc, bold_text, normal_text):
    """Add paragraph with bold prefix"""
    p = doc.add_paragraph()
    r1 = p.add_run(bold_text)
    r1.bold = True
    r1.font.size = Pt(11)
    r2 = p.add_run(normal_text)
    r2.font.size = Pt(11)
    return p

def bullet(doc, text, prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    if prefix:
        r1 = p.add_run(prefix)
        r1.bold = True
        r1.font.size = Pt(11)
        r2 = p.add_run(text)
        r2.font.size = Pt(11)
    else:
        p.text = text
    return p

def add_img(doc, filename, width=Inches(5.5)):
    path = os.path.join(img_dir, filename)
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(path, width=width)

# ============ TITLE ============
title = doc.add_heading("Mathematics for Computer Science (计算机科学中的数学) \u2014 全书深度总结", level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in title.runs:
    run.font.size = Pt(24)

doc.add_paragraph("Based on MIT 6.042J / Textbook by Meyer & Lehman (基于MIT 6.042J课程教材)", style='Intense Quote')

# ============ OVERVIEW ============
doc.add_heading("Overview (全书总论)", level=1)

doc.add_paragraph(
    "本书是MIT Computer Science (计算机科学) 专业的核心数学教材，系统性地介绍了CS所需的数学基础。"
    "全书分为五大部分：Proofs (证明)、Structures (结构)、Counting (计数)、Probability (概率) 和 Recurrences (递推)，"
    "共21章。全书的核心思想是：Proof (证明) 是理解数学和计算机科学的 Cornerstone (基石)，"
    "而 Discrete Mathematics (离散数学) 结构是 Modeling (建模) 计算问题的核心工具。"
)

doc.add_heading("Logical Relationship of Five Parts (五大部分的逻辑关系)", level=2)
doc.add_paragraph(
    "Part I Proofs (证明) 是 Foundation (地基) \u2014 教你如何严格地推理；"
    "Part II Structures (结构) 是 Language (语言) \u2014 教你用数学对象来建模计算问题；"
    "Part III Counting (计数) 是 Tool (工具) \u2014 教你精确地数 How many possibilities (有多少种可能)；"
    "Part IV Probability (概率) 是 Upgrade (升级) \u2014 在 Uncertainty (不确定性) 的世界中做推理；"
    "Part V Recurrences (递推) 是 Synthesis (综合) \u2014 用 Recurrence Relation (递推关系) 把前面所有工具串联起来分析 Algorithm (算法)。"
)

doc.add_heading("Core Philosophy (核心哲学)", level=2)
doc.add_paragraph(
    "全书一个贯穿始终的哲学是：Intuition (直觉) 不可靠，Proof (证明) 才可靠。"
    "书中反复用 Counterexample (反例) 说明 \u201c看起来显然\u201d的 Proposition (命题) 可能是错的，"
    "而 \u201c看起来不对\u201d的命题可能是对的。例如，Euler (欧拉) 曾 Conjecture (猜想) 对于 n>2，"
    "n次幂的和至少需要n个n次幂才能凑成另一个n次幂，这个猜想看起来很自然，但被 Counterexample 推翻了。"
    "而 Four Color Theorem (四色定理) 看起来太强了，却是对的。"
)

# ================================================================
# PART I
# ================================================================
doc.add_heading("Part I: Proofs (第一部分：证明)", level=1)
doc.add_paragraph(
    "本部分共7章（Chapter 1-7），是全书的基础。Computer Science 区别于其他 Engineering (工程) 学科的一个核心特征是："
    "我们能够（也必须）证明 Program (程序) 和 System (系统) 的 Correctness (正确性)，而不仅仅是 Testing (测试)。"
    "Testing 只能发现 Bug，Proof 才能保证没有 Bug。Intel 在1990年代的 Floating-point Division Bug (浮点除法错误) "
    "造成了数亿美元损失，这推动了 Chip Verification (芯片验证) 中 Formal Proof (形式化证明) 的广泛应用。"
)

# ---- Chapter 1 ----
doc.add_heading("Chapter 1: What is a Proof? (第1章：什么是证明)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "这是全书的起点。本章要回答一个 Fundamental Question (根本问题)：在数学中，"
    "\u201cTruth (真理)\u201d是如何确立的？不同领域有不同的真理标准\u2014\u2014"
    "Court (法庭) 靠 Evidence (证据) 和 Jury (陪审团)，Science (科学) 靠 Experiment (实验)，"
    "Business (商业) 靠 Authority (权威)，Statistics (统计) 靠 Significance (显著性)。"
    "数学有自己独特的标准：从 Axiom (公理) 出发，通过 Logical Deduction (逻辑推导) 链到达 Conclusion (结论)。"
)

doc.add_heading("Core Abilities Trained (训练的核心能力)", level=3)
bullet(doc, "识别什么样的 Statement (语句) 可以被 Prove (证明) 或 Disprove (证伪)", "Proposition Awareness (命题意识)：")
bullet(doc, "理解 Axiom \u2192 Logical Deduction \u2192 Theorem 的证明链条", "Axiomatic Thinking (公理化思维)：")
bullet(doc, "掌握证明 Implication (蕴含)、Iff (当且仅当)、Contradiction (反证法)、Case Analysis (分类讨论) 的标准 Template (模板)", "Proof Construction (证明构造能力)：")
bullet(doc, "区分 \u201c看起来对\u201d和 \u201c严格对\u201d，警惕 Intuition Trap (直觉陷阱)", "Anti-intuition Awareness (反直觉意识)：")

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "1. Proposition (命题)：", "一个非真即假的 Declarative Statement (判断语句)。例如 \u201c2+3=5\u201d是 True Proposition (真命题)，\u201c2+3=6\u201d是 False Proposition (假命题)。")

bp(doc, "2. Predicate (谓词)：", "Truth Value (真值) 依赖于 Variable (变量) 的命题。如 P(n)=\u201cn是 Perfect Square (完全平方数)\u201d，P(4)为 True，P(5)为 False。Predicate 是后续 Logical Formula (逻辑公式) 和 Quantifier (量词) 的基础。")

bp(doc, "3. Axiom (公理)：", "不加证明即接受为真的 Basic Assumption (基本假设)。全书采用 \u201c高中数学常识\u201d作为 Axiomatic Foundation (公理基础)，而非最底层的 ZFC Set Theory (ZFC集合论) \u2014\u2014因为用 ZFC 证明 2+2=4 需要超过 20,000 步！")

bp(doc, "4. Logical Deduction Rules (逻辑推导规则)：", "最基本的是 Modus Ponens (假言推理)：从 P 和 P\u2192Q 推出 Q。推导规则必须是 Sound (可靠的)，即 Premise (前提) 为 True 时 Conclusion (结论) 必为 True。")

bp(doc, "5. Theorem (定理)、Lemma (引理)、Corollary (推论)：", "Theorem 是重要的 True Proposition，Lemma 是为证明 Theorem 服务的辅助命题，Corollary 是从 Theorem 几步就能推出的结论。界限并非绝对\u2014\u2014有些 Lemma 比 Theorem 还重要。")

doc.add_paragraph("")
bp(doc, "Proving Implication (证明蕴含) 的两种方法：", "")
bullet(doc, "Assume (假设) P 为 True，Deduce (推导) Q 为 True。这是最直接的方法。", "Method 1 - Direct Proof (直接证明)：")
bullet(doc, "证明 Contrapositive (逆否命题) \u00acQ\u2192\u00acP，因为 P\u2192Q 等价于 \u00acQ\u2192\u00acP。例如证明 \u201c若 r 是 Irrational Number (无理数) 则 \u221ar 也是无理数\u201d，直接证很绕，但证逆否 \u201c若 \u221ar 是 Rational Number (有理数) 则 r 是有理数\u201d就非常简洁。", "Method 2 - Contrapositive Proof (逆否证明)：")

bp(doc, "Proving If and Only If (证明当且仅当) 的两种方法：", "")
bullet(doc, "分别证明 P\u2192Q 和 Q\u2192P", "Method 1: Two-directional Proof (双向证明)")
bullet(doc, "构造 P\u2261R\u2081\u2261R\u2082\u2261...\u2261Q 的 Equivalence Chain (等价链)", "Method 2: Chain of Iff (等价链)")

bp(doc, "Proof by Contradiction (反证法)：", "Assume (假设) 命题为 False，Derive Contradiction (推出矛盾)。经典例子：\u221a2 是 Irrational Number。")

bp(doc, "Proof by Cases (分类讨论)：", "将证明分成若干 Case (情况) 分别证明。经典例子：6个人中必有3人互相认识或3人互相不认识\u2014\u2014这是 Ramsey Theory (拉姆齐理论) 的入门例子。")

doc.add_heading("Deep Intention (深层意图)", level=3)
doc.add_paragraph(
    "本章表面上是教你 \u201cProof 是什么\u201d，但更深层的是在培养一种 Thinking Habit (思维习惯)："
    "不要相信 Intuition，要相信 Logic。Four Color Theorem 看起来 \u201c不太可能对\u201d却是对的，"
    "Euler Conjecture 看起来 \u201c显然对\u201d却是错的，Goldbach Conjecture (哥德巴赫猜想) 至今不知道对不对。"
    "这些例子都在说同一件事：在数学中，\u201c看起来对\u201d不等于 \u201c确实对\u201d，只有 Proof 才能给出 Certain Answer (确定的答案)。"
)

# ---- Chapter 2 ----
doc.add_heading("Chapter 2: Well Ordering Principle (第2章：良序原理)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "良序原理是全书第一个真正的 Proof Tool (证明工具)。它看似简单\u2014\u2014"
    "\u201c每个非空 Natural Number Set (自然数集) 都有 Minimum Element (最小元素)\u201d\u2014\u2014"
    "但这个朴素的陈述蕴含着强大的证明力量，同时为 Chapter 5 的 Induction (归纳法) 奠定 Theoretical Basis (理论基础)。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Well Ordering Principle (良序原理)：", "每个非空 Natural Number Set 都有一个 Minimum Element。这是关于 Natural Number 的基本性质，当作 Axiom 接受。")

bp(doc, "Well Ordering Proof Template (良序证明模板)：", "")
add_img(doc, "well_ordering.png", Inches(5.5))

bp(doc, "Classic Application (经典应用)：", "证明 Summation Formula (求和公式) 1+2+3+...+n = n(n+1)/2。假设存在使公式不成立的 n，取 Minimum Counterexample (最小反例) m，则对 m-1 公式成立，由此推出对 m 也成立\u2014\u2014 Contradiction。")

doc.add_heading("Deep Intention (深层意图)", level=3)
doc.add_paragraph(
    "良序原理的深层意义在于：它提供了一种将 Infinite Verification (无限验证) 转化为 Finite Contradiction (有限矛盾) 的方法。"
    "要证明 \u201c对所有 Natural Number n，P(n) 成立\u201d，你不可能逐个验证 Infinite 个 n，"
    "但 Well Ordering Principle 让你只需处理 \u201cMinimum Counterexample\u201d 这一个情况。"
    "这种从 Infinite 到 Finite 的 Reduction (归约)，是 Mathematical Reasoning (数学推理) 的核心技巧之一。"
)

# ---- Chapter 3 ----
doc.add_heading("Chapter 3: Logical Formulas (第3章：逻辑公式)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "将日常语言中的 Reasoning (推理) Formalize (形式化)、Precise (精确化)。在 CS 中，Logic 是 Hardware Design (硬件设计/布尔电路)、"
    "Program Verification (程序验证/霍尔逻辑)、Database Query (数据库查询/SQL的数学基础)、"
    "Artificial Intelligence (人工智能/知识表示) 的共同 Foundation。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Propositional Logic Connectives (命题逻辑连接词)：", "")
bullet(doc, "Conjunction (合取) P\u2227Q（且）、Disjunction (析取) P\u2228Q（或）、Negation (否定) \u00acP（非）", "Basic Connectives (基本连接词)：")
bullet(doc, "P\u2192Q 等价于 \u00acP\u2228Q\u2014\u2014这是理解 Implication 的关键", "Implication (蕴含) P\u2192Q：")
bullet(doc, "P\u2194Q = (P\u2192Q)\u2227(Q\u2192P)", "Biconditional (双条件) P\u2194Q：")

bp(doc, "Logical Equivalence (逻辑等价) 与 Basic Laws (基本定律)：", "")
bullet(doc, "\u00ac(P\u2227Q) \u2261 \u00acP\u2228\u00acQ，\u00ac(P\u2228Q) \u2261 \u00acP\u2227\u00acQ", "De Morgan\u2019s Laws (德摩根定律)：")
bullet(doc, "P\u2192Q \u2261 \u00acQ\u2192\u00acP", "Contrapositive (逆否)：")
bullet(doc, "P\u2227(Q\u2228R) \u2261 (P\u2227Q)\u2228(P\u2227R)", "Distributive Law (分配律)：")

bp(doc, "Predicate Logic & Quantifiers (谓词逻辑与量词)：", "")
bullet(doc, "\u2200x.P(x) 表示 \u201cFor all x, P(x) is true (对所有x，P(x)为真)\u201d", "Universal Quantifier (全称量词) \u2200：")
bullet(doc, "\u2203x.P(x) 表示 \u201cThere exists x such that P(x) is true (存在某个x使P(x)为真)\u201d", "Existential Quantifier (存在量词) \u2203：")
bullet(doc, "\u00ac\u2200x.P(x) \u2261 \u2203x.\u00acP(x)，\u00ac\u2203x.P(x) \u2261 \u2200x.\u00acP(x)\u2014\u2014量词否定时翻转", "Negation of Quantifiers (量词的否定)：")
bullet(doc, "\u2200x\u2203y 和 \u2203y\u2200x 含义完全不同\u2014\u2014Quantifier Order (量词顺序) 不能随意交换", "Important Warning (重要警告)：")

bp(doc, "Satisfiability Problem (可满足性问题/SAT)：", "给定一个 Propositional Formula (命题公式)，是否存在一组 Truth Assignment (真值赋值) 使公式为 True？"
    "SAT 是第一个被证明为 NP-Complete (NP完全) 的问题（Cook-Levin Theorem），"
    "是 Computational Complexity Theory (计算复杂性理论) 的核心。")

# ---- Chapter 4 ----
doc.add_heading("Chapter 4: Mathematical Data Types (第4章：数学数据类型)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "介绍 Discrete Mathematics 的三大基本 Data Type (数据类型)：Set (集合)、Function (函数)、Relation (关系)。"
    "这些是所有高级 Mathematical Structure 和 CS 中 Data Modeling (数据建模) 的 \u201cAtom (原子)\u201d。"
    "Database (数据库) 是 Relation 的 Set，Program 是 Function，Type System (类型系统) 依赖于 Set Theory (集合论)。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Set (集合)：", "")
bullet(doc, "Subset (子集)、Proper Subset (真子集)、Power Set (幂集) P(A)（A的所有子集构成的集合，|P(A)|=2^|A|）", "")
bullet(doc, "Union (并) A\u222aB、Intersection (交) A\u2229b、Set Difference (差) A-B、Complement (补) \u0100", "")
bullet(doc, "Cartesian Product (笛卡尔积) A\u00d7B = {(a,b) | a\u2208A, b\u2208B}，|A\u00d7B| = |A|\u00b7|B|", "")
bullet(doc, "Set Builder Notation (集合构造记号)：{x \u2208 A | P(x)}", "")

bp(doc, "Function (函数)：", "")
bullet(doc, "f: A\u2192B 是 Surjection (满射) 若每个 b\u2208B 都有 a\u2208A 使 f(a)=b", "Surjection (满射)：")
bullet(doc, "f 是 Injection (单射) 若 f(a\u2081)=f(a\u2082) 则 a\u2081=a\u2082", "Injection (单射)：")
bullet(doc, "f 是 Bijection (双射) 当且仅当既是 Injection 又是 Surjection", "Bijection (双射)：")
bullet(doc, "Bijection Principle (双射原理)：若存在 A 到 B 的 Bijection，则 |A|=|B|\u2014\u2014这是 Counting (计数) 的基本工具", "Bijection Principle (双射原理)：")

bp(doc, "Binary Relation (二元关系)：", "")
bullet(doc, "Reflexivity (自反性)：\u2200a, aRa", "")
bullet(doc, "Symmetry (对称性)：aRb \u2192 bRa", "")
bullet(doc, "Antisymmetry (反对称性)：aRb \u2227 bRa \u2192 a=b", "")
bullet(doc, "Transitivity (传递性)：aRb \u2227 bRc \u2192 aRc", "")
bullet(doc, "Equivalence Relation (等价关系) = Reflexive + Symmetric + Transitive，与 Set Partition (集合划分) 一一对应", "")
bullet(doc, "Partial Order (偏序) = Reflexive + Antisymmetric + Transitive，是 Topological Sort (拓扑排序) 和 Lattice Theory (格论) 的基础", "")

# ---- Chapter 5 ----
doc.add_heading("Chapter 5: Induction (第5章：归纳法)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Induction 是 Discrete Mathematics 中最重要的 Proof Method (证明方法)，没有之一。"
    "Induction 与 Recursion (递归) 是一体两面：每个 Recursive Program 的 Correctness 都需要 Induction 来证明，"
    "每个 Inductive Definition 都对应一个 Recursive Algorithm。本章系统讲解 Ordinary Induction (普通归纳法) 和 "
    "Strong Induction (强归纳法)，并展示它们在 Program Verification 中的应用。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Ordinary Induction (普通归纳法)：", "")
add_img(doc, "induction.png", Inches(5.5))
doc.add_paragraph("证明对所有 n\u22650，P(n) 成立：")
doc.add_paragraph("  Base Case (基础步)：Prove P(0)")
doc.add_paragraph("  Inductive Step (归纳步)：Assume P(n) (Inductive Hypothesis/归纳假设)，Prove P(n+1)")

bp(doc, "Strong Induction (强归纳法)：", "Inductive Hypothesis 不再是 \u201cP(n) 成立\u201d，而是 \u201cP(0), P(1), ..., P(n) 都成立\u201d。"
    "在证明关于 Integer 的命题时特别有用，如：任何大于1的 Integer 都可以分解为 Prime Number (素数) 的乘积。")

bp(doc, "Equivalence (等价性)：", "Ordinary Induction、Strong Induction、Well Ordering Principle 三者等价\u2014\u2014任何一个都可以推出另外两个。")

bp(doc, "Common Pitfalls (常见错误)：", "")
bullet(doc, "Inductive Hypothesis 使用不当", "Misusing Inductive Hypothesis：")
bullet(doc, "Base Case 遗漏 n=0 或 n=1 的情况，导致 Inductive Chain (归纳链) 断裂", "Missing Base Case：")
bullet(doc, "\u201c所有马的颜色相同\u201d的 \u201c证明\u201d\u2014\u2014错误出在从 n=1 到 n=2 的 Inductive Step", "Classic Counterexample (经典反例)：")

bp(doc, "Induction & Program Verification (归纳法与程序验证)：", "")
doc.add_paragraph("Loop Invariant (循环不变式) 是 Induction 在 Program Verification 中的直接应用：")
doc.add_paragraph("  Base Case：Loop 开始前 Invariant 成立")
doc.add_paragraph("  Inductive Step：如果 Loop 开始时 Invariant 成立，则 Loop Body (循环体) 执行后 Invariant 仍然成立")
doc.add_paragraph("  Termination (终止)：Loop 结束时 Invariant 给出 Program 的 Correctness")
doc.add_paragraph("这种方法被 Intel 等公司用于 Chip Verification (芯片验证)。")

doc.add_heading("Deep Intention (深层意图)", level=3)
doc.add_paragraph(
    "Induction 的深层意义在于：它是一种处理 Infinite 的 Finite Method (有限方法)。"
    "你不需要验证 Infinite 个情况，只需要验证一个 Base Case 和一个 \u201cTransmission Rule (传递规则)\u201d。"
    "理解 Induction，就是理解了 \u201c如何用 Finite 的推理把握 Infinite 的对象\u201d。"
)

# ---- Chapter 6 ----
doc.add_heading("Chapter 6: Recursive Data Types (第6章：递归数据类型)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "将 Inductive Thinking 从 Natural Number 推广到更一般的 Recursive Structure (递归结构)。"
    "在 CS 中，Linked List (链表)、Tree (树)、Expression (表达式)、String (字符串) 等都是 Recursively Defined (递归定义) 的。"
    "本章教你对这类结构进行 Structural Induction (结构归纳法)，"
    "这是 Functional Programming (函数式编程) 和 Compiler Design (编译器设计) 中的核心 Reasoning Tool (推理工具)。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Recursive Definition Pattern (递归定义模式)：", "")
doc.add_paragraph("  Base Case (基础情况)：定义最简单的 Element")
doc.add_paragraph("  Constructor Rule (构造规则)：从已有 Element 构造新 Element 的规则")
doc.add_paragraph("例如，Matching Parentheses String (匹配括号字符串) 的递归定义：")
doc.add_paragraph("  Base：Empty String (空串) \u03b5 是匹配的")
doc.add_paragraph("  Rule 1：若 s 是匹配的，则 (s) 是匹配的")
doc.add_paragraph("  Rule 2：若 s 和 t 是匹配的，则 st 是匹配的")

bp(doc, "Structural Induction (结构归纳法)：", "Base Case 证明 Proposition 对 Base Case 成立，Inductive Step 假设 Proposition 对已有结构成立，"
    "证明通过 Constructor Rule 得到的新结构也满足 Proposition。"
    "Structural Induction 比 Natural Number Induction 更一般\u2014\u2014Natural Number Induction 是 Structural Induction 的特例。")

bp(doc, "Connection to CFG (与上下文无关文法的联系)：", "Recursive Data Type 与 Context-Free Grammar (上下文无关文法) 本质上是同一回事\u2014\u2014"
    "Base Case 对应 Terminal Symbol (终结符)，Constructor Rule 对应 Production Rule (产生式)。"
    "理解了这种联系，你就理解了为什么 Compiler 能 Parsing (解析) 程序代码。")

# ---- Chapter 7 ----
doc.add_heading("Chapter 7: Infinite Sets (第7章：无限集合)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "从 Finite World (有限世界) 进入 Infinite World (无限世界)。本章回答一个 Profound Question (深刻的问题)："
    "Infinity (无穷) 也有大小之分吗？Cantor (康托尔) 的回答是：是的。"
    "有些 Infinity 比另一些 Infinity 更大。这不仅是一个 Mathematical Surprise (数学惊奇)，"
    "更是理解 Computation Theory (计算理论) 中 Decidability (可判定性) 与 Computability (可计算性) 的基础\u2014\u2014"
    "Halting Problem (停机问题) 的 Undecidability (不可判定性) 本质上就是 Diagonal Argument (对角线论证法) 在计算领域的应用。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Countable Set (可数集)：", "")
bullet(doc, "Set S 是 Countable 当且仅当存在从 N 到 S 的 Bijection（或 Surjection）", "Definition (定义)：")
bullet(doc, "Natural Number N、Integer Z、Rational Number Q 都是 Countable 的", "Classic Examples (经典例子)：")
bullet(doc, "将 Positive Rational Number 排列成 Infinite Matrix (无穷矩阵)，沿 Diagonal (对角线) 枚举", "Proof that Q is Countable (有理数可数的证明)：")

bp(doc, "Uncountable Set & Diagonal Argument (不可数集与对角线论证法)：", "")
bullet(doc, "Real Number Set R 是 Uncountable 的", "")
bullet(doc, "Cantor 的证明：假设 [0,1] 中的 Real Number 可 Enumerate (枚举) 为 r\u2081,r\u2082,...，"
    "构造 Real Number d 使其第 i 位与 r\u1d62 的第 i 位不同，则 d 不在 Enumeration 中\u2014\u2014Contradiction", "Diagonal Argument (对角线论证)：")
bullet(doc, "Diagonal Argument 在 Logic (G\u00f6del Incompleteness Theorem/哥德尔不完备定理) 和 Computation Theory (Halting Problem/停机问题) 中都有 Core Application", "Far-reaching Impact (深远影响)：")

bp(doc, "Cantor\u2019s Theorem (康托尔定理/幂集定理)：", "")
bullet(doc, "对任何 Set A，|P(A)| > |A|", "")
bullet(doc, "Corollary (推论)：不存在 \u201c最大\u201d的 Infinity\u2014\u2014N, P(N), P(P(N)), ... 严格递增", "")

bp(doc, "Connection to Computation Theory (与计算理论的联系)：", "")
doc.add_paragraph(
    "Halting Problem（给定 Program P 和 Input x，P(x) 是否 Halt (停机)？）是 Undecidable 的。"
    "其证明本质上就是 Diagonal Argument 在 Computation 领域的应用："
    "假设存在判定 Halting 的 Program H，构造 Program D 使得 D(P) 的行为与 H(P,P) 相反，"
    "则 D(D) 产生 Contradiction。理解了 Diagonal Argument，你就理解了为什么有些问题是 Computer 永远无法解决的。"
)

# ================================================================
# PART II
# ================================================================
doc.add_heading("Part II: Structures (第二部分：结构)", level=1)
doc.add_paragraph(
    "本部分共5章（Chapter 8-12），介绍 Discrete Mathematics 的核心 Structure："
    "Number Theory (数论)、Directed Graph (有向图)、Communication Network (通信网络)、"
    "Simple Graph (无向图) 和 Planar Graph (平面图)。"
    "训练目标：让你学会将 Actual Problem (实际问题) Abstract (抽象) 为 Mathematical Structure，"
    "然后用 Structure 本身的 Property 来解决问题。"
)

# ---- Chapter 8 ----
doc.add_heading("Chapter 8: Number Theory (第8章：数论)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Number Theory 是 \u201cQueen of Mathematics (数学女王)\u201d，也是现代 Cryptography (密码学) 的 Mathematical Foundation (数学基础)。"
    "本章从最基本的 Divisibility (整除性) 出发，逐步构建 Modular Arithmetic (模运算) 理论，"
    "最终达到 RSA Public Key Cryptosystem (RSA公钥加密系统)\u2014\u2014这是 Pure Math 应用于 CS 最成功的案例之一。"
    "通过本章，你将理解 Online Banking (网上银行)、HTTPS、Digital Signature (数字签名) 背后的数学原理。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Divisibility & GCD (整除性与最大公约数)：", "")
bullet(doc, "a | b 表示 a Divides b (a整除b)，即 b = ka 对某个 Integer k", "")
bullet(doc, "gcd(a,b)：a 和 b 的 Greatest Common Divisor (最大公约数)", "")

bp(doc, "Euclidean Algorithm (欧几里得算法)：", "")
doc.add_paragraph("基于 gcd(a,b) = gcd(b, a mod b)，Recursive (递归) 计算直到 b=0。")
doc.add_paragraph("Time Complexity (时间复杂度)：O(log(min(a,b)))\u2014\u2014每两次 Recursive Call，参数至少减半。")
doc.add_paragraph("Extended Euclidean Algorithm (扩展欧几里得算法)：同时求出 gcd(a,b) 和满足 gcd(a,b)=sa+tb 的 Integer s,t（B\u00e9zout\u2019s Identity/贝祖定理）。")

bp(doc, "Modular Arithmetic (模运算)：", "")
bullet(doc, "a \u2261 b (mod m) 当且仅当 m | (a-b)", "Congruence (同余) 的定义：")
bullet(doc, "Congruence 是 Equivalence Relation，模 m 的 Equivalence Class 称为 Residue Class (剩余类)", "Equivalence Relation (等价关系)：")
bullet(doc, "Addition, Subtraction, Multiplication 在 Congruence 下保持", "Arithmetic Properties (运算性质)：")
bullet(doc, "Multiplicative Inverse (乘法逆元)：a 在模 m 下有 Inverse 当且仅当 gcd(a,m)=1", "Inverse (逆元)：")

bp(doc, "Fermat\u2019s Little Theorem & Euler\u2019s Theorem (费马小定理与欧拉定理)：", "")
bullet(doc, "Fermat\u2019s Little Theorem：若 p 是 Prime (素数) 且 gcd(a,p)=1，则 a^(p-1) \u2261 1 (mod p)", "")
bullet(doc, "Euler\u2019s Theorem：若 gcd(a,m)=1，则 a^\u03c6(m) \u2261 1 (mod m)，其中 \u03c6(m) 是 Euler\u2019s Totient Function (欧拉函数)（1到m中与m互素的Integer个数）", "")
bullet(doc, "Fermat\u2019s Little Theorem 是 Euler\u2019s Theorem 的 Special Case (特例)（p是Prime时 \u03c6(p)=p-1）", "")

bp(doc, "RSA Encryption System (RSA加密系统)：", "")
add_img(doc, "rsa.png", Inches(5.5))

# ---- Chapter 9 ----
doc.add_heading("Chapter 9: Directed Graphs (第9章：有向图)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Directed Graph 是 Modeling State Machine (状态机)、Dependency (依赖关系)、Network Flow (网络流) 等 Computation Problem 的核心 Structure。"
    "本章重点是 DAG (有向无环图) 和 Topological Sort (拓扑排序)\u2014\u2014"
    "这在 Task Scheduling (任务调度)、Compiler Design (编译器设计)、Data Analysis (数据分析) 中都有直接应用。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Directed Graph Basics (有向图基本概念)：", "")
bullet(doc, "Directed Graph G = (V, E)，V 是 Vertex Set (顶点集)，E 是 Directed Edge Set (有向边集)", "")
bullet(doc, "Directed Path (有向路径)：Vertex 序列 v\u2080\u2192v\u2081\u2192...\u2192v\u2096", "")
bullet(doc, "Directed Cycle (有向环)：起点和终点相同的 Directed Path", "")

bp(doc, "DAG & Topological Sort (有向无环图与拓扑排序)：", "")
bullet(doc, "DAG：不含有向环的 Directed Graph", "")
bullet(doc, "Topological Sort：将 DAG 的 Vertex 排成 Linear Sequence (线性序列)，使每条 Edge 的起点在终点之前", "")
bullet(doc, "DAG 存在 Topological Sort 当且仅当图 Acyclic (无环)", "")
bullet(doc, "Application：Course Prerequisite (课程先修关系)、Task Scheduling、Makefile Dependency Analysis (依赖分析)", "")

bp(doc, "Strongly Connected Components (强连通分量)：", "")
bullet(doc, "Vertex u 和 v Strongly Connected 当且仅当互相 Reachable (可达)", "")
bullet(doc, "任何 Directed Graph 都可以分解为 Strongly Connected Component 的 DAG", "")

# ---- Chapter 10 ----
doc.add_heading("Chapter 10: Communication Networks (第10章：通信网络)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "本章是 Graph Theory 的 Application Chapter (应用篇)，展示数学如何指导 Network Engineering Design (网络工程设计)。"
    "通过分析特定 Network Topology (网络拓扑) 的 Performance，让你理解 Engineering Design 中的 Mathematical Tradeoff (数学权衡)。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Performance Metrics (性能指标)：", "")
bullet(doc, "Switch Count (交换节点数)：Network Cost (网络成本)", "")
bullet(doc, "Network Diameter (网络直径)：最短路径的最大长度\u2014\u2014决定 Latency (延迟)", "")
bullet(doc, "Congestion (拥塞)：最繁忙 Link 上的最大 Path 数\u2014\u2014决定 Throughput (吞吐量)", "")

bp(doc, "Classic Network Topologies (经典网络拓扑)：", "")
bullet(doc, "Complete Binary Tree (完全二叉树)：Diameter O(log N)，但 Root Congestion 严重", "")
bullet(doc, "Butterfly Network (蝶形网络)：Diameter O(log N)，Congestion O(\u221aN)", "")
bullet(doc, "Benes Network：Non-blocking Network (无阻塞网络)，任意 Permutation (排列) 都可实现\u2014\u2014但节点数更多", "")

# ---- Chapter 11 ----
doc.add_heading("Chapter 11: Simple Graphs (第11章：简单图/无向图)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Undirected Graph (无向图) 是 Modeling Social Network (社交网络)、Communication Network、Molecular Structure (分子结构) 等的基础工具。"
    "本章系统介绍 Handshaking Lemma (握手定理)、Connectivity (连通性)、Tree (树)、Bipartite Graph (二部图)、"
    "Eulerian Path (欧拉路径)、Graph Coloring (图着色) 等。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Handshaking Lemma (握手定理)：", "Undirected Graph 中所有 Vertex 的 Degree (度数) 之和等于 2 倍 Edge 数：\u03a3 deg(v) = 2|E|。"
    "Corollary (推论)：Degree 为奇数的 Vertex 个数必为 Even (偶数)。这个 Theorem 虽然简单，但在 Graph Theory Proof 中经常使用。")

bp(doc, "Connectivity & Trees (连通性与树)：", "")
bullet(doc, "Connected Graph (连通图)：任意两 Vertex 间有 Path", "")
bullet(doc, "Tree (树)：Connected Acyclic Graph (连通无环图)，Edge 数 = Vertex 数 - 1", "")
bullet(doc, "Spanning Tree (生成树)：Connected Graph 的 Minimal Connected Subgraph (极小连通子图)", "")
bullet(doc, "Minimum Spanning Tree (最小生成树)：Edge Weight (边权) 之和最小的 Spanning Tree（Kruskal/Prim Algorithm）", "")

bp(doc, "Bipartite Graph (二部图)：", "")
bullet(doc, "Vertex Set 可分为两部分，所有 Edge 连接不同部分的 Vertex", "")
bullet(doc, "Equivalent Condition (等价条件)：Graph 是 Bipartite 当且仅当不含 Odd-length Cycle (奇数长度的环)", "")
bullet(doc, "Matching Problem (匹配问题)：Bipartite Graph 中的 Maximum Matching (最大匹配) 可用 Hungarian Algorithm (匈牙利算法) 求解", "")

bp(doc, "Eulerian Path & Circuit (欧拉路径与欧拉回路)：", "")
bullet(doc, "Eulerian Path：经过每条 Edge 恰好一次的 Path", "")
bullet(doc, "Eulerian Circuit：起点和终点相同的 Eulerian Path", "")
bullet(doc, "Eulerian Circuit 存在当且仅当 Graph Connected 且每个 Vertex Degree 为 Even", "")
bullet(doc, "Eulerian Path 存在当且仅当 Graph Connected 且恰有 0 个或 2 个 Odd-degree Vertex", "")

bp(doc, "Graph Coloring (图着色)：", "")
bullet(doc, "Chromatic Number \u03c7(G) (色数)：给 Vertex 着色使 Adjacent Vertex (相邻顶点) 颜色不同的最少颜色数", "")
bullet(doc, "Bipartite Graph 的 Chromatic Number 为 2", "")
bullet(doc, "Graph Coloring 是 NP-hard Problem\u2014\u2014没有已知的 Polynomial-time Algorithm (多项式时间算法)", "")

# ---- Chapter 12 ----
doc.add_heading("Chapter 12: Planar Graphs (第12章：平面图)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Planar Graph Theory 在 Circuit Board Design (电路板设计)、Map Coloring (地图着色) 等领域有直接应用。"
    "本章通过 Euler\u2019s Formula 和 Five Color Theorem (五色定理)，展示 Graph Theory 中 \u201cStructure determines Property (结构决定性质)\u201d 的深刻思想。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Euler\u2019s Formula (欧拉公式)：", "")
add_img(doc, "euler.png", Inches(4.5))
doc.add_paragraph("对 Connected Planar Graph：V - E + F = 2，其中 V 是 Vertex 数，E 是 Edge 数，F 是 Face 数（包括 Outer Face/外部面）。"
    "这是数学中最优美的公式之一，与 Polyhedron (多面体) 的 Euler Formula V-E+F=2 是同一回事。")

bp(doc, "Necessary Conditions for Planar Graph (平面图的必要条件)：", "")
bullet(doc, "若 G 是 Simple Planar Graph（V\u22653），则 E \u2264 3V - 6", "")
bullet(doc, "若 G 不含 Triangle (三角形)，则 E \u2264 2V - 4", "")
bullet(doc, "Corollary：K\u2085 和 K\u2083,\u2083 不是 Planar Graph", "")

bp(doc, "Kuratowski\u2019s Theorem (Kuratowski定理)：", "Graph 是 Planar 当且仅当它不包含 K\u2085 或 K\u2083,\u2083 的 Subdivision (细分)。这给出了 Planar Graph 的一个 Complete Characterization (完全刻画)。")

bp(doc, "Five Color Theorem (五色定理)：", "Planar Graph 的 Chromatic Number 不超过 5。Proof 使用 Euler\u2019s Formula 和 Induction，"
    "是 \u201cStructure + Induction\u201d 的经典组合。"
    "Four Color Theorem (四色定理)：Chromatic Number 不超过 4\u2014\u2014Proof 需要 Computer-assisted Proof (计算机辅助证明)，"
    "是第一个重要的 Computer-assisted Proof。")

# ================================================================
# PART III
# ================================================================
doc.add_heading("Part III: Counting (第三部分：计数)", level=1)
doc.add_paragraph(
    "本部分共3章（Chapter 13-15），介绍 Counting Technique 和 Asymptotic Analysis (渐近分析)。"
    "Counting 是 Algorithm Analysis 的核心\u2014\u2014你需要知道 How many possibilities (有多少种可能)、"
    "Algorithm 需要多少 Step (步骤)、某种情况发生的 Probability 是多少。"
)

# ---- Chapter 13 ----
doc.add_heading("Chapter 13: Sums and Asymptotics (第13章：求和与渐近)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Common Summation Formulas (常见求和公式)：", "")
bullet(doc, "Arithmetic Series (等差级数)：1+2+...+n = n(n+1)/2", "")
bullet(doc, "Geometric Series (等比级数)：1+r+r\u00b2+...+r\u207f = (r\u207f\u207a\u00b9-1)/(r-1)", "")
bullet(doc, "Harmonic Series (调和级数)：H_n = 1+1/2+1/3+...+1/n \u2248 ln n + \u03b3（\u03b3\u22480.5772 是 Euler-Mascheroni Constant/欧拉常数）", "")

bp(doc, "Asymptotic Notation (渐近记号)：", "")
bullet(doc, "f = O(g)：存在 c 和 n\u2080，n\u2265n\u2080 时 f(n)\u2264c\u00b7g(n)\u2014\u2014Upper Bound (上界)", "")
bullet(doc, "f = \u03a9(g)：存在 c 和 n\u2080，n\u2265n\u2080 时 f(n)\u2265c\u00b7g(n)\u2014\u2014Lower Bound (下界)", "")
bullet(doc, "f = \u0398(g)：f = O(g) 且 f = \u03a9(g)\u2014\u2014Tight Bound (紧界)", "")
bullet(doc, "f = o(g)：lim f(n)/g(n) = 0\u2014\u2014Strictly smaller (严格小于)", "")

bp(doc, "Stirling\u2019s Formula (Stirling公式)：", "n! \u2248 (n/e)\u207f \u00b7 \u221a(2\u03c0n)。这是估计 Factorial (阶乘) 和 Binomial Coefficient (组合数) 的核心工具。")

# ---- Chapter 14 ----
doc.add_heading("Chapter 14: Cardinality Rules (第14章：基数规则)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Product Rule (乘法规则)：", "若 |A|=a，|B|=b，则 |A\u00d7B|=a\u00b7b。用于 Step-by-step Counting (分步计数)：每步有若干 Choice，总 Choice 数等于各步 Choice 数之积。")

bp(doc, "Sum Rule (加法规则)：", "若 A 和 B Disjoint (不交)，则 |A\u222aB|=|A|+|B|。用于 Case-by-case Counting (分类计数)。")

bp(doc, "Permutation & Combination (排列与组合)：", "")
bullet(doc, "Permutation (排列)：从 n 个元素中选 k 个有序排列，P(n,k) = n!/(n-k)!", "")
bullet(doc, "Combination (组合)：从 n 个元素中选 k 个无序子集，C(n,k) = n!/(k!(n-k)!)", "")
bullet(doc, "Pascal\u2019s Rule (帕斯卡法则)：C(n,k) = C(n-1,k-1) + C(n-1,k)", "")

bp(doc, "Inclusion-Exclusion Principle (容斥原理)：", "")
doc.add_paragraph("Two Set (两个集合)：|A\u2081\u222aA\u2082| = |A\u2081|+|A\u2082|-|A\u2081\u2229A\u2082|")
doc.add_paragraph("n Set (n个集合)：|A\u2081\u222a...\u222aA\u2099| = \u03a3|A\u1d62| - \u03a3|A\u1d62\u2229A\u2c7c| + \u03a3|A\u1d62\u2229A\u2c7c\u2229A\u2096| - ...")
doc.add_paragraph("Application：Derangement Problem (错排问题)\u2014\u2014n个元素的 Permutation 中没有任何元素在原位的 Permutation 数")

bp(doc, "Pigeonhole Principle (鸽巢原理)：", "")
doc.add_paragraph("若 n+1 个 Object 放入 n 个 Box，则至少有一个 Box 有至少 2 个 Object。")
doc.add_paragraph("Generalized Form (推广形式)：若 N 个 Object 放入 k 个 Box，则至少有一个 Box 有至少 \u2308N/k\u2309 个 Object。")

# ---- Chapter 15 ----
doc.add_heading("Chapter 15: Generating Functions (第15章：生成函数)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "将 Counting Problem Algebraize (代数化)。Generating Function 是 Counting 中最 Elegant (优雅)、最强大的 Tool\u2014\u2014"
    "它将复杂的 Combinatorial Problem (组合问题) 转化为 Power Series (幂级数) 的 Algebraic Operation (代数运算)。"
    "Generating Function 也是求解 Recurrence Relation 的重要工具，连接了 Counting 和 Recurrence 两个主题。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Ordinary Generating Function / OGF (普通生成函数)：", "")
doc.add_paragraph("Sequence (序列) a\u2080,a\u2081,a\u2082,... 的 Generating Function 是 G(x) = a\u2080 + a\u2081x + a\u2082x\u00b2 + ...")
doc.add_paragraph("Generating Function 的 Coefficient (系数) 就是我们要 Count 的对象，x\u207f 的 Coefficient 就是 a\u2099。")

bp(doc, "Common Generating Functions (常见生成函数)：", "")
bullet(doc, "1,1,1,... \u2192 1/(1-x)", "")
bullet(doc, "1,2,3,... \u2192 1/(1-x)\u00b2", "")
bullet(doc, "1,k,k\u00b2,... \u2192 1/(1-kx)", "")
bullet(doc, "C(n,0),C(n,1),...,C(n,n) \u2192 (1+x)\u207f（Binomial Theorem/二项式定理）", "")

bp(doc, "Operations on Generating Functions (生成函数的运算)：", "")
bullet(doc, "Addition (加法)：对应 Coefficient 相加\u2014\u2014对应 Counting 的 Classification (分类)", "")
bullet(doc, "Multiplication / Convolution (乘法/卷积)：(G\u00b7H)\u2099 = \u03a3a\u1d62b\u2099\u208b\u1d62\u2014\u2014对应 Counting 的 Step-by-step (分步)", "")
bullet(doc, "Differentiation (求导)：G\u2019(x) 对应 Sequence a\u2081,2a\u2082,3a\u2083,...", "")

bp(doc, "Solving Recurrence with GF (用生成函数解递推)：", "")
doc.add_paragraph("以 Fibonacci Sequence (斐波那契数列) 为例：F(n)=F(n-1)+F(n-2)，F(0)=0，F(1)=1")
doc.add_paragraph("Generating Function G(x) = x/(1-x-x\u00b2)，通过 Partial Fraction Decomposition (部分分式分解) 可得 Binet Formula：F(n) = (\u03c6\u207f-\u03c8\u207f)/\u221a5")

# ================================================================
# PART IV
# ================================================================
doc.add_heading("Part IV: Probability (第四部分：概率)", level=1)
doc.add_paragraph(
    "本部分共5章（Chapter 16-20），是现代 CS 中最重要的 Mathematical Tool 之一。"
    "从 Machine Learning 到 Cryptography，从 Network Analysis 到 Randomized Algorithm (随机算法)，"
    "Probability Theory 无处不在。"
)

# ---- Chapter 16 ----
doc.add_heading("Chapter 16: Events and Probability (第16章：事件与概率)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Three Elements of Probability Space (概率空间三要素)：", "")
bullet(doc, "Sample Space \u03a9 (样本空间)：所有 Possible Outcome (可能结果) 的集合", "")
bullet(doc, "Event (事件)：Sample Space 的 Subset", "")
bullet(doc, "Probability Measure P (概率测度)：满足三条 Axiom (公理)\u2014\u2014Non-negativity (非负性)、Normalization (规范性, P(\u03a9)=1)、Additivity (可加性，Mutually Exclusive/互斥 Event 之并的概率等于概率之和)", "")

bp(doc, "Uniform Probability Model (等可能模型)：", "若 \u03a9 有 n 个 Equally Likely Outcome (等可能结果)，则 P(A) = |A|/n。"
    "这是最 Intuitive (直观) 的 Probability Model，将 Probability Problem 转化为 Counting Problem。")

# ---- Chapter 17 ----
doc.add_heading("Chapter 17: Conditional Probability (第17章：条件概率)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "Conditional Probability 是 Probability Theory 中最重要的 Concept。"
    "Bayes\u2019 Theorem (贝叶斯定理) 是 Machine Learning 和 Statistical Inference (统计推断) 的 Mathematical Foundation。"
    "本章特别强调 Conditional Probability 中的 Intuition Trap (直觉陷阱)\u2014\u2014人的 Intuition 在 Conditional Probability 问题上经常出错。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

add_img(doc, "bayes.png", Inches(4.5))

bp(doc, "Conditional Probability (条件概率)：", "P(A|B) = P(A\u2229B)/P(B)，表示 \u201cGiven B has occurred, what is the probability of A (已知B发生时A发生的概率)\u201d。")

bp(doc, "Bayes\u2019 Theorem (贝叶斯定理)：", "P(A|B) = P(B|A)\u00b7P(A)/P(B)")
doc.add_paragraph("这是 \u201cInverse Probability (逆概率)\u201d 问题的核心：从 Observed Result (观察到的结果) B Infer (推断) Cause (原因) A 的 Probability。")
doc.add_paragraph("Application：Spam Filtering (垃圾邮件过滤/Naive Bayes)、Medical Diagnosis (医疗诊断)、ML Classifier (机器学习分类器)。")

bp(doc, "Independence (独立性)：", "P(A\u2229B) = P(A)\u00b7P(B) 等价于 P(A|B)=P(A)，即 B 的发生不影响 A 的 Probability。")
doc.add_paragraph("Important：Mutually Exclusive \u2260 Independent！事实上，若 P(A)>0 且 P(B)>0，则 A 和 B Mutually Exclusive 意味着它们 Not Independent。")

bp(doc, "Classic Intuition Traps (经典直觉陷阱)：", "")
doc.add_paragraph("Prosecutor\u2019s Fallacy (检察官谬误)：P(Evidence|Innocent) 很小不等于 P(Innocent|Evidence) 很小。")
doc.add_paragraph("Simpson\u2019s Paradox (辛普森悖论)：在 Subgroup Data (分组数据) 中观察到的 Trend，在 Combined Data (合并数据) 中可能反转。")

# ---- Chapter 18 ----
doc.add_heading("Chapter 18: Random Variables & Expectation (第18章：随机变量与期望)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Random Variable (随机变量)：", "从 Sample Space 到 Real Number (实数) 的 Function。"
    "不是 \u201crandom\u201d 的也不是 \u201cvariable\u201d\u2014\u2014它是一个确定的 Function，只是 Input 是随机的。")

bp(doc, "Expected Value (期望值)：", "E[X] = \u03a3 x\u00b7P(X=x)。Expectation 是 Random Variable 的 \u201cWeighted Average (加权平均)\u201d。")

bp(doc, "Linearity of Expectation (期望的线性性)：", "")
doc.add_paragraph("E[X\u2081+X\u2082+...+X\u2099] = E[X\u2081]+E[X\u2082]+...+E[X\u2099]")
doc.add_paragraph("Key：不需要 X\u2081,...,X\u2099 Mutually Independent！这是 Linearity of Expectation 最强大的地方。")
doc.add_paragraph("Classic Application：Hat-Check Problem (帽子检查问题)\u2014\u2014n个人随机取帽子，Expectation 有1个人拿到自己的帽子（无论n多大）。")

bp(doc, "Variance (方差)：", "")
doc.add_paragraph("Var(X) = E[(X-E[X])\u00b2] = E[X\u00b2] - (E[X])\u00b2")
doc.add_paragraph("Variance 衡量 Random Variable 偏离 Expectation 的程度，是分析 Algorithm Stability (算法稳定性) 的关键指标。")

bp(doc, "Common Distributions (常见分布)：", "")
bullet(doc, "Bernoulli Distribution (伯努利分布)：P(X=1)=p, P(X=0)=1-p，E[X]=p", "")
bullet(doc, "Binomial Distribution (二项分布)：n次 Independent Bernoulli Trial 的成功次数，E[X]=np", "")
bullet(doc, "Geometric Distribution (几何分布)：首次成功所需的 Trial 次数，E[X]=1/p", "")

# ---- Chapter 19 ----
doc.add_heading("Chapter 19: Deviation from the Mean (第19章：偏离均值)", level=2)

doc.add_heading("Purpose (章节目的)", level=3)
doc.add_paragraph(
    "只知道 Expectation 是不够的\u2014\u2014你还需要知道 Random Variable 有多大的 Probability 偏离 Expectation。"
    "Concentration Inequality (集中不等式) 给出了这种 Probability 的 Upper Bound (上界)，"
    "是分析 Randomized Algorithm、Machine Learning Generalization Error (泛化误差)、Network Reliability (网络可靠性) 的核心工具。"
)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

add_img(doc, "concentration.png", Inches(5.5))

bp(doc, "Markov\u2019s Inequality (马尔可夫不等式)：", "")
doc.add_paragraph("P(X \u2265 c) \u2264 E[X]/c（X Non-negative 时）")
doc.add_paragraph("只需知道 Expectation 就能给出 Tail Probability (尾部概率) 的 Upper Bound，但 Bound 较 Loose (松)。")

bp(doc, "Chebyshev\u2019s Inequality (切比雪夫不等式)：", "")
doc.add_paragraph("P(|X-E[X]| \u2265 c) \u2264 Var(X)/c\u00b2")
doc.add_paragraph("利用 Variance 给出 Tighter Bound (更紧的界限)。Application：Opinion Poll (民意调查) 的 Sample Size Estimation (样本量估计)。")

bp(doc, "Chernoff Bound (切尔诺夫界)：", "")
doc.add_paragraph("对 Independent Bernoulli Random Variable 之和 X=X\u2081+...+X\u2099：")
doc.add_paragraph("P(X \u2265 (1+\u03b4)\u03bc) \u2264 (e\u1dbb/(1+\u03b4)^(1+\u03b4))^\u03bc，其中 \u03bc=E[X]")
doc.add_paragraph("Chernoff Bound 是 Exponential (指数级) 的 Tail Bound\u2014\u2014比 Chebyshev 的 1/c\u00b2 衰减快得多。")
doc.add_paragraph("Application：Load Balancing (负载均衡)、Randomized Algorithm 的 Error Probability Analysis (错误概率分析)。")

# ---- Chapter 20 ----
doc.add_heading("Chapter 20: Random Walks (第20章：随机游走)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

add_img(doc, "pagerank.png", Inches(5))

bp(doc, "Random Walk Basics (随机游走基本概念)：", "")
doc.add_paragraph("从某个 Vertex 出发，每步随机选择一条 Edge 走到 Neighbor。Core Questions：")
bullet(doc, "Expected steps to return to start (返回时间)", "")
bullet(doc, "Expected steps to visit all vertices (覆盖时间)", "")
bullet(doc, "Long-run proportion of time at each vertex (平稳分布)", "")

bp(doc, "PageRank Algorithm (PageRank算法)：", "")
doc.add_paragraph("Web Page Ranking (网页排名) = Stationary Distribution (平稳分布) of Random Walk on Graph。")
doc.add_paragraph("Core Idea：一个 Web Page 重要当且仅当很多重要的 Web Page Link (链接) 到它。")
doc.add_paragraph("Mathematical Model：以 Probability \u03b1 随机点击 Link，以 Probability 1-\u03b1 跳转到 Random Page（避免陷入 Dead End/死胡同）。")
doc.add_paragraph("Stationary Distribution 的 Existence (存在性) 和 Uniqueness (唯一性) 由 Markov Chain Theory (马尔可夫链理论) 保证。")

# ================================================================
# PART V
# ================================================================
doc.add_heading("Part V: Recurrences (第五部分：递推)", level=1)
doc.add_paragraph(
    "本部分只有1章（Chapter 21），但它是全书的 Synthesis (综合应用)。"
    "Recurrence Relation 是将前面所有工具串联起来的纽带："
    "Divide-and-Conquer Algorithm (分治算法) 产生 Recurrence Relation，"
    "Recurrence Relation 需要 Counting 和 Summation 来求解，"
    "Random Recurrence 需要 Probability Tool 来分析。"
)

# ---- Chapter 21 ----
doc.add_heading("Chapter 21: Recurrences (第21章：递推关系)", level=2)

doc.add_heading("Key Concepts (核心知识点详解)", level=3)

bp(doc, "Substitution Method (替换法)：", "Guess a Solution (猜一个解)，Verify by Induction (用归纳法验证)。"
    "Key 在于如何 Guess\u2014\u2014通常先 Expand (展开) 前几项找 Pattern (规律)。")

bp(doc, "Plug-and-Chug Method (展开法)：", "将 Recurrence Relation 反复 Expand，直到出现 Known Term，然后 Summation (求和)。")
doc.add_paragraph("例如：T(n) = T(n-1) + n \u2192 T(n) = T(0) + 1+2+...+n = O(n\u00b2)")

bp(doc, "Master Theorem (主定理)：", "")
add_img(doc, "master.png", Inches(5.5))
doc.add_paragraph("对 Recurrence T(n) = aT(n/b) + O(n\u1d48)：")
doc.add_paragraph("  Case 1：d < log_b(a) \u2192 T(n) = O(n^{log_b(a)})\u2014\u2014Leaves dominate (叶节点主导)")
doc.add_paragraph("  Case 2：d = log_b(a) \u2192 T(n) = O(n\u1d48 log n)\u2014\u2014Balanced (每层均衡)")
doc.add_paragraph("  Case 3：d > log_b(a) \u2192 T(n) = O(n\u1d48)\u2014\u2014Root dominates (根节点主导)")

bp(doc, "Classic Applications (经典应用)：", "")
bullet(doc, "Merge Sort (归并排序)：T(n)=2T(n/2)+O(n) \u2192 Case 2 \u2192 O(n log n)", "")
bullet(doc, "Binary Search (二分搜索)：T(n)=T(n/2)+O(1) \u2192 O(log n)", "")
bullet(doc, "Strassen Matrix Multiplication (Strassen矩阵乘法)：T(n)=7T(n/2)+O(n\u00b2) \u2192 Case 1 \u2192 O(n^2.81)", "")

# ================================================================
# SUMMARY TABLE
# ================================================================
doc.add_heading("Summary: Ability System Overview (全书能力体系总览)", level=1)

table = doc.add_table(rows=7, cols=4, style='Light Grid Accent 1')
headers = ["Part (部分)", "Core Math Tools (核心数学工具)", "Core Abilities Trained (训练的核心能力)", "CS Applications (计算机科学应用)"]
for i, h in enumerate(headers):
    table.cell(0, i).text = h

data = [
    ["I Proofs (证明)", "Induction (归纳法), Contradiction (反证法), Logical Formula (逻辑公式), Well Ordering (良序原理)",
     "Logical Reasoning (逻辑推理), Proof Construction (证明构造), Intuition-to-Proof (直觉转证明)",
     "Program Correctness (程序正确性), Formal Methods (形式化方法), Chip Verification (芯片验证)"],
    ["II Structures (结构)", "Number Theory (数论), Graph Theory (图论), Relation (关系), Network Topology (网络拓扑)",
     "Structural Modeling (结构化建模), Graph Analysis (图论分析), Modular Arithmetic (模运算推理)",
     "Cryptography/RSA (密码学), Network Design (网络设计), Scheduling (调度优化)"],
    ["III Counting (计数)", "Permutation/Combination (排列组合), Inclusion-Exclusion (容斥原理), Generating Function (生成函数), Asymptotics (渐近)",
     "Exact Counting (精确计数), Asymptotic Estimation (渐近估计), Algebraic Counting (代数化计数)",
     "Algorithm Complexity (算法复杂度), Combinatorial Optimization (组合优化), Probability Calculation"],
    ["IV Probability (概率)", "Bayes\u2019 Theorem (贝叶斯), Linearity of Expectation (期望线性性), Concentration Inequality (集中不等式)",
     "Probability Modeling (概率建模), RV Analysis (随机变量分析), Deviation Estimation (偏差估计)",
     "Machine Learning, Randomized Algorithm (随机算法), PageRank, Reliability (可靠性)"],
    ["V Recurrences (递推)", "Master Theorem (主定理), Substitution (替换法), Plug-and-Chug (展开法)",
     "Recurrence Setup & Solving (递推建立与求解), Divide-and-Conquer Analysis (分治分析)",
     "Divide-and-Conquer Algorithm (分治算法), Recursive Algorithm Complexity (递归算法复杂度)"],
    ["Synthesis (综合)", "Combination of All Above (以上所有工具的组合)",
     "Mathematical Modeling + Rigorous Reasoning + Algorithm Analysis (数学建模+严格推理+算法分析)",
     "Independently Solve Math Problems in CS Research (独立解决CS研究中的数学问题)"],
]

for i, row_data in enumerate(data):
    for j, cell_text in enumerate(row_data):
        table.cell(i+1, j).text = cell_text

doc.add_paragraph("")

# ================================================================
# STUDY ADVICE
# ================================================================
doc.add_heading("Study Advice (学习建议)", level=1)

doc.add_paragraph(
    "1. Proofs 是 Foundation 中的 Foundation。Induction 和 Logical Reasoning 如果不扎实，后续所有内容都会吃力。"
    "建议：每学一种 Proof Template，至少做5道题，直到 Template 变成 Muscle Memory (肌肉记忆)。", style='List Number')

doc.add_paragraph(
    "2. Structures 的核心是 \u201cModeling (建模)\u201d\u2014\u2014学会将 Actual Problem 翻译成 Graph Theory / Number Theory Language，"
    "这比记住 Theorem 更重要。建议：每学一种 Graph Structure，想一个 Real Application，自己画出 Graph 并分析。", style='List Number')

doc.add_paragraph(
    "3. Counting 的核心是 \u201cSystematic (系统化)\u201d\u2014\u2014不要靠 \u201c数数\u201d，要用 Rule 和 Formula。"
    "Inclusion-Exclusion 和 Generating Function 是最强大的两个 Tool，必须反复练习。", style='List Number')

doc.add_paragraph(
    "4. Probability 的核心是 \u201cLinearity of Expectation (期望线性性)\u201d\u2014\u2014这个技巧在 Algorithm Analysis 中出现频率极高，必须熟练掌握。"
    "同时要警惕 Conditional Probability 中的 Intuition Trap，Bayes\u2019 Theorem 必须能熟练运用。", style='List Number')

doc.add_paragraph(
    "5. Recurrences 的核心是 \u201cMaster Theorem (主定理)\u201d\u2014\u2014它是分析 Divide-and-Conquer Algorithm 的标准 Tool，"
    "必须能快速判断三种 Case。建议：对常见的 Divide-and-Conquer Algorithm 都用 Master Theorem 分析一遍。", style='List Number')

doc.add_paragraph(
    "6. 全书最重要的 Thinking Habit (思维习惯)：不要相信 Intuition，要相信 Proof。"
    "每当你觉得 \u201cObviously (显然)\u201d 的时候，问自己：我能 Rigorously Prove (严格证明) 吗？"
    "如果不行，你的 \u201cObviously\u201d 可能就是错的。", style='List Number')

# Save
output_path = os.path.join(output_dir, "计算机科学数学_全书重点总结_v2.docx")
doc.save(output_path)
print(f"Word document saved to: {output_path}")
