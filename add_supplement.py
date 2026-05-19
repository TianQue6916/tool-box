
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
import os

output_dir = r"c:\Users\27063\Desktop\AI   文件夹"
doc_path = os.path.join(output_dir, "计算机科学数学_全书重点总结_v2.docx")

doc = Document(doc_path)

def bp(doc, bold_text, normal_text):
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

# ================================================================
# SUPPLEMENT: Missing Sections
# ================================================================
doc.add_page_break()
doc.add_heading("Supplement: Key Missing Topics (补充：遗漏的重要主题)", level=1)
doc.add_paragraph(
    "以下内容是初版总结中遗漏的重要主题，它们在原书中占有重要篇幅，"
    "且对理解 Computer Science 的数学基础至关重要。"
)

# ---- Ch1.9 Good Proofs ----
doc.add_heading("1.9 Good Proofs in Practice (良好证明的实践)", level=2)
doc.add_paragraph(
    "原书不仅教你如何证明，还教你如何写好证明。这是很多数学教材忽略的 Practical Skill (实用技能)。"
)
bp(doc, "Good Proof 的标准：", "")
bullet(doc, "A good proof should be clear, correct, and complete (清晰、正确、完整)", "")
bullet(doc, "Each step should follow from previous steps by basic reasoning (每一步都应由前一步通过基本推理得出)", "")
bullet(doc, "The proof should be organized in a linear fashion (证明应以线性方式组织)", "")
bullet(doc, "Avoid excessive formal notation when plain language is clearer (当自然语言更清晰时避免过多形式化记号)", "")

bp(doc, "Bad Proof 的常见问题：", "")
bullet(doc, "Jumping steps without justification (跳步无理由)", "")
bullet(doc, "Using circular reasoning (循环论证)", "")
bullet(doc, "Confusing implication with biconditional (混淆蕴含与双条件)", "")
bullet(doc, "Proof by intimidation: claiming something is obvious when it is not (恐吓式证明：声称某事显然但其实不然)", "")

doc.add_paragraph(
    "Key Insight (核心洞见)：写证明就像写程序——要面向人类读者，要有良好的结构和命名，"
    "要有适当的注释。一个晦涩但正确的证明，价值不如一个清晰且正确的证明。"
    "这和 Software Engineering 中的 Code Readability (代码可读性) 是同一个道理。"
)

# ---- Ch3.2 Logic in Programs ----
doc.add_heading("3.2 Propositional Logic in Computer Programs (命题逻辑在程序中的应用)", level=2)
doc.add_paragraph(
    "这一节直接将 Logical Formula 与 Programming Language 联系起来，"
    "是全书最直接面向 CS 的内容之一。"
)

bp(doc, "Conditional Statement (条件语句)：", "")
doc.add_paragraph(
    "if 语句中的条件就是 Propositional Formula。例如 if (NOT (p AND q)) 等价于 if (NOT p OR NOT q)，"
    "这就是 De Morgan Law 在程序中的直接应用。理解 Logical Equivalence 可以帮助你 Simplify Code (简化代码) "
    "和 Avoid Redundant Conditions (避免冗余条件)。"
)

bp(doc, "Loop Invariant (循环不变式)：", "")
doc.add_paragraph(
    "while 循环的条件也是一个 Proposition。更重要的是，Loop Invariant 是证明循环正确性的核心工具："
    "它在循环开始前为 True，每次迭代后保持 True，循环结束时结合终止条件给出所需结论。"
    "这是 Hoare Logic (霍尔逻辑) 的基础，也是 Formal Verification (形式化验证) 的起点。"
)

bp(doc, "Short-circuit Evaluation (短路求值)：", "")
doc.add_paragraph(
    "在大多数编程语言中，p AND q 如果 p 为 False 则不计算 q，p OR q 如果 p 为 True 则不计算 q。"
    "这不仅是 Performance Optimization (性能优化)，更是 Semantic Difference (语义差异)——"
    "当 q 有 Side Effect (副作用) 时，短路求值和非短路求值的结果不同。"
)

# ---- Ch5.4 State Machines ----
doc.add_heading("5.4 State Machines (状态机)", level=2)
doc.add_paragraph(
    "State Machine 是 Induction 的最重要的 CS 应用之一。"
    "它将程序抽象为 States (状态) 和 Transitions (转移)，"
    "然后用 Invariant (不变式) 证明程序的正确性。"
)

bp(doc, "State Machine 的定义：", "")
bullet(doc, "States (状态)：程序可能处于的所有 Configuration (配置)", "")
bullet(doc, "Transitions (转移)：从一个状态到另一个状态的有向边", "")
bullet(doc, "Start State (起始状态)：程序的初始配置", "")
bullet(doc, "Invariant (不变式)：在所有 Reachable State (可达状态) 中保持为 True 的性质", "")

bp(doc, "用 Invariant 证明正确性的模板：", "")
doc.add_paragraph(
    "Step 1: 定义 State Machine (将程序抽象为状态机)\n"
    "Step 2: 声明 Invariant P(s) (不变式)\n"
    "Step 3: 证明 P 在 Start State 成立 (基础步)\n"
    "Step 4: 证明若 P 在状态 s 成立，则 P 在 s 的所有后继状态也成立 (归纳步)\n"
    "Step 5: 结合 Invariant 和终止条件，推出程序的正确性结论"
)

bp(doc, "经典应用：", "")
bullet(doc, "证明 Fast Exponentiation (快速幂) 算法的正确性", "")
bullet(doc, "证明 GCD 算法终止且结果正确", "")
bullet(doc, "证明 Distributed Algorithm (分布式算法) 的 Safety Property (安全性)", "")

doc.add_paragraph(
    "Key Insight (核心洞见)：State Machine + Invariant 是 Program Verification (程序验证) 的标准范式。"
    "它本质上就是 Induction——对 Reachable State 的集合做归纳。"
    "这个思想在 Model Checking (模型检测)、Concurrent Programming (并发编程)、"
    "Protocol Verification (协议验证) 中无处不在。"
)

# ---- Ch6.2 Matched Brackets ----
doc.add_heading("6.2 Strings of Matched Brackets (匹配括号字符串)", level=2)
doc.add_paragraph(
    "这一节是 Structural Induction (结构归纳法) 的经典教学案例，"
    "也是 Compiler (编译器) 中 Parsing (语法分析) 的数学基础。"
)

bp(doc, "Matched Bracket 的递归定义：", "")
bullet(doc, "空串是 Matched (匹配的)", "Base Case：")
bullet(doc, "若 S 是 Matched，则 [S] 也是 Matched", "Constructor 1：")
bullet(doc, "若 S 和 T 都是 Matched，则 ST 也是 Matched", "Constructor 2：")

bp(doc, "关键定理：", "")
doc.add_paragraph(
    "一个括号字符串是 Matched 当且仅当：(1) 从左到右扫描，右括号数不超过左括号数；"
    "(2) 左右括号总数相等。这个定理的证明是 Structural Induction 的完美示范。"
)

doc.add_paragraph(
    "Key Insight (核心洞见)：Recursive Definition (递归定义) 和 Structural Induction (结构归纳法) "
    "是处理 Recursive Data Type (递归数据类型) 的标准方法。"
    "在 Compiler 中，这对应着 Context-Free Grammar (上下文无关文法) 和 Recursive Descent Parsing (递归下降分析)。"
)

# ---- Ch7.2 Halting Problem ----
doc.add_heading("7.2 The Halting Problem (停机问题)", level=2)
doc.add_paragraph(
    "停机问题是 Computability Theory (可计算性理论) 中最重要的结果，"
    "它证明了存在计算机永远无法解决的问题。"
)

bp(doc, "Halting Problem 的陈述：", "")
doc.add_paragraph(
    "给定一个程序 P 和输入 w，是否存在一个算法能够判断 P(w) 是否会终止 (Halt)？"
    "答案是：不存在这样的算法。这是一个 Absolute Impossibility Result (绝对不可能性结果)，"
    "不是因为我们不够聪明，而是逻辑上就不可能。"
)

bp(doc, "证明方法：Diagonalization (对角线论证)", "")
doc.add_paragraph(
    "假设存在这样的算法 Halt(P, w)。构造程序 D(w)：
"
    "  if Halt(w, w) then loop forever else halt
"
    "现在考虑 D(D)：
"
    "  若 D(D) halts，则 Halt(D, D) = true，则 D(D) loops forever——矛盾
"
    "  若 D(D) loops，则 Halt(D, D) = false，则 D(D) halts——矛盾
"
    "无论哪种情况都矛盾，故 Halt 不存在。"
)

doc.add_paragraph(
    "Key Insight (核心洞见)：这是 Godel Incompleteness Theorem (哥德尔不完备定理) 在 CS 中的等价物。"
    "它告诉我们：Computability (可计算性) 是有边界的，不是所有问题都能用算法解决。"
    "这个结果深刻影响了整个 CS 的发展方向——从追求万能算法转向研究问题的 Complexity Class (复杂性类)。"
)

# ---- Ch8.3 Prime Mysteries ----
doc.add_heading("8.3 Prime Mysteries (素数之谜)", level=2)
doc.add_paragraph(
    "这一节展示了 Number Theory 中关于素数的未解之谜，"
    "以及已解决的重要定理，是理解 Cryptography (密码学) 安全性基础的关键。"
)

bp(doc, "已解决的重要结果：", "")
bullet(doc, "素数有无穷多个——Euclid 的经典证明（反证法）", "Euclid Theorem：")
bullet(doc, "每个大于1的整数可唯一分解为素数之积", "Fundamental Theorem of Arithmetic (算术基本定理)：")
bullet(doc, "存在任意大的素数间隙——n!+2, n!+3, ..., n!+n 都是合数", "Prime Gap (素数间隙)：")

bp(doc, "未解之谜：", "")
bullet(doc, "每个大于2的偶数是否都能表示为两个素数之和？至今未被证明", "Goldbach Conjecture (哥德巴赫猜想)：")
bullet(doc, "是否存在无穷多对相差2的素数？2013年张益唐证明了存在无穷多对相差不超过7000万的素数，后被改进到246", "Twin Prime Conjecture (孪生素数猜想)：")
bullet(doc, "P(n) = n^2 + n + 41 对 n=0..39 都是素数，但 n=40 不是", "Prime-generating Polynomial (素数生成多项式)：")

doc.add_paragraph(
    "Key Insight (核心洞见)：Cryptography (密码学) 的安全性依赖于素数的分布性质。"
    "RSA 的安全性基于 Integer Factorization (整数分解) 的困难性——我们知道分解是可能的，"
    "但目前没有高效算法。如果有人发现了快速分解算法或 Quantum Computer (量子计算机) 实用化，"
    "RSA 就会被破解。理解素数的性质是理解现代密码学的基础。"
)

# ---- Ch8.5 Alan Turing ----
doc.add_heading("8.5 Alan Turing (艾伦·图灵)", level=2)
doc.add_paragraph(
    "原书专门用一节介绍 Alan Turing 的生平和贡献，这不仅是历史回顾，"
    "更是理解 CS 起源的重要背景。"
)

bp(doc, "Turing 的三大贡献：", "")
bullet(doc, "定义了 Turing Machine (图灵机)，奠定了 Computability Theory 的基础", "Turing Machine：")
bullet(doc, "证明了 Halting Problem 不可解，揭示了计算的边界", "Halting Problem：")
bullet(doc, "二战期间破解了 Enigma 密码，据估计将战争缩短了2年", "Enigma Code-breaking：")

doc.add_paragraph(
    "Key Insight (核心洞见)：Turing 的工作告诉我们，Computer Science 不仅仅是关于计算机的，"
    "它是关于 Computation (计算) 本质的学科——什么可以被计算，什么不可以，以及效率如何。"
    "这个思想贯穿了从 Cryptography 到 Artificial Intelligence 的整个 CS 领域。"
)

# ---- Ch8.8 Turing's Code ----
doc.add_heading("8.8 Turing's Code (图灵密码)", level=2)
doc.add_paragraph(
    "这是原书中一个精彩的 Historical Case Study (历史案例研究)，"
    "展示了 Modular Arithmetic (模运算) 在密码学中的实际应用和攻击方法。"
)

bp(doc, "Turing's Code Version 2.0：", "")
doc.add_paragraph(
    "加密：将明文 m 乘以一个 Secret Key (密钥) k，得到密文 c = m*k mod p，其中 p 是一个大素数。
"
    "解密：m = c * k^(-1) mod p，其中 k^(-1) 是 k 的 Multiplicative Inverse (乘法逆元)。"
)

bp(doc, "攻击方法：", "")
doc.add_paragraph(
    "如果攻击者知道两个密文 c1 = m1*k mod p 和 c2 = m2*k mod p，"
    "则 c1/c2 = m1/m2 mod p，泄露了明文的比例关系。
"
    "更严重的是，如果攻击者知道任何一个明文-密文对，就可以恢复密钥 k。
"
    "这个弱点说明了为什么 Modern Cryptography (现代密码学) 使用更复杂的数学结构。"
)

# ---- Ch9.6 Partial Orders ----
doc.add_heading("9.6 Partial Orders (偏序)", level=2)
doc.add_paragraph(
    "Partial Order 是 Scheduling (调度)、Version Control (版本控制)、"
    "和 Dependency Analysis (依赖分析) 的数学基础。"
)

bp(doc, "Partial Order 的定义：", "")
doc.add_paragraph(
    "集合 A 上的关系 R 是 Partial Order 当且仅当满足三个性质：
"
    "1. Reflexivity (自反性)：a R a
"
    "2. Antisymmetry (反对称性)：若 a R b 且 b R a，则 a = b
"
    "3. Transitivity (传递性)：若 a R b 且 b R c，则 a R c
"
    "记作 (A, R) 或 (A, <=)，读作偏序集 Poset。"
)

bp(doc, "重要概念：", "")
bullet(doc, "偏序中并非所有元素都可比较。如子集包含关系：{1} 和 {2} 互不包含", "Comparability (可比性)：")
bullet(doc, "偏序集中最大的全序子集。在 Task Scheduling 中对应最长依赖链", "Chain and Antichain (链与反链)：")
bullet(doc, "任何有限偏序集都可以扩展为全序——这是 Topological Sort (拓扑排序) 的数学基础", "Topological Sort (拓扑排序)：")

bp(doc, "经典应用：", "")
bullet(doc, "课程先修关系构成偏序，Topological Sort 给出合法选课顺序", "Course Prerequisites (课程先修)：")
bullet(doc, "软件包之间的依赖关系构成偏序", "Package Dependencies (包依赖)：")
bullet(doc, "Dilworth Theorem：偏序集的最少链分解数等于最长反链长度", "Dilworth Theorem (迪尔沃思定理)：")

# ---- Ch9.10 Equivalence Relations ----
doc.add_heading("9.10 Equivalence Relations (等价关系)", level=2)

bp(doc, "等价关系的定义：", "")
doc.add_paragraph(
    "集合 A 上的关系 R 是 Equivalence Relation 当且仅当满足：
"
    "1. Reflexivity (自反性)：a R a
"
    "2. Symmetry (对称性)：若 a R b，则 b R a
"
    "3. Transitivity (传递性)：若 a R b 且 b R c，则 a R c
"
    "等价关系将集合划分为互不相交的 Equivalence Class (等价类)。"
)

bp(doc, "CS 中的应用：", "")
bullet(doc, "Union-Find 数据结构维护的就是等价类的动态变化", "Union-Find (并查集)：")
bullet(doc, "两个状态等价当且仅当它们对任意输入产生相同输出", "State Minimization (状态最小化)：")
bullet(doc, "同余关系是等价关系，模运算在等价类上操作", "Modular Arithmetic (模运算)：")

# ---- Ch11.4 Isomorphism ----
doc.add_heading("11.4 Graph Isomorphism (图同构)", level=2)

bp(doc, "Graph Isomorphism 的定义：", "")
doc.add_paragraph(
    "两个图 G 和 H 同构，如果存在一个 Bijection (双射) f: V(G) -> V(H)，"
    "使得 (u,v) 是 G 的边当且仅当 (f(u), f(v)) 是 H 的边。"
    "直觉上：同构的图是同一个图的不同画法。"
)

bp(doc, "Graph Isomorphism Problem 的特殊地位：", "")
doc.add_paragraph(
    "判断两个图是否同构的问题既没有被证明是 P 也不被证明是 NP-Complete。"
    "它被认为是 NP-Intermediate (NP中间问题)——如果 P != NP，则存在既不在 P 中也不在 NP-Complete 中的问题。"
    "2015年 Babai 给出了 Quasi-polynomial Time (准多项式时间) 算法，这是重大突破。"
)

doc.add_paragraph(
    "Key Insight (核心洞见)：Graph Isomorphism 展示了 Computational Complexity 中一个深刻的现象——"
    "并非所有问题都能被清晰分类。有些问题似乎介于 Easy (P) 和 Hard (NP-Complete) 之间，"
    "这暗示着 Complexity Hierarchy (复杂性层次) 可能比我们想象的更丰富。"
)

# ---- Ch11.6 Stable Marriage ----
doc.add_heading("11.6 The Stable Marriage Problem (稳定婚姻问题)", level=2)
doc.add_paragraph(
    "这是算法设计中最优雅的问题之一，也是 2012 年诺贝尔经济学奖的核心内容。"
)

bp(doc, "问题定义：", "")
doc.add_paragraph(
    "n 个男人和 n 个女人，每人有一个对异性的 Preference List (偏好列表)。"
    "一个 Matching (匹配) 是男女的一一配对。"
    "一个匹配是 Stable (稳定的) 当且仅当不存在一对 (man, woman) 使得："
    "他们没有配对，但彼此都比当前伴侣更偏好对方。"
    "这样的对称为 Rogue Couple (私奔对)。"
)

bp(doc, "Gale-Shapley Algorithm：", "")
doc.add_paragraph(
    "算法流程：
"
    "1. 每个 Free Man (自由男人) 向他尚未求过婚的最偏好女人求婚
"
    "2. 女人收到求婚：若她未婚，则暂时接受；若她已订婚，则比较新求婚者与当前未婚夫，选择更偏好的
"
    "3. 被拒绝的男人变为 Free，继续向下一个偏好求婚
"
    "4. 重复直到所有人配对"
)

bp(doc, "关键性质：", "")
bullet(doc, "算法最多 n^2 轮后终止", "Termination (终止性)：")
bullet(doc, "结果一定是 Stable Matching", "Correctness (正确性)：")
bullet(doc, "每个男人得到所有稳定匹配中最优的伴侣", "Man-optimality (男方最优性)：")
bullet(doc, "每个女人得到所有稳定匹配中最差的伴侣", "Woman-pessimality (女方最差性)：")

doc.add_paragraph(
    "Key Insight (核心洞见)：Gale-Shapley 算法揭示了 Mechanism Design (机制设计) 中的一个深刻教训——"
    "主动求婚的一方获得最优结果，被动接受的一方获得最差结果。"
    "这个结论在 Real-world Matching Market (现实匹配市场) 中有重要应用："
    "医院-实习生匹配、学校录取、器官移植分配等。"
    "Lloyd Shapley 和 Alvin Roth 因此获得 2012 年 Nobel Prize in Economics (诺贝尔经济学奖)。"
)

# ---- Ch11.2 Sexual Demographics ----
doc.add_heading("11.2 Sexual Demographics in America (性别人口统计)", level=2)
doc.add_paragraph(
    "这一节用 Graph Theory 解释了一个看似矛盾的社会统计现象，"
    "是 Discrete Math 应用于 Social Science 的经典案例。"
)

bp(doc, "The Paradox (悖论)：", "")
doc.add_paragraph(
    "调查显示：异性恋男性报告的平均女性性伴侣数 远大于 "
    "异性恋女性报告的平均男性性伴侣数。这在数学上应该相等（因为每次异性关系涉及一男一女），"
    "但调查数据却显示男性报告的数量约为女性的 1.5-2 倍。"
)

bp(doc, "Graph Theory 解释：", "")
doc.add_paragraph(
    "将人作为 Vertex，性关系作为 Edge，构成 Bipartite Graph (二部图)。
"
    "数学上：sum of degrees on male side = sum of degrees on female side = |E|。
"
    "所以 Average Male Degree = |E|/|M|，Average Female Degree = |E|/|F|。
"
    "若男女数量相等，平均值必须相等。差异的可能原因：
"
    "1. 男性倾向于 Over-report (多报)，女性倾向于 Under-report (少报)
"
    "2. 男女对性关系的 Definition (定义) 不同
"
    "3. 采样偏差：High-degree Individual 更容易被调查到
"
    "4. Prostitution (卖淫) 使少数女性有极高的度数"
)

# ---- Monty Hall Problem ----
doc.add_heading("补充经典习题：Monty Hall Problem (三门问题)", level=2)
doc.add_paragraph(
    "这是 Conditional Probability 最著名的经典问题，也是 Intuition 最容易出错的地方。"
)

bp(doc, "问题描述：", "")
doc.add_paragraph(
    "有三扇门，一扇后面有车，两扇后面有羊。你选一扇门后，主持人（知道门后是什么）"
    "打开另一扇有羊的门。现在你可以换门或坚持。应该换吗？"
)

bp(doc, "正确答案：", "")
doc.add_paragraph(
    "应该换！换门的胜率是 2/3，不换的胜率是 1/3。

"
    "分析：
"
    "- 你最初选对的概率 = 1/3，此时换门会输
"
    "- 你最初选错的概率 = 2/3，此时主持人只能打开另一个羊门，换门必赢
"
    "- 所以换门赢的概率 = 2/3

"
    "用 Bayes Theorem 分析：
"
    "P(车在门1 | 主持人开门3) = P(主持人开门3 | 车在门1) * P(车在门1) / P(主持人开门3)
"
    "= (1/2) * (1/3) / (1/2) = 1/3
"
    "P(车在门2 | 主持人开门3) = P(主持人开门3 | 车在门2) * P(车在门2) / P(主持人开门3)
"
    "= 1 * (1/3) / (1/2) = 2/3"
)

doc.add_paragraph(
    "Key Insight (核心洞见)：Monty Hall 问题揭示了 Conditional Probability 的反直觉性。"
    "关键在于主持人不是随机开门——他知道车在哪里，这个 Information (信息) 改变了概率分布。"
    "如果主持人是随机开门且碰巧开到羊，则换不换概率相同（都是 1/2）。"
    "这个区别说明了 Information Structure (信息结构) 对概率推理的决定性影响。"
)

# ---- Ch4.5 Finite Cardinality ----
doc.add_heading("4.5 Finite Cardinality (有限基数)", level=2)

bp(doc, "核心定理：", "")
doc.add_paragraph(
    "集合 A 和 B 有相同的 Cardinality (基数) 当且仅当存在 Bijection (双射) f: A -> B。
"
    "这个看似简单的定义蕴含着深刻的计数原理：
"
    "- |A| = |B| 当且仅当存在 A 到 B 的双射
"
    "- |A| <= |B| 当且仅当存在 A 到 B 的单射
"
    "- |A| >= |B| 当且仅当存在 A 到 B 的满射
"
    "- Schroder-Bernstein Theorem：若 |A| <= |B| 且 |B| <= |A|，则 |A| = |B|"
)

bp(doc, "计数中的 Bijection Principle (双射原理)：", "")
doc.add_paragraph(
    "如果你要数集合 A 的元素个数，但 A 不好数，可以找一个与 A 等势的集合 B，"
    "数 B 的元素个数即可。这是 Combinatorics (组合学) 中最强大的计数策略之一。"
    "例如：数从 (0,0) 到 (n,n) 只走右和上的路径数 = 数 n 个 R 和 n 个 U 的排列数 = C(2n, n)。"
)

# ---- Save ----
output_path = os.path.join(output_dir, "计算机科学数学_全书重点总结_v2.docx")
try:
    doc.save(output_path)
    print(f"Document updated and saved to: {output_path}")
except PermissionError:
    alt_path = os.path.join(output_dir, "计算机科学数学_全书重点总结_v3.docx")
    doc.save(alt_path)
    print(f"Original file locked. Saved to: {alt_path}")
