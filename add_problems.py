
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os

output_dir = r'c:\Users\27063\Desktop\AI   文件夹'
doc_path = os.path.join(output_dir, '计算机科学数学_全书重点总结_v2.docx')

doc = Document(doc_path)

# ============ Helper Functions ============
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

# ============ Add Classic Problems Section ============
doc.add_page_break()
doc.add_heading('Classic Problems & Key Insights (经典习题与核心洞见)', level=1)
doc.add_paragraph(
    '本书的习题（Problems）是课程的精华所在。MIT的传统是：真正理解一个概念，不是通过阅读，而是通过做题。'
    '以下精选了各章最具代表性的习题，它们不仅训练特定技能，更蕴含着深刻的数学洞见。'
    '每道题都标注了它所训练的核心能力（Core Ability Trained）和背后的核心洞见（Key Insight）。'
)

# ---- Chapter 1 Problems ----
doc.add_heading('Chapter 1: What is a Proof? (经典习题)', level=2)

bp(doc, 'Problem 1.2 - Bogus Proof that 1=1 implies 2=1 (伪证：1=1推出2=1)：', '')
doc.add_paragraph(
    '这道题展示了一个经典的伪证：利用 sqrt(1) = +/-1 的 Ambiguity (歧义)，'
    '错误地推出了 2=1。题目要求你 Precisely identify the mistake (精确找出错误)。'
)
bullet(doc, '识别 Proof 中的 Logical Fallacy (逻辑谬误)', 'Core Ability (训练能力)：')
bullet(doc, 'Square Root Convention (平方根约定)：sqrt(r) 只指 Positive Square Root，这是消除歧义的关键约定。'
    '许多 Bogus Proof 都利用了 sqrt 的歧义性。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 1.5 - Surprise Quiz Paradox (惊喜测验悖论)：', '')
doc.add_paragraph(
    'Albert宣布下周会给学生一个 Surprise Quiz。学生推理：不可能是周五（因为到周四晚上就知道了），'
    '也不可能是周四（因为周三晚上就知道了）...依次类推，不可能在任何一天。结果周二测验时，真的是惊喜！'
)
bullet(doc, '识别 Self-Reference (自指) 推理中的漏洞', 'Core Ability (训练能力)：')
bullet(doc, '这是 Epistemology (认识论) 中的经典悖论，与 Godel Incompleteness Theorem (哥德尔不完备定理) 有深层联系。'
    '学生的推理看似严密，但 Knowledge about future (对未来的认知) 和 Actual event (实际事件) 之间存在微妙的逻辑鸿沟。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 1.7 - Irrational to Irrational Power (无理数的无理数次幂)：', '')
doc.add_paragraph(
    '无理数的无理数次幂可以是有理数吗？考虑 sqrt(2)^sqrt(2)：若它是有理数，则 a=sqrt(2), b=sqrt(2) 即为答案；'
    '若它是无理数，则 (sqrt(2)^sqrt(2))^sqrt(2) = sqrt(2)^2 = 2 是有理数。'
)
bullet(doc, 'Nonconstructive Proof (非构造性证明) 的运用', 'Core Ability (训练能力)：')
bullet(doc, 'Proof by Cases (分类证明) 可以证明存在性而不需要构造具体实例。'
    '这种 Nonconstructive Method 在数学中非常重要——有时我们知道某物存在却无法具体指出它。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 1.15 - Rational Root Theorem (有理根定理)：', '')
doc.add_paragraph(
    '整系数多项式的实根要么是整数要么是无理数。这直接推出：当k不是m次幂时，m次根号k是无理数。'
)
bullet(doc, '从 Special Case 推广到 General Theorem', 'Core Ability (训练能力)：')
bullet(doc, '许多看似不同的 Irrationality Proof (无理性证明) 其实是同一个定理的特例。'
    '数学的力量在于 Generalization (推广)——找到一个统一框架来涵盖多个具体结果。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 1.17 - No Polynomial Generates Only Primes (没有多项式只生成素数)：', '')
doc.add_paragraph(
    '证明：对于非常数整系数多项式q(n)，必存在n使q(n)不是素数。'
    '关键步骤：q(c*m)是q(0)=c的倍数，而非常数多项式趋向无穷，故必有合数值。'
)
bullet(doc, '用 Divisibility (整除性) 证明 Number Theory 结果', 'Core Ability (训练能力)：')
bullet(doc, 'n^2+n+41 对n=0..39都是素数，但n=40时=1681=41^2不是素数。'
    '这个例子完美说明了：有限验证不能代替 Proof！40个连续成功案例也不足以建立普遍真理。', 'Key Insight (核心洞见)：')

# ---- Chapter 2 Problems ----
doc.add_heading('Chapter 2: Well Ordering Principle (经典习题)', level=2)

bp(doc, 'Problem 2.8 - Lehman Equation (莱曼方程)：', '')
doc.add_paragraph(
    'Euler Conjecture (a^4+b^4+c^4=d^4) 被反例推翻，但 Lehman方程 (8a^4+4b^4+2c^4=d^4) 确实无正整数解。'
    '提示：考虑所有解中a的最小值。'
)
bullet(doc, 'Well Ordering Proof 的实际运用', 'Core Ability (训练能力)：')
bullet(doc, '同样的方程形式，一个有解一个无解——这再次说明 Intuition 不可靠。'
    'Well Ordering Principle 在 Diophantine Equation (丢番图方程) 的 Impossible Proof (不可能性证明) 中非常有用。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 2.10 - Postage Stamp Problem (邮票问题)：', '')
doc.add_paragraph(
    '用10分和15分邮票组合的邮费一定是5的倍数。证明使用 Well Ordering Principle，'
    '但有一个 subtle mistake：需要验证 Base Case（m必须大于某个值才能保证m-10或m-15非负）。'
)
bullet(doc, '发现和修复 Proof 中的 Subtle Gap', 'Core Ability (训练能力)：')
bullet(doc, 'Well Ordering Proof 中最容易忽略的就是 Base Case。'
    '最小反例m必须足够大，使得m-10或m-15仍然在定义域内。'
    '这个 Gap 在 Induction Proof 中同样常见——Base Case 不完整是最常见的证明错误。', 'Key Insight (核心洞见)：')

# ---- Chapter 3 Problems ----
doc.add_heading('Chapter 3: Logical Formulas (经典习题)', level=2)

bp(doc, 'SAT and 3-SAT Problems (可满足性问题)：', '')
doc.add_paragraph(
    '习题中大量出现 k-clause 的 Satisfiability 问题。核心结果：若 n < 2^k，'
    '则 n 个 k-clause 构成的集合一定可满足。证明使用 Expectation Method (期望法)：'
    '随机赋值下每个 k-clause 为真的概率 >= 1 - 1/2^k，期望真 clause 数 > n-1 >= 0，故存在全满足的赋值。'
)
bullet(doc, 'Probabilistic Method (概率法) 证明存在性', 'Core Ability (训练能力)：')
bullet(doc, '这是 Paul Erdos 发明的 Probabilistic Method 的入门案例：'
    '通过计算期望来证明存在性，而不需要构造具体实例。'
    '这种方法在 Combinatorics (组合学) 和 Theoretical CS 中极其重要。', 'Key Insight (核心洞见)：')

# ---- Chapter 5 Problems ----
doc.add_heading('Chapter 5: Induction (经典习题)', level=2)

bp(doc, "The All Horses Same Color Fallacy (所有马同色的谬误)：", '')
doc.add_paragraph(
    '经典的 Induction Fallacy：声称证明n匹马同色。Base Case n=1 成立。'
    'Inductive Step：n+1匹马中，前n匹同色（归纳假设），后n匹同色（归纳假设），'
    '所以n+1匹同色。错误出在n=1到n=2的步骤：两匹马{h1}和{h2}没有 Overlap (重叠)，无法传递颜色。'
)
bullet(doc, '识别 Induction Step 中的 Hidden Assumption (隐含假设)', 'Core Ability (训练能力)：')
bullet(doc, 'Induction 从 n=1 到 n=2 的第一步是最脆弱的。'
    '当 Inductive Step 依赖于两个子集的 Overlap 时，必须验证最小情况下 Overlap 确实存在。'
    '这个 Fallacy 是所有 CS 学生必须知道的经典陷阱。', 'Key Insight (核心洞见)：')

# ---- Chapter 8 Problems ----
doc.add_heading('Chapter 8: Number Theory (经典习题)', level=2)

bp(doc, 'RSA Implementation and Security Analysis (RSA实现与安全性分析)：', '')
doc.add_paragraph(
    '习题要求完整实现 RSA 的 Key Generation、Encryption、Decryption，并分析：'
    '为什么知道 n 和 e 不能推出 d？如果有人找到了高效的 Integer Factorization (整数分解) 算法，RSA会怎样？'
)
bullet(doc, '从数学原理到工程实现的完整链路', 'Core Ability (训练能力)：')
bullet(doc, 'Cryptographic Security (密码安全性) 建立在 Computational Hardness (计算困难性) 之上，'
    '而非 Mathematical Impossibility (数学不可能性)。分解大整数不是不可能，只是目前没有高效算法。'
    'Quantum Computing (量子计算) 中的 Shor Algorithm 可能改变这一切。', 'Key Insight (核心洞见)：')

# ---- Chapter 11 Problems ----
doc.add_heading('Chapter 11: Simple Graphs (经典习题)', level=2)

bp(doc, 'Problem 11.68 - 2-Colorability Characterization (二着色刻画)：', '')
doc.add_paragraph(
    '证明：图G是 2-colorable 当且仅当不含 Odd-length Closed Walk (奇数长度闭途径)。'
    '证明分两步：(1) 2-colorable 推出无奇闭途径（显然）；(2) 无奇闭途径推出 2-colorable（对 Spanning Tree 着色，再证明非树边也连接不同颜色）。'
)
bullet(doc, '利用 Spanning Tree 简化 Graph Problem', 'Core Ability (训练能力)：')
bullet(doc, 'Spanning Tree 是 Graph Theory 中的脚手架——许多图论问题可以先在树上解决，'
    '再推广到一般图。这种 Reduce to Tree (归约到树) 的策略在算法设计中非常常见。', 'Key Insight (核心洞见)：')

bp(doc, 'MST Uniqueness (最小生成树唯一性)：', '')
doc.add_paragraph(
    '若所有 Edge Weight (边权) 互不相同，则 MST 唯一。'
    '证明思路：假设有两个不同MST M和N，考虑在M中但不在N中的最小边e，'
    '则N+e包含 Cycle，该Cycle上必有比e重的边，替换后得到更小的生成树——矛盾。'
)
bullet(doc, 'Exchange Argument (交换论证)', 'Core Ability (训练能力)：')
bullet(doc, 'Exchange Argument 是证明 Greedy Algorithm (贪心算法) 正确性的核心方法：'
    '假设最优解不使用贪心选择，通过交换证明可以替换为贪心选择而不变差。'
    'MST、Huffman Coding 等贪心算法的正确性都依赖这种论证。', 'Key Insight (核心洞见)：')

# ---- Chapter 17 Problems ----
doc.add_heading('Chapter 17: Conditional Probability (经典习题)', level=2)

bp(doc, 'Problem 17.28 - Product Rule Does Not Imply Independence (乘法规则不推出独立性)：', '')
doc.add_paragraph(
    '构造三个事件A, B, C满足 P(A∩B∩C) = P(A)P(B)P(C) 但不 Mutual Independent (相互独立)。'
)
bullet(doc, '区分 Pairwise Independence (两两独立) 和 Mutual Independence (相互独立)', 'Core Ability (训练能力)：')
bullet(doc, 'Pairwise Independence 不等于 Mutual Independence！三个事件两两独立且满足乘法规则，'
    '仍可能不是相互独立的。这个区别在 Probabilistic Analysis 中至关重要——'
    '错误假设 Independence 是最常见的概率推理错误之一。', 'Key Insight (核心洞见)：')

bp(doc, "Simpson's Paradox in University Admission (辛普森悖论与大学录取)：", '')
doc.add_paragraph(
    '习题中的经典场景：每个院系分别看，女生录取率不低于男生；但合并数据后，女生整体录取率反而更低。'
    '法官的证明看似严密，但忽略了 Conditional Probability 的 Direction (方向) 问题。'
)
bullet(doc, "识别和避免 Simpson's Paradox", 'Core Ability (训练能力)：')
bullet(doc, 'Conditional Probability 的方向至关重要：P(A|B) 和 P(A|C) 分别看可能得出与 P(A|B∪C) 完全相反的结论。'
    '这在 Real-world Data Analysis (现实数据分析) 中极其常见——'
    'Confounding Variable (混淆变量) 可以完全逆转统计结论。', 'Key Insight (核心洞见)：')

# ---- Chapter 18 Problems ----
doc.add_heading('Chapter 18: Random Variables & Expectation (经典习题)', level=2)

bp(doc, 'Problem 18.21 - Flirting Couples (调情情侣对数)：', '')
doc.add_paragraph(
    "4x4教室中男女随机就座，若男女相邻则构成调情对，求 Expected Number (期望数)。"
    '解法：对每对相邻座位定义 Indicator Variable (指示变量) I_i，E[I_i]=P(男女相邻)=1/2，'
    '总期望 = E[sum I_i] = sum E[I_i] = 相邻对数 * 1/2。'
)
bullet(doc, 'Indicator Variable + Linearity of Expectation 的标准范式', 'Core Ability (训练能力)：')
bullet(doc, '这是 Linearity of Expectation 最典型的应用模式：'
    '(1) 将复杂 Random Variable 分解为 Indicator Variable 之和；'
    '(2) 每个 Indicator 的期望容易计算；'
    '(3) 由 Linearity 直接求和。'
    '这种方法不需要知道变量之间的 Independence 关系——这是 Linearity 最强大的地方。', 'Key Insight (核心洞见)：')

bp(doc, 'Problem 18.23 - k-SAT Satisfiability via Expectation (用期望法证明k-SAT可满足性)：', '')
doc.add_paragraph(
    'n个k-clause，若n < 2^k，则一定可满足。'
    '证明：随机赋值下每个clause为真概率=1-1/2^k，期望真clause数=n(1-1/2^k)>n-1，'
    '故存在赋值使所有clause为真。'
)
bullet(doc, 'Probabilistic Method 证明 Combinatorial Existence', 'Core Ability (训练能力)：')
bullet(doc, '期望 > 0 推出存在性——这是 Probabilistic Method 的核心思想。'
    '更一般地，若 E[X] >= mu，则存在 omega 使 X(omega) >= mu。'
    '这个原理在 Randomized Algorithm 和 Extremal Combinatorics (极值组合) 中无处不在。', 'Key Insight (核心洞见)：')

# ---- Chapter 19 Problems ----
doc.add_heading('Chapter 19: Deviation from the Mean (经典习题)', level=2)

bp(doc, 'Polling and Sample Size (民意调查与样本量)：', '')
doc.add_paragraph(
    '用 Chebyshev Inequality 估计：要使民意调查误差不超过 +/-3% 的概率 >= 95%，需要多少样本？'
    '答案：n >= 1/(4*0.03^2*0.95) 约等于 556。而用 Chernoff Bound 可以得到更紧的估计。'
)
bullet(doc, '用 Concentration Inequality 估计样本量', 'Core Ability (训练能力)：')
bullet(doc, '三种 Concentration Inequality 给出不同精度的估计：'
    'Markov 最粗（只需期望）、Chebyshev 中等（需要方差）、Chernoff 最精（需要独立性）。'
    '在实际应用中，选择哪种不等式取决于你有多少信息。'
    '这个 Polling Problem 是统计学中 Confidence Interval (置信区间) 的离散版本。', 'Key Insight (核心洞见)：')

# ---- Chapter 20 Problems ----
doc.add_heading('Chapter 20: Random Walks (经典习题)', level=2)

bp(doc, 'Expected Return Time and Stationary Distribution (期望返回时间与平稳分布)：', '')
doc.add_paragraph(
    '核心定理：对 Connected Undirected Graph，从顶点v出发的 Expected Return Time (期望返回时间) = 2|E|/deg(v)。'
    'Stationary Distribution (平稳分布)：pi(v) = deg(v)/(2|E|)。'
)
bullet(doc, '将 Graph Structure 与 Probability Distribution 联系起来', 'Core Ability (训练能力)：')
bullet(doc, 'Stationary Distribution 只依赖于 Degree (度数)，不依赖于图的其他结构。'
    '这意味着 PageRank 的基本版本对 High-degree Node 有偏好——'
    '这也是为什么 Google 引入了 Damping Factor (阻尼因子) 来修正这种偏差。', 'Key Insight (核心洞见)：')

# ---- Chapter 21 Problems ----
doc.add_heading('Chapter 21: Recurrences (经典习题)', level=2)

bp(doc, 'Akra-Bazzi Method (阿克拉-巴齐方法) - Master Theorem 的推广：', '')
doc.add_paragraph(
    'Master Theorem 只处理 T(n) = aT(n/b) + O(n^d) 的形式，但实际递推可能更复杂，如：'
    'T(n) = T(n/3) + T(2n/3) + O(n)（子问题大小不等）。'
    'Akra-Bazzi Method 可以处理 T(n) = sum a_i * T(n/b_i) + O(n^d) 的一般形式。'
)
bullet(doc, '处理 Non-uniform Divide-and-Conquer (非均匀分治) 的递推', 'Core Ability (训练能力)：')
bullet(doc, 'Master Theorem 的三种情况本质上是在比较 Subproblem Growth (子问题增长) 和 Merge Cost (合并代价)。'
    'Akra-Bazzi Method 延续了这个思想，但允许不同大小的子问题。'
    '理解这个核心思想比记住公式更重要——当你遇到新的递推形式时，可以从这个角度分析。', 'Key Insight (核心洞见)：')

# ============ Cross-chapter Insights ============
doc.add_heading('Cross-Chapter Insights (跨章节核心洞见)', level=1)

doc.add_paragraph(
    '以下是从习题中提炼的贯穿多个章节的核心思想，它们是本书最宝贵的思维财富。'
)

bp(doc, '1. The Probabilistic Method (概率法)：', '')
doc.add_paragraph(
    '从 Chapter 3 的 SAT 问题到 Chapter 18 的 k-clause 可满足性，'
    'Probability 不仅仅是计算工具，更是证明 Existence (存在性) 的强大方法。'
    '核心逻辑：若 E[X] > 0，则存在 omega 使 X(omega) > 0。'
    '这种方法由 Paul Erdos 系统化，已成为现代 Combinatorics 的基石。'
)

bp(doc, '2. The Exchange Argument (交换论证)：', '')
doc.add_paragraph(
    '从 Chapter 8 的 GCD 算法到 Chapter 11 的 MST 唯一性，'
    'Exchange Argument 是证明 Greedy Choice (贪心选择) 正确性的通用方法。'
    "核心逻辑：假设最优解不用贪心选择，通过交换证明可以替换而不变差。"
)

bp(doc, '3. Indicator Variable + Linearity (指示变量+线性性范式)：', '')
doc.add_paragraph(
    '从 Hat-Check Problem 到 Flirting Couples 到 k-SAT，'
    '这是概率论中最实用的计算范式：将复杂变量分解为 Indicator 之和，利用 Linearity 求期望。'
    '关键优势：不需要 Independence 假设！'
)

bp(doc, '4. Reduce to Simpler Structure (归约到更简结构)：', '')
doc.add_paragraph(
    '从 Chapter 7 的对角线论证到 Chapter 11 的 Spanning Tree 简化，'
    '数学中最重要的策略之一是将问题归约到更简单的结构。'
    "Graph Problem -> Tree Problem, Infinite -> Finite, General -> Special Case。"
)

bp(doc, '5. Finite Verification != Proof (有限验证不等于证明)：', '')
doc.add_paragraph(
    '从 n^2+n+41 的素数陷阱到 Euler Conjecture 的推翻，'
    '全书反复强调：无论验证了多少个 Case，都不能代替 Proof。'
    '这是 CS 中 Formal Verification (形式化验证) 存在的根本理由。'
)

# Save
output_path = os.path.join(output_dir, '计算机科学数学_全书重点总结_v2.docx')
try:
    doc.save(output_path)
    print(f"Document updated and saved to: {output_path}")
except PermissionError:
    alt_path = os.path.join(output_dir, '计算机科学数学_全书重点总结_v3.docx')
    doc.save(alt_path)
    print(f"Original file locked. Saved to: {alt_path}")
