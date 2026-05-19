
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# Set default font
style = doc.styles['Normal']
font = style.font
font.name = '微软雅黑'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# ========== Title ==========
title = doc.add_heading('《计算机科学中的数学》全书重点总结', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph('（基于MIT课程教材 Mathematics for Computer Science）', style='Intense Quote')

# ========== Overview ==========
doc.add_heading('全书概览', level=1)
doc.add_paragraph(
    '本书是MIT计算机科学专业的核心数学教材，系统性地介绍了计算机科学所需的数学基础。'
    '全书分为五大部分：证明（Proofs）、结构（Structures）、计数（Counting）、概率（Probability）和递推（Recurrences），'
    '共21章。全书的核心思想是：证明是理解数学和计算机科学的基石，而离散数学结构是建模计算问题的核心工具。'
)

doc.add_paragraph(
    '每一部分都遵循"定义→定理→证明→应用"的严密逻辑链条，'
    '旨在训练学生从直觉思维转向严格的数学推理思维，培养构造性证明和算法分析能力。'
)

# ========== Part I ==========
doc.add_heading('第一部分：证明（Proofs）', level=1)
doc.add_paragraph(
    '本部分是全书的基础，目标是建立严格的数学证明思维。计算机科学不仅需要编程能力，'
    '更需要证明程序正确性、验证算法性质的能力。本部分训练的核心能力是：逻辑推理能力、'
    '证明构造能力、以及将直觉转化为严格论证的能力。'
)

# Ch1
doc.add_heading('第1章 什么是证明（What is a Proof）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '建立数学证明的基本概念框架，让学生理解什么是数学真理，以及如何通过逻辑推理确立真理。'
    '这是整本书的起点——如果你不知道什么是证明，就无法做任何严格的数学推理。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('命题识别与分类能力：区分命题、公理、定理、引理、推论')
doc.add_paragraph('逻辑演绎的基本素养：理解从公理出发通过逻辑推导得出结论的过程')
doc.add_paragraph('反证法的运用能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('命题（Proposition）：非真即假的判断语句', style='List Bullet')
doc.add_paragraph('公理（Axiom）：不加证明即接受为真的基本假设', style='List Bullet')
doc.add_paragraph('逻辑演绎（Logical Deduction）：从已知命题推出新命题的推理规则', style='List Bullet')
doc.add_paragraph('证明的定义：从公理出发，通过逻辑推导链到达目标命题', style='List Bullet')
doc.add_paragraph('反证法（Proof by Contradiction）：假设命题为假，导出矛盾', style='List Bullet')
doc.add_paragraph('经典案例：四色定理、费马大定理的历史与证明方法讨论', style='List Bullet')

# Ch2
doc.add_heading('第2章 良序原理（Well Ordering Principle）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '引入良序原理作为证明的基本工具。良序原理看似简单（每个非空自然数集有最小元），'
    '但它是数学归纳法的理论基础，也是证明许多离散数学命题的利器。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('利用极小元反例进行证明的能力')
doc.add_paragraph('将"无限情况"归约为"有限矛盾"的思维能力')
doc.add_paragraph('为后续归纳法学习打基础')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('良序原理：每个非空自然数集合都有最小元素', style='List Bullet')
doc.add_paragraph('良序证明模板：假设存在反例→取最小反例→推出矛盾', style='List Bullet')
doc.add_paragraph('应用：证明求和公式 1+2+...+n = n(n+1)/2', style='List Bullet')
doc.add_paragraph('整除性质的基本证明', style='List Bullet')
doc.add_paragraph('良序原理与归纳法的等价性预告', style='List Bullet')

# Ch3
doc.add_heading('第3章 逻辑公式（Logical Formulas）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '将日常语言中的推理形式化。计算机科学需要精确的语言来描述和验证逻辑关系，'
    '命题逻辑和谓词逻辑是程序验证、电路设计、数据库查询的数学基础。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('命题逻辑的形式化表达能力')
doc.add_paragraph('真值表分析与等价推理能力')
doc.add_paragraph('谓词逻辑与量词的使用能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('命题逻辑：合取(∧)、析取(∨)、否定(¬)、蕴含(→)、双条件(↔)', style='List Bullet')
doc.add_paragraph('真值表与逻辑等价', style='List Bullet')
doc.add_paragraph('德摩根定律、分配律等基本等价式', style='List Bullet')
doc.add_paragraph('谓词逻辑：全称量词(∀)与存在量词(∃)', style='List Bullet')
doc.add_paragraph('量词的否定：¬∀x.P(x) ≡ ∃x.¬P(x)', style='List Bullet')
doc.add_paragraph('有效推理与可靠性', style='List Bullet')
doc.add_paragraph('可满足性（SAT）问题简介——这是计算机科学中的核心NP问题', style='List Bullet')

# Ch4
doc.add_heading('第4章 数学数据类型（Mathematical Data Types）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '介绍离散数学的基本数据结构——集合、函数、关系。这些是所有高级数学结构和计算机科学中数据建模的基础。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('集合运算与证明能力')
doc.add_paragraph('函数性质（单射、满射、双射）的判断与证明')
doc.add_paragraph('关系性质的识别能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('集合：子集、并、交、差、补、幂集、笛卡尔积', style='List Bullet')
doc.add_paragraph('函数：单射(injection)、满射(surjection)、双射(bijection)', style='List Bullet')
doc.add_paragraph('双射与计数：双射原理（映射计数法）', style='List Bullet')
doc.add_paragraph('二元关系：自反、对称、传递、等价关系', style='List Bullet')
doc.add_paragraph('等价关系与划分的一一对应', style='List Bullet')
doc.add_paragraph('偏序关系与偏序集', style='List Bullet')

# Ch5
doc.add_heading('第5章 归纳法（Induction）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '归纳法是离散数学中最重要的证明方法，也是计算机科学中证明算法正确性的核心工具。'
    '本章系统讲解普通归纳法和强归纳法，并建立它们与良序原理的等价性。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('归纳证明的构造能力——这是CS最核心的证明技能')
doc.add_paragraph('识别归纳结构的能力（递归定义→归纳证明）')
doc.add_paragraph('强归纳法的灵活运用')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('普通归纳法：基础步P(0) + 归纳步P(n)→P(n+1)', style='List Bullet')
doc.add_paragraph('强归纳法：归纳假设可使用P(0),P(1),...,P(n)', style='List Bullet')
doc.add_paragraph('归纳法与良序原理的等价性', style='List Bullet')
doc.add_paragraph('归纳模板：明确声明、基础步、归纳假设、归纳步', style='List Bullet')
doc.add_paragraph('常见错误：归纳步中的逻辑漏洞', style='List Bullet')
doc.add_paragraph('应用：证明程序正确性、递归算法终止性', style='List Bullet')

# Ch6
doc.add_heading('第6章 递归数据类型（Recursive Data Types）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '将归纳思想从自然数推广到递归定义的数据结构。在计算机科学中，链表、树、表达式等都是递归定义的，'
    '需要结构归纳法来证明其性质。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('递归定义的构造能力')
doc.add_paragraph('结构归纳法的运用能力')
doc.add_paragraph('将递归数据结构转化为归纳证明的能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('递归定义：基础情况+构造规则', style='List Bullet')
doc.add_paragraph('字符串的递归定义', style='List Bullet')
doc.add_paragraph('结构归纳法：对递归定义的结构进行归纳', style='List Bullet')
doc.add_paragraph('匹配括号字符串的递归定义与证明', style='List Bullet')
doc.add_paragraph('上下文无关文法与递归数据类型的关系', style='List Bullet')
doc.add_paragraph('递归函数与计算：从递归定义到可计算性', style='List Bullet')

# Ch7
doc.add_heading('第7章 无限集合（Infinite Sets）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '从有限世界进入无限世界，理解不同"大小"的无穷。这不仅是一个数学惊奇，'
    '更是理解计算理论中可判定性与可计算性的基础。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('对无穷概念的严格数学理解')
doc.add_paragraph('对角线论证法的掌握')
doc.add_paragraph('可数性与不可数性的判断能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('可数集：与自然数一一对应的集合（如整数、有理数）', style='List Bullet')
doc.add_paragraph('不可数集：实数集比自然数集"大"', style='List Bullet')
doc.add_paragraph('康托尔对角线论证法：证明实数不可数', style='List Bullet')
doc.add_paragraph('幂集定理：|P(A)| > |A|，不存在最大无穷', style='List Bullet')
doc.add_paragraph('连续统假设简介', style='List Bullet')
doc.add_paragraph('与计算理论的联系：不可判定问题的存在性', style='List Bullet')

# ========== Part II ==========
doc.add_heading('第二部分：结构（Structures）', level=1)
doc.add_paragraph(
    '本部分介绍离散数学的核心结构——数论、图论和网络。这些结构是建模实际计算问题的数学语言，'
    '从密码学到网络设计到算法分析都离不开它们。训练的核心能力是：结构化建模能力和算法性质分析能力。'
)

# Ch8
doc.add_heading('第8章 数论（Number Theory）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '数论是现代密码学的数学基础。本章从整除性出发，逐步建立模运算理论，'
    '最终达到RSA公钥加密系统——这是数学应用于计算机科学最成功的案例之一。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('模运算的计算与推理能力')
doc.add_paragraph('算法设计与正确性证明能力（欧几里得算法）')
doc.add_paragraph('理解密码系统的数学原理')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('整除性与最大公约数（GCD）', style='List Bullet')
doc.add_paragraph('欧几里得算法：高效计算GCD', style='List Bullet')
doc.add_paragraph('贝祖定理：gcd(a,b) = sa + tb 的整数解', style='List Bullet')
doc.add_paragraph('模运算：同余关系、加法/乘法逆元', style='List Bullet')
doc.add_paragraph('费马小定理：a^(p-1) ≡ 1 (mod p)', style='List Bullet')
doc.add_paragraph('欧拉定理与欧拉函数φ(n)', style='List Bullet')
doc.add_paragraph('RSA加密系统：密钥生成、加密、解密的完整流程', style='List Bullet')
doc.add_paragraph('中国剩余定理', style='List Bullet')

# Ch9
doc.add_heading('第9章 有向图（Directed Graphs）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '有向图是建模状态机、依赖关系、网络流等计算问题的核心结构。'
    '本章建立图论的基本语言和证明方法。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('图论建模能力：将实际问题抽象为有向图')
doc.add_paragraph('路径与连通性的分析能力')
doc.add_paragraph('DAG与拓扑排序的应用能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('有向图的定义：顶点集+有向边集', style='List Bullet')
doc.add_paragraph('路径、环、可达性', style='List Bullet')
doc.add_paragraph('强连通分量', style='List Bullet')
doc.add_paragraph('有向无环图（DAG）与拓扑排序', style='List Bullet')
doc.add_paragraph('邻接矩阵与邻接表表示', style='List Bullet')
doc.add_paragraph('应用：任务调度、依赖分析', style='List Bullet')

# Ch10
doc.add_heading('第10章 通信网络（Communication Networks）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '将图论应用于实际网络设计。本章通过分析特定网络拓扑结构的性能，'
    '展示数学如何指导工程设计——这是理论与实践结合的典范。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('网络拓扑的分析与比较能力')
doc.add_paragraph('交换网络性能指标的数学分析')
doc.add_paragraph('用图论语言描述和评估工程系统')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('交换网络的基本指标：交换节点数、直径、链路数', style='List Bullet')
doc.add_paragraph('完全二叉树网络', style='List Bullet')
doc.add_paragraph('蝶形网络（Butterfly Network）', style='List Bullet')
doc.add_paragraph('Benes网络：无阻塞网络结构', style='List Bullet')
doc.add_paragraph('网络直径与拥塞的权衡', style='List Bullet')

# Ch11
doc.add_heading('第11章 简单图（Simple Graphs）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '无向图是建模社交网络、通信网络、分子结构等的基础工具。'
    '本章系统介绍无向图的性质和经典定理，为后续平面图等高级主题做准备。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('无向图的性质分析能力')
doc.add_paragraph('图着色问题的理解与处理')
doc.add_paragraph('二部图的识别与应用能力')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('无向图基本概念：度、路径、环、连通分量', style='List Bullet')
doc.add_paragraph('握手定理：度数之和 = 2×边数', style='List Bullet')
doc.add_paragraph('连通性：连通图、连通分量', style='List Bullet')
doc.add_paragraph('树：连通无环图，边数 = 顶点数 - 1', style='List Bullet')
doc.add_paragraph('二部图：可用两种颜色着色的图', style='List Bullet')
doc.add_paragraph('图着色：色数χ(G)的定义与计算', style='List Bullet')
doc.add_paragraph('欧拉路径与欧拉回路', style='List Bullet')
doc.add_paragraph('哈密顿路径（简介）', style='List Bullet')

# Ch12
doc.add_heading('第12章 平面图（Planar Graphs）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '平面图理论在电路设计、地图着色等领域有直接应用。'
    '本章通过欧拉公式和五色定理，展示图论中"结构决定性质"的深刻思想。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('平面图的判断与分析能力')
doc.add_paragraph('利用欧拉公式进行图论推导的能力')
doc.add_paragraph('理解图论中"不可能性"证明的方法')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('平面图定义：可在平面上画出且边不相交的图', style='List Bullet')
doc.add_paragraph('欧拉公式：V - E + F = 2（连通平面图）', style='List Bullet')
doc.add_paragraph('平面图的必要条件：E ≤ 3V - 6', style='List Bullet')
doc.add_paragraph('K₅和K₃,₃不是平面图（Kuratowski定理）', style='List Bullet')
doc.add_paragraph('五色定理：平面图的色数 ≤ 5', style='List Bullet')
doc.add_paragraph('四色定理简介', style='List Bullet')

# ========== Part III ==========
doc.add_heading('第三部分：计数（Counting）', level=1)
doc.add_paragraph(
    '本部分介绍计数技术和渐近分析。计数是算法分析的核心——你需要知道有多少种可能、'
    '算法需要多少步骤。训练的核心能力是：精确计数能力、渐近估计能力和生成函数的代数化计数能力。'
)

# Ch13
doc.add_heading('第13章 求和与渐近（Sums and Asymptotics）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '算法分析的核心技能——估计求和式的增长速度。当你分析一个算法的时间复杂度时，'
    '本质上就是在做求和与渐近分析。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('常见求和公式的熟练运用')
doc.add_paragraph('大O、Ω、Θ记号的准确使用')
doc.add_paragraph('渐近估计的直觉与技巧')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('等差/等比求和公式', style='List Bullet')
doc.add_paragraph('调和级数：H_n ≈ ln n', style='List Bullet')
doc.add_paragraph('大O记号：f = O(g) 表示f增长不超过g', style='List Bullet')
doc.add_paragraph('Ω记号：下界；Θ记号：紧界', style='List Bullet')
doc.add_paragraph('积分近似求和', style='List Bullet')
doc.add_paragraph('Stirling公式：n! ≈ (n/e)^n √(2πn)', style='List Bullet')

# Ch14
doc.add_heading('第14章 基数规则（Cardinality Rules）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '系统化计数方法。从乘法规则到容斥原理，这些是组合计数的核心工具，'
    '直接应用于概率计算和算法分析中的情况枚举。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('乘法规则与加法规则的灵活运用')
doc.add_paragraph('排列组合的计算能力')
doc.add_paragraph('容斥原理的运用——处理重叠计数问题')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('乘法规则：|A×B| = |A|·|B|', style='List Bullet')
doc.add_paragraph('加法规则：不交集的并 |A∪B| = |A|+|B|', style='List Bullet')
doc.add_paragraph('排列：P(n,k) = n!/(n-k)!', style='List Bullet')
doc.add_paragraph('组合：C(n,k) = n!/(k!(n-k)!)', style='List Bullet')
doc.add_paragraph('容斥原理：|A₁∪A₂| = |A₁|+|A₂|-|A₁∩A₂|', style='List Bullet')
doc.add_paragraph('广义容斥原理（n个集合）', style='List Bullet')
doc.add_paragraph('鸽巢原理及其应用', style='List Bullet')
doc.add_paragraph('多项式系数与划分', style='List Bullet')

# Ch15
doc.add_heading('第15章 生成函数（Generating Functions）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '将计数问题代数化。生成函数是计数中最优雅的工具——将复杂的组合问题转化为幂级数的代数运算，'
    '许多看似困难的计数问题在生成函数框架下变得系统可解。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('将计数问题转化为生成函数的能力')
doc.add_paragraph('幂级数的代数运算能力')
doc.add_paragraph('从生成函数提取系数的技巧')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('普通生成函数（OGF）的定义', style='List Bullet')
doc.add_paragraph('生成函数的加法、乘法（卷积）运算', style='List Bullet')
doc.add_paragraph('常见序列的生成函数：1/(1-x), e^x等', style='List Bullet')
doc.add_paragraph('用生成函数解递推关系', style='List Bullet')
doc.add_paragraph('用生成函数解组合计数问题', style='List Bullet')
doc.add_paragraph('部分分式分解求系数', style='List Bullet')

# ========== Part IV ==========
doc.add_heading('第四部分：概率（Probability）', level=1)
doc.add_paragraph(
    '本部分是现代计算机科学中最重要的数学工具之一。从机器学习到密码学，从网络分析到随机算法，'
    '概率论无处不在。训练的核心能力是：概率建模能力、随机变量分析能力和偏差估计能力。'
)

# Ch16
doc.add_heading('第16章 事件与概率（Events and Probability）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '建立概率论的基本框架。从样本空间和事件出发，定义概率测度，'
    '为后续所有概率分析奠定严格基础。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('概率空间的建模能力')
doc.add_paragraph('概率的基本计算能力')
doc.add_paragraph('将直觉概率转化为严格计算')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('样本空间、事件、概率测度', style='List Bullet')
doc.add_paragraph('概率公理：非负性、规范性、可加性', style='List Bullet')
doc.add_paragraph('等可能模型：古典概型', style='List Bullet')
doc.add_paragraph('概率的基本规则：加法规则、补规则', style='List Bullet')
doc.add_paragraph('包含-排斥的概率版本', style='List Bullet')
doc.add_paragraph('无穷样本空间与连续概率简介', style='List Bullet')

# Ch17
doc.add_heading('第17章 条件概率（Conditional Probability）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '条件概率是概率论中最重要的概念之一。贝叶斯定理是机器学习和统计推断的数学基础，'
    '理解条件概率对于正确处理不确定性至关重要。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('条件概率的计算与链式规则运用')
doc.add_paragraph('贝叶斯定理的应用能力')
doc.add_paragraph('识别和避免概率直觉错误（如检察官谬误）')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('条件概率：P(A|B) = P(A∩B)/P(B)', style='List Bullet')
doc.add_paragraph('乘法规则：P(A∩B) = P(A)·P(B|A)', style='List Bullet')
doc.add_paragraph('全概率公式', style='List Bullet')
doc.add_paragraph('贝叶斯定理：P(A|B) = P(B|A)P(A)/P(B)', style='List Bullet')
doc.add_paragraph('先验概率与后验概率', style='List Bullet')
doc.add_paragraph('独立性：P(A∩B) = P(A)P(B)', style='List Bullet')
doc.add_paragraph('条件独立性', style='List Bullet')
doc.add_paragraph('经典应用：医疗检测、垃圾邮件过滤', style='List Bullet')

# Ch18
doc.add_heading('第18章 随机变量（Random Variables）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '从事件概率过渡到数值分析。随机变量是量化随机现象的工具，'
    '期望和方差是分析随机算法性能的核心指标。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('随机变量的定义与分布分析')
doc.add_paragraph('期望的线性性质运用——最强大的概率工具之一')
doc.add_paragraph('方差的计算与意义理解')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('随机变量：从样本空间到实数的函数', style='List Bullet')
doc.add_paragraph('概率质量函数（PMF）', style='List Bullet')
doc.add_paragraph('期望：E[X] = Σ x·P(X=x)', style='List Bullet')
doc.add_paragraph('期望的线性性：E[X+Y] = E[X]+E[Y]（无论是否独立！）', style='List Bullet')
doc.add_paragraph('条件期望', style='List Bullet')
doc.add_paragraph('方差：Var(X) = E[(X-μ)²] = E[X²] - (E[X])²', style='List Bullet')
doc.add_paragraph('独立随机变量的方差可加性', style='List Bullet')
doc.add_paragraph('常见分布：伯努利、二项、几何、均匀', style='List Bullet')

# Ch19
doc.add_heading('第19章 偏离均值（Deviation from the Mean）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '仅有期望和方差是不够的——你需要知道随机变量偏离期望值的概率有多大。'
    '集中不等式是分析随机算法可靠性和机器学习泛化能力的核心工具。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('集中不等式的运用能力')
doc.add_paragraph('概率界限的估计能力')
doc.add_paragraph('理解"高概率"保证的含义')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('马尔可夫不等式：P(X≥a) ≤ E[X]/a', style='List Bullet')
doc.add_paragraph('切比雪夫不等式：P(|X-μ|≥k) ≤ Var(X)/k²', style='List Bullet')
doc.add_paragraph('切尔诺夫界：独立伯努利变量之和的指数集中性', style='List Bullet')
doc.add_paragraph('弱大数定律', style='List Bullet')
doc.add_paragraph('应用：负载均衡、哈希冲突分析', style='List Bullet')

# Ch20
doc.add_heading('第20章 随机游走（Random Walks）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '随机游走是概率论与图论的交汇点，在算法设计（如随机化算法、PageRank）中有广泛应用。'
    '本章将之前学到的概率工具应用于动态随机过程。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('随机过程的基本分析能力')
doc.add_paragraph('马尔可夫链的直观理解')
doc.add_paragraph('返回时间与期望步数的计算')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('一维随机游走：期望返回时间', style='List Bullet')
doc.add_paragraph('图上的随机游走', style='List Bullet')
doc.add_paragraph('平稳分布', style='List Bullet')
doc.add_paragraph('返回时间与图结构的关系', style='List Bullet')
doc.add_paragraph('应用：PageRank算法的数学原理', style='List Bullet')
doc.add_paragraph('赌徒破产问题', style='List Bullet')

# ========== Part V ==========
doc.add_heading('第五部分：递推（Recurrences）', level=1)
doc.add_paragraph(
    '递推关系是分析分治算法和递归算法时间复杂度的核心工具。'
    '训练的核心能力是：从递归算法导出递推关系，并求解递推关系得到闭式解。'
)

# Ch21
doc.add_heading('第21章 递推关系（Recurrences）', level=2)
doc.add_heading('章节目的', level=3)
doc.add_paragraph(
    '将递归算法的时间复杂度分析问题转化为递推关系的求解问题。'
    '本章提供系统的方法来求解常见的递推关系，是算法分析的必备技能。'
)
doc.add_heading('训练能力', level=3)
doc.add_paragraph('从递归算法导出递推关系的能力')
doc.add_paragraph('递推关系的求解能力')
doc.add_paragraph('主定理的运用——分析分治算法的利器')
doc.add_heading('核心知识点', level=3)
doc.add_paragraph('递推关系的定义：T(n) = aT(n/b) + f(n)', style='List Bullet')
doc.add_paragraph('展开法（迭代法）：反复展开递推', style='List Bullet')
doc.add_paragraph('主定理：分治递推的通用解', style='List Bullet')
doc.add_paragraph('主定理三种情况：根据f(n)与n^(log_b a)的关系', style='List Bullet')
doc.add_paragraph('线性递推的求解', style='List Bullet')
doc.add_paragraph('应用：归并排序O(n log n)、快速排序期望分析', style='List Bullet')

# ========== Summary ==========
doc.add_heading('全书能力训练体系总结', level=1)

table = doc.add_table(rows=6, cols=3)
table.style = 'Light Grid Accent 1'

# Header
hdr = table.rows[0].cells
hdr[0].text = '部分'
hdr[1].text = '核心训练能力'
hdr[2].text = '应用场景'

row1 = table.rows[1].cells
row1[0].text = '证明'
row1[1].text = '逻辑推理、证明构造、归纳思维'
row1[2].text = '程序正确性验证、算法证明'

row2 = table.rows[2].cells
row2[0].text = '结构'
row2[1].text = '结构化建模、图论分析、数论计算'
row2[2].text = '密码学、网络设计、调度优化'

row3 = table.rows[3].cells
row3[0].text = '计数'
row3[1].text = '精确计数、渐近估计、代数化计数'
row3[2].text = '算法复杂度分析、组合优化'

row4 = table.rows[4].cells
row4[1].text = '概率建模、随机变量分析、集中不等式'
row4[2].text = '机器学习、随机算法、可靠性分析'
row4[0].text = '概率'

row5 = table.rows[5].cells
row5[0].text = '递推'
row5[1].text = '递推关系建立与求解'
row5[2].text = '分治算法分析、递归算法复杂度'

doc.add_paragraph('')
doc.add_heading('学习建议', level=1)
doc.add_paragraph('证明部分是基础中的基础——如果归纳法和逻辑推理不扎实，后续所有内容都会吃力。务必多做题。', style='List Number')
doc.add_paragraph('结构部分注重建模——学会将实际问题翻译成图论/数论语言，这比记住定理更重要。', style='List Number')
doc.add_paragraph('计数部分的核心是系统化——不要靠"数数"，要用规则和公式。生成函数是最强大的工具。', style='List Number')
doc.add_paragraph('概率部分的核心是期望线性性——这个技巧在算法分析中出现频率极高，必须熟练掌握。', style='List Number')
doc.add_paragraph('递推部分的核心是主定理——它是分析分治算法的标准工具，必须能快速判断三种情况。', style='List Number')

doc.save('c:/Users/27063/Desktop/AI   文件夹/计算机科学数学_全书重点总结.docx')
print("Word文档已生成！")
