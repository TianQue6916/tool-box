# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# --- Style settings ---
style = doc.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(11)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.25

# Helper functions
def add_title(text, level=0):
    size_map = {0: Pt(22), 1: Pt(18), 2: Pt(15), 3: Pt(12)}
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = size_map.get(level, Pt(11))
    if level == 0:
        run.font.color.rgb = RGBColor(0x1a, 0x3c, 0x6e)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:
        run.font.color.rgb = RGBColor(0x2e, 0x5e, 0x8e)
    elif level == 2:
        run.font.color.rgb = RGBColor(0x3a, 0x7a, 0xa4)
    p.paragraph_format.space_before = Pt(12 if level > 0 else 0)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_body(text):
    p = doc.add_paragraph(text)
    return p

def add_bullet(text, level=0):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.left_indent = Cm(1.5 + level * 1.0)
    return p

# ============= TITLE PAGE =============
add_title('《计算机科学数学》全书重点总结', level=0)
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Mathematics for Computer Science — MIT')
run.italic = True
run.font.size = Pt(13)

doc.add_paragraph()
add_body(
    '本总结涵盖全书五个部分、共二十一章的核心内容，逐一说明每章的目的、'
    '训练的核心能力、以及必须掌握的关键知识点。适合作为复习大纲和查漏补缺的参考。'
)

# ============= INTRODUCTION =============
add_title('全书总览', level=1)
add_body(
    'MIT 的《计算机科学数学》并非传统的「应用数学课本」，而是一门训练计算机科学家'
    '数学思维方法的教材。全书围绕五大模块展开，每一个模块训练一种核心的数学能力：'
)
add_bullet('第一部分 · 证明（Proofs）：学会如何用数学语言精确地「说理」，建立严密的逻辑推理能力。')
add_bullet('第二部分 · 结构（Structures）：深入理解计算机科学中最基础的离散结构——数论、图论与网络。')
add_bullet('第三部分 · 计数（Counting）：掌握如何系统地「数数」，处理组合爆炸问题，分析算法复杂度。')
add_bullet('第四部分 · 概率（Probability）：在不确定性的世界中，为随机算法、数据分析和机器学习准备概率论工具。')
add_bullet('第五部分 · 递推（Recurrences）：学会处理递归算法和递推关系，是分析分治、动态规划等算法的基础。')

# ============= PART I =============
add_title('第一部分：证明（Proofs）—— 第 1～7 章', level=1)
add_body(
    '目的：培养「数学编程语言」的使用能力。学会阅读、理解并书写严谨的数学证明，'
    '这是计算机科学中几乎所有理论课程（算法分析、形式化方法、密码学等）的前置能力。'
)
add_body(
    '训练的核心能力：命题逻辑与谓词逻辑的运用、归纳法思维、构造性证明与反证法的灵活切换，'
    '以及将「直觉描述」转化为「形式化陈述」的能力。'
)

# Ch1
add_title('第 1 章：什么是证明（What is a Proof）', level=2)
add_body('目的：本章是整本书的起点，旨在从零开始建立「数学证明」的概念。')
add_body('训练能力：从直觉推理到形式化证明的转型能力；识别「看似合理但实际错误」的论证。')
add_bullet('命题（Proposition）：一个或真或假的陈述。')
add_bullet('谓词（Predicate）：依赖于变量、需要变量被具体化后才具有真值的陈述。')
add_bullet('公理化方法：从一组公理（axioms）出发，通过逻辑推导得到定理（theorems）。')
add_bullet('证明的结构：假设 -> 逻辑推导 -> 结论。')
add_bullet('反例（Counterexample）：用来证明全称命题为假的简单方法。')
add_bullet('常见错误：循环论证、偷换概念、以偏概全。')

# Ch2
add_title('第 2 章：良序原理（The Well Ordering Principle）', level=2)
add_body('目的：引入最基本但极其强大的证明方法——良序原理，并展示它与归纳法的等价性。')
add_body('训练能力：用「最小反例」法构造简洁证明的能力；理解自然数基本性质。')
add_bullet('良序原理：任何非空的非负整数集合都有一个最小元素。')
add_bullet('良序原理 <-> 数学归纳法：二者在逻辑上等价，可以互相推导。')
add_bullet('通过「假设存在反例，取最小的反例，推出矛盾」来证明定理。')
add_bullet('经典例证：sqrt(2) 是无理数（使用良序原理的经典证明）。')
add_bullet('应用场景：数论中的整除问题、分数化简的唯一性。')

# Ch3
add_title('第 3 章：逻辑公式（Logical Formulas）', level=2)
add_body('目的：提供一套完整的形式化逻辑语言，让读者能够精确地表达和操作数学陈述。')
add_body('训练能力：命题逻辑和简单谓词逻辑的符号运算能力；真值表分析；逻辑等价变形。')
add_bullet('逻辑连接词：NOT、AND、OR、IMPLIES、IFF。')
add_bullet('蕴含（IMPLIES）的真值表与日常语言的区别——假前提蕴含一切。')
add_bullet('等价式（Equivalence）与永真式（Validity）：区分语法等价和语义永真。')
add_bullet('SAT（可满足性）问题：是否存在一种赋值使公式为真——计算机验证的前沿问题。')
add_bullet('量词：全称量词 A 和存在量词 E；量词交换带来的含义变化。')
add_bullet('德摩根律：NOT(Ax.P(x)) = Ex.NOT(P(x))；NOT(Ex.P(x)) = Ax.NOT(P(x))。')

# Ch4
add_title('第 4 章：数学数据类型（Mathematical Data Types）', level=2)
add_body('目的：将编程语言中「数据类型」的概念用数学语言严格定义，在代码与数学之间建立一一映射。')
add_body('训练能力：用集合论定义复杂数据结构的能力；理解类型与不变量的关系。')
add_bullet('集合（Sets）：元素的定义、空集、子集、幂集、并/交/差/补集。')
add_bullet('有序对（Pairs）与笛卡尔积（Cartesian Product）：二元组的严格定义。')
add_bullet('关系（Relations）：定义域、值域、逆关系；关系的复合。')
add_bullet('函数（Functions）：是一种特殊的关系——每个输入有且仅有一个输出；单射/满射/双射。')
add_bullet('序列（Sequences）：函数视角下的序列定义（定义域为自然数）。')
add_bullet('这些定义与编程中 tuple、set、map、list 等类型的一一对应关系。')

# Ch5
add_title('第 5 章：归纳法（Induction）', level=2)
add_body('目的：全面掌握数学归纳法（普通归纳法和强归纳法），这是分析递归算法、证明算法正确性的第一工具。')
add_body('训练能力：归纳假设的建立与使用；将「从 n 到 n+1」的逻辑链条严密化。')
add_bullet('普通归纳法：证明 P(0) 为真；假设 P(n) 为真，证明 P(n+1) 为真。')
add_bullet('强归纳法：假设对所有 k < n，P(k) 为真，证明 P(n) 为真。')
add_bullet('归纳法为什么有效：根植于良序原理。')
add_bullet('构造性归纳：在归纳步骤中「构造」出更大规模的解。')
add_bullet('经典例题：马棋覆盖问题、奇偶性推理、斐波那契数列的各种恒等式证明。')
add_bullet('常见错误：归纳假设太弱（需要加强归纳命题）、错误使用 P(n)->P(n+1) 的证明方向。')

# Ch6
add_title('第 6 章：递归数据类型（Recursive Data Types）', level=2)
add_body('目的：将归纳法推广到数据结构上——用递归的方式定义数据类型，并用结构归纳法证明其性质。')
add_body('训练能力：用递归定义构建复杂结构（如树、表达式）的能力；结构归纳法与编程中递归函数的紧密联系。')
add_bullet('递归定义（Recursive Definition）：基例（Base Case）+ 构造步骤（Constructor Step）。')
add_bullet('结构归纳法（Structural Induction）：对递归定义的结构按「构造层次」进行归纳证明。')
add_bullet('字符串（Strings）：最基本的递归数据类型；Kleene 星号（Sigma*）的定义。')
add_bullet('二叉树（Binary Trees）与表达式树（Expression Trees）。')
add_bullet('匹配括号（Matched Brackets）的递归定义与性质证明——编译器与解析器的数学基础。')
add_bullet('无限集合由有限规则生成的哲学含义。')

# Ch7
add_title('第 7 章：无限集（Infinite Sets）', level=2)
add_body('目的：直面「无限」带来的直觉挑战，建立可数/不可数集的概念，为计算理论和复杂度理论打下基础。')
add_body('训练能力：理解「大小比较」在无限集上的扩展；对角线法的灵活运用；识别是否存在不能被计算的问题。')
add_bullet('有限集 vs 无限集：无限集可以与自身真子集一一对应（Dedekind 定义）。')
add_bullet('可数集（Countable）：能与 N 建立双射的集合；Z、Q 都是可数的。')
add_bullet('不可数集（Uncountable）：对角线论证法——R 不可数，{0,1} 的无穷序列集合不可数。')
add_bullet('图灵机与不可判定问题：程序（有限的二进制串）是可数的，但 N->N 的函数却是不可数的——因此存在「不可计算」的函数。')
add_bullet('Cantor-Schroeder-Bernstein 定理：若 A <= B 且 B <= A，则 A = B（等势）。')

# ============= PART II =============
add_title('第二部分：结构（Structures）—— 第 8～12 章', level=1)
add_body('目的：围绕数论和图论两大离散结构，掌握它们在密码学、网络、算法等计算机科学核心领域的应用。')
add_body('训练的核心能力：模运算的熟练使用；图论建模能力（将真实问题抽象为图）；理解图的可平面性与网络优化。')

# Ch8
add_title('第 8 章：数论（Number Theory）', level=2)
add_body('目的：数论是 RSA 加密等现代密码学的数学基础。本章专注于建立模运算体系并理解基本数论定理。')
add_body('训练能力：模运算的直觉化运算、最大公约数算法（欧几里得算法）及其扩展版本。')
add_bullet('整除：a | b 的定义；带余除法：a = qb + r。')
add_bullet('最大公约数（GCD）与欧几里得算法：辗转相除法，O(log n) 复杂度。')
add_bullet('扩展欧几里得算法：求解 ax + by = gcd(a,b) 的整数解 x, y——解线性丢番图方程的基础。')
add_bullet('模运算系统：Z_n 上的加法与乘法；乘法逆元存在的条件（gcd(a,n)=1）。')
add_bullet('费马小定理（Fermat Little Theorem）：若 p 为素数，则 a^{p-1} = 1 (mod p)。')
add_bullet('欧拉定理与欧拉 phi 函数；RSA 加密的数学原理（大素数乘积的因数分解困难性）。')
add_bullet('中国剩余定理：解同余方程组的系统方法。')

# Ch9
add_title('第 9 章：有向图（Directed Graphs）', level=2)
add_body('目的：有向图是网络流、调度问题、状态机、依赖关系的核心建模工具。')
add_body('训练能力：用有向图建模「因果/流向/依赖」关系的抽象能力；掌握路径、环和连通性的精确定义。')
add_bullet('有向图的形式化定义：顶点集 V、有向边集 E _in_ V x V。')
add_bullet('路径（Walk）、迹（Trail）、轨（Path）的精细区分。')
add_bullet('有向无环图（DAG）与拓扑排序：任务调度的数学模型。')
add_bullet('强连通分量：任意两点可互达的子图；Kosaraju / Tarjan 算法思路。')
add_bullet('邻接矩阵与路径计数：A^k 中 (i,j) 元素 = 从 i 到 j 的长度为 k 的路径数。')
add_bullet('可达性与传递闭包：Warshall 算法的数学基础。')

# Ch10
add_title('第 10 章：通信网络（Communication Networks）', level=2)
add_body('目的：以图论为工具，研究网络中数据传输的极限能力——路由与容错。')
add_body('训练能力：理解网络的瓶颈（割集）；最大流问题的建模与直觉培养。')
add_bullet('网络：有向图 + 容量函数；源点与汇点。')
add_bullet('路由、延迟与吞吐量：带宽限制下的最优传输策略。')
add_bullet('割（Cut）与割容量：任何割的容量都限制了从源到汇的最大流。')
add_bullet('最大流-最小割定理（Max-Flow Min-Cut Theorem）的直觉与含义。')
add_bullet('应用：二分图匹配、网络可靠性分析、并行计算中的通信优化。')

# Ch11
add_title('第 11 章：简单图（Simple Graphs）', level=2)
add_body('目的：系统建立无向简单图的理论框架，为后续平面图和社会网络分析奠定基础。')
add_body('训练能力：将社交网络、分子结构、电路布局等抽象为简单图并分析性质。')
add_bullet('简单图：无自环、无重边的无向图。')
add_bullet('度（Degree）、握手引理：所有顶点的度之和 = 2|E|。')
add_bullet('子图与导出子图：给定顶点子集或边子集生成的局部结构。')
add_bullet('同构（Isomorphism）：图的结构等价性——两个图是否「本质上相同」。')
add_bullet('二分图（Bipartite Graph）的判定条件：图不含奇环。')
add_bullet('匹配（Matching）与霍尔婚姻定理（Hall Marriage Theorem）的直观解释。')
add_bullet('着色（Coloring）：色数 chi(G) 的定义；贪心着色算法的最坏情况分析。')
add_bullet('欧拉回路与汉密尔顿回路的存在性条件。')

# Ch12
add_title('第 12 章：平面图（Planar Graphs）', level=2)
add_body('目的：研究可以画在平面上且没有边交叉的图——芯片布线、地图绘制的理论基础。')
add_body('训练能力：读懂并运用欧拉公式分析平面结构的约束；训练「空间想象力」与「图的可视化」能力。')
add_bullet('平面嵌入（Planar Embedding）：边不交叉的平面画法。')
add_bullet('欧拉公式：v - e + f = 2（对连通平面图，其中 f 为面数）。')
add_bullet('边上限：任何简单平面图满足 e <= 3v - 6。')
add_bullet('K5 与 K33 的非平面性：Kuratowski 定理——所有非平面图都包含 K5 或 K33 的细分。')
add_bullet('四色定理（Four-Color Theorem）陈述：任何平面图只需 4 种颜色即可对面着色。')
add_bullet('应用：PCB 布线、地图绘制、几何数据结构。')

# ============= PART III =============
add_title('第三部分：计数（Counting）—— 第 13～15 章', level=1)
add_body(
    '目的：系统学习组合数学中的计数方法，从基础的加/乘法原理到高级的生成函数方法。'
)
add_body(
    '训练的核心能力：将复杂计数问题分解为标准组合对象（排列、组合、分割）的能力；'
    '渐近分析的能力（理解大 O 的真正含义）；生成函数的代数操作能力。'
)

# Ch13
add_title('第 13 章：和与渐近（Sums and Asymptotics）', level=2)
add_body('目的：为算法复杂度分析提供数学工具——求闭合形式和渐近上/下界。')
add_body('训练能力：将循环累加转化为封闭公式；用大 O / 大 Omega / 大 Theta 精确描述函数增长。')
add_bullet('常见和的闭合公式：等差数列和、几何级数和、平方和、立方和。')
add_bullet('裂项法（Telescoping）：通过巧妙的项抵消将长和化简。')
add_bullet('扰动法（Perturbation Method）：给和起个名字 S，加减一项后构造关于 S 的方程。')
add_bullet('积分近似的思想：用积分近似求和，为连续-离散过渡做准备。')
add_bullet('大 O 记号的严格定义：f = O(g) 意味着存在 c, n0 使得对所有 n>=n0, f(n)<=c*g(n)。')
add_bullet('大 Omega、大 Theta 的定义与对应含义（下界与紧界）。')
add_bullet('常见渐近增长排序：O(1) < O(log n) < O(n) < O(n log n) < O(n^2) < O(2^n) < O(n!)。')
add_bullet('斯特林公式（Stirling Formula）：n! 的渐近近似——Theta(n log n) 到 O(n^2) 的分界线。')

# Ch14
add_title('第 14 章：计数法则（Cardinality Rules）', level=2)
add_body('目的：掌握系统计数的四大基本法则及组合对象的标准定义。这是排列组合的核心方法论。')
add_body('训练能力：将现实问题映射到标准组合模型（排列/组合/多重集/分割）的能力；用规则组合解决问题。')
add_bullet('乘法法则（Product Rule）：|A x B| = |A| * |B|。')
add_bullet('加法法则（Sum Rule）：不相交集合的并 = 大小之和。')
add_bullet('双射法则（Bijection Rule）：若 A 与 B 之间存在双射，则 |A| = |B|——「间接计数」的基础。')
add_bullet('补集法则（Complement Rule）：总数 - 不满足条件的个数。')
add_bullet('排列（Permutations）：n 个不同元素的有序排列数为 n!；k-排列的计算。')
add_bullet('组合（Combinations）：C(n,k) = n! / (k!(n-k)!)；二项式系数 (n choose k) 的含义。')
add_bullet('有放回排列与组合：n^k 与 C(n+k-1, k)（星棒法推导）。')
add_bullet('多重集组合（Multisets）与整数分割（Integer Partitions）的区分。')
add_bullet('容斥原理（Inclusion-Exclusion Principle）：|A U B| = |A|+|B|-|A n B|，推广到 n 个集合的通式。')
add_bullet('鸽巢原理（Pigeonhole Principle）：若 n+1 个物品放入 n 个盒子，必有一盒有 >=2 个物品——强有力的存在性论证。')

# Ch15
add_title('第 15 章：生成函数（Generating Functions）', level=2)
add_body('目的：引入代数方法来处理计数问题——将无限序列编码为形式幂级数，将组合问题转化为代数操作。')
add_body('训练能力：数列 <-> 幂级数的转化能力；用生成函数求解线性递推的机械化方法；组合类操作的生成函数对应。')
add_bullet('形式幂级数（Formal Power Series）：A(x) = Sigma a_n x^n；不关心收敛性，只看系数。')
add_bullet('生成函数的基本运算：加法（对应序列逐项加）、乘法（对应卷积）、导数、积分。')
add_bullet('卷积的含义：C(x) = A(x)*B(x) => c_n = Sigma a_k b_{n-k}，将两个组合对象「合并」的数学描述。')
add_bullet('常见序列的生成函数：(1-x)^{-1} = Sigma x^n；(1-x)^{-k} 对应 C(n+k-1, n) 的序列。')
add_bullet('用部分分式分解求解线性常系数递推（例如斐波那契数列的闭合公式推导）。')
add_bullet('指数生成函数（EGF）在标号结构计数中的应用——区分「带标签」和「不带标签」的组合。')

# ============= PART IV =============
add_title('第四部分：概率（Probability）—— 第 16～20 章', level=1)
add_body(
    '目的：建立适合计算机科学的概率论体系——从离散样本空间出发，直到随机变量的高级应用。'
)
add_body(
    '训练的核心能力：概率建模能力（用概率空间精确描述问题）；条件概率的深层理解与贝叶斯推理；'
    '随机算法的分析与设计能力；集中不等式（如切比雪夫/切尔诺夫界）的运用。'
)

# Ch16
add_title('第 16 章：事件与概率（Events and Probability）', level=2)
add_body('目的：从零建立概率论的形式化基础——样本空间、事件空间、概率函数的三元组模型。')
add_body('训练能力：概率模型的严格建立；直觉概率推断的形式化验证。')
add_bullet('概率空间 (S, F, P)：样本空间 S、可测事件族 F、概率测度 P。')
add_bullet('柯尔莫哥洛夫公理：非负性、规范性、可数可加性。')
add_bullet('离散均匀概率模型：每个基本结果等可能 => P(E) = |E| / |S|。')
add_bullet('补事件、交事件、并事件的概率计算规则。')
add_bullet('条件概率的定义：P(A|B) = P(A n B) / P(B)。')
add_bullet('全概率公式（Law of Total Probability）：用样本空间的一个划分加权平均计算概率。')
add_bullet('常见谬误：赌徒谬误（认为独立事件的概率会「补偿」）、检察官谬误（混淆 P(A|B) 和 P(B|A)）。')

# Ch17
add_title('第 17 章：条件概率（Conditional Probability）', level=2)
add_body('目的：深入理解条件概率的微妙之处，并掌握贝叶斯定理作为推理工具。')
add_body('训练能力：正确使用贝叶斯定理进行反向推理；区分因果关系和统计关系；用树形图分解复杂条件概率。')
add_bullet('贝叶斯定理：P(B|A) = P(A|B)*P(B) / P(A)——用「果」推理「因」的数学基础。')
add_bullet('先验概率、似然度、后验概率：贝叶斯更新的三步过程。')
add_bullet('独立性：P(A n B) = P(A)*P(B)；条件独立（Given C 的独立）不等同于无条件独立。')
add_bullet('多重事件独立：成对独立 != 多重相互独立。')
add_bullet('蒙提霍尔问题（Monty Hall Problem）的详细分析：为什么换门获胜概率是 2/3？')
add_bullet('医疗检测中的 Bayes 悖论：即使检测准确率 99%，罕见病的阳性预测值仍可能很低。')

# Ch18
add_title('第 18 章：随机变量（Random Variables）', level=2)
add_body('目的：从「事件」的语言过渡到「变量」的数学工具——随机变量、期望值、分布。')
add_body('训练能力：定义和使用随机变量来分析复杂随机系统；期望的线性性质（最被低估的强大工具）。')
add_bullet('随机变量（Random Variable）：定义在样本空间上的实值函数。')
add_bullet('概率质量函数（PMF）与累积分布函数（CDF）。')
add_bullet('期望值（Expectation）：概率加权平均 E[X] = Sigma x*P(X=x)。')
add_bullet('期望的线性性质：E[X+Y] = E[X] + E[Y]——不需要 X 和 Y 独立！')
add_bullet('方差（Variance）：Var[X] = E[(X-mu)^2] = E[X^2] - (E[X])^2。')
add_bullet('重要离散分布：伯努利分布、二项分布、几何分布、泊松分布及其期望/方差。')
add_bullet('指示变量（Indicator Variable）技巧：用 0/1 随机变量累加求期望，将复杂期望拆分。')
add_bullet('随机变量独立性：X 和 Y 独立当且仅当对所有 a,b, P(X=a,Y=b) = P(X=a)*P(Y=b)。')

# Ch19
add_title('第 19 章：偏离均值（Deviation from the Mean）', level=2)
add_body('目的：当 n 很大时，随机变量以怎样的程度集中在期望附近？引入集中不等式。')
add_body('训练能力：用马尔可夫/切比雪夫/切尔诺夫界对尾部分布做估算；为随机算法提供概率保证。')
add_bullet('马尔可夫不等式（Markov Inequality）：对非负随机变量 X，P(X>=a) <= E[X]/a。')
add_bullet('切比雪夫不等式（Chebyshev Inequality）：P(|X-mu|>=a) <= Var[X] / a^2。')
add_bullet('切尔诺夫界（Chernoff Bound）：对独立指示变量之和的极端集中性——指数级别的尾部分布衰减。')
add_bullet('大数定律（Law of Large Numbers）：样本均值依概率收敛于期望。')
add_bullet('中心极限定理（Central Limit Theorem）的思想：独立同分布随机变量之和近似正态分布。')
add_bullet('应用：随机负载均衡、随机采样可靠性分析、Polling 中的置信区间。')

# Ch20
add_title('第 20 章：随机游走（Random Walks）', level=2)
add_body('目的：以随机游走为具体模型，综合运用前面所学的概率、期望、递推、集中不等式等工具。')
add_body('训练能力：复杂随机过程的建模与分析；二维/高维随机游走的极限性质；用生成函数与递推分析随机系统。')
add_bullet('一维简单随机游走（Simple Random Walk）：每一步 +-1，等概率。')
add_bullet('反射原理（Reflection Principle）与首达时间（First Passage Time）的分析。')
add_bullet('赌徒破产问题（Gambler Ruin）：带吸收壁的随机游走，计算破产概率。')
add_bullet(
    'Polya 经典定理：一维和二维随机游走是常返的（必然会回到原点无穷多次），'
    '但三维及以上是非能返的（有正概率永远不回来）！'
)
add_bullet(
    '随机游走在计算机科学中的应用：PageRank 的随机冲浪者模型、蒙特卡洛方法、'
    '计算机网络中的随机退避算法。'
)

# ============= PART V =============
add_title('第五部分：递推（Recurrences）—— 第 21 章', level=1)
add_body(
    '目的：系统学习递推关系的建立与求解——这是分析递归算法（分治、动态规划）的数学语言。'
)
add_body(
    '训练的核心能力：从递归算法/递归数据结构中提取递推方程；'
    '闭合解的多种求解方法（猜+验证、特征方程、生成函数、递归树/主干定理）；'
    '用大 O 描述递推解的增长阶。'
)

# Ch21
add_title('第 21 章：递推（Recurrences）', level=2)
add_body('目的：递推是描述「一个问题如何由规模更小的同类问题的解构造出来」的语言。')
add_body('训练能力：递推建模（将自然语言/代码中的递归译成递推方程）；多种求解方法的灵活切换。')
add_bullet('递推方程的建立：基例 + 递推关系（如 T(n) = 2T(n/2) + n）。')
add_bullet('猜证法（Guess & Verify）：猜测一个闭合形式，用归纳法严格证明——猜的时候可以借助数值算例。')
add_bullet('特征方程法（Characteristic Equation）：解常系数线性齐次递推——a_n = c1*a_{n-1}+...+ck*a_{n-k}。')
add_bullet('非齐次递推的待定系数法：解 = 齐次通解 + 特解。')
add_bullet('生成函数法（Generating Function Method）：将递推转化为关于 G(x)=Sigma a_n x^n 的方程，求解后展开。')
add_bullet('分治递归与 Akra-Bazzi 定理：形如 T(n)=Sigma a_i T(n/b_i)+f(n) 的通用渐近解。')
add_bullet(
    '常用递推复杂度速查：二分搜索 T(n)=T(n/2)+O(1)->O(log n)；'
    '归并排序 T(n)=2T(n/2)+O(n)->O(n log n)；'
    '汉诺塔 T(n)=2T(n-1)+1->O(2^n)。'
)
add_bullet('递推与微分方程的类比：离散 vs 连续的「导数」对应差分操作。')

# ============= CLOSING SUMMARY =============
add_title('总结：五大能力的贯通训练', level=1)
add_body('本书五大部分由浅入深，形成了一条「数学思维工具链」：')
add_bullet('证明能力（Part I）：让你能够严谨地说清楚「为什么算法正确」、「为什么不能做得更好」。')
add_bullet('结构分析能力（Part II）：让你能够将复杂系统抽象为离散数学结构（图、模运算系统）并做推理。')
add_bullet('计数与渐近分析能力（Part III）：让你能够估计算法消耗的资源（时间、空间），并比较不同方案的效率。')
add_bullet('概率推理能力（Part IV）：让你能够在不确定性中做出最优决策，设计并分析随机算法。')
add_bullet('递推求解能力（Part V）：让你能够分析递归算法的时间复杂度，把「递归代码」翻译为「闭合复杂度公式」。')
add_body('')
add_body(
    '以上五者并非孤立——证明能力使你能够验证任何数学陈述；结构分析使你有东西可证；'
    '计数概率使你能做资源预测；递推则是将算法的递归结构数学化。'
    '建议在学完每一章后，尝试用该章的方法去分析一个你熟悉的算法或数据结构（如快速排序、哈希表、'
    'DFS/BFS、二叉搜索树等），以此巩固「从数学到代码」的桥梁。'
)

# Save
output_path = r'计算机离散数学_MIT_全书重点总结.docx'
doc.save(output_path)
print('Document saved successfully to: ' + output_path)