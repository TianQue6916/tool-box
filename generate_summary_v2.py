
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

doc = Document()

# ============ Style Setup ============
style = doc.styles['Normal']
font = style.font
font.name = '微软雅黑'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
style.paragraph_format.line_spacing = 1.3
style.paragraph_format.space_after = Pt(4)

for i in range(1, 5):
    hs = doc.styles[f'Heading {i}']
    hs.font.name = '微软雅黑'
    hs.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    hs.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)

# ============ Helper Functions ============
def add_bold_para(doc, bold_text, normal_text):
    p = doc.add_paragraph()
    run_b = p.add_run(bold_text)
    run_b.bold = True
    run_b.font.size = Pt(11)
    run_n = p.add_run(normal_text)
    run_n.font.size = Pt(11)
    return p

def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.size = Pt(11)
        run_n = p.add_run(text)
        run_n.font.size = Pt(11)
    else:
        p.runs[0].font.size = Pt(11) if p.runs else None
        p.text = text
    return p

# ============ TITLE ============
title = doc.add_heading('《计算机科学中的数学》全书深度总结', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in title.runs:
    run.font.size = Pt(26)

doc.add_paragraph('基于MIT教材 Mathematics for Computer Science (Meyer & Lehman)', style='Intense Quote')
doc.add_paragraph('')

# ============ 全书总论 ============
doc.add_heading('全书总论', level=1)

doc.add_paragraph(
    '本书是MIT 6.042J课程的核心教材，是MIT计算机科学专业本科生的必修数学课。'
    '它不是一本传统的离散数学教材——它的核心特色在于"以证明为纲"：全书围绕如何严格地建立数学真理展开，'
    '每一章都在训练你一种特定的数学思维能力。与普通数学教材不同，本书始终将数学与计算机科学的实际需求紧密结合，'
    '从程序正确性验证到密码学，从网络设计到随机算法分析，让你看到每一个数学工具"为什么需要"以及"怎么用"。'
)

doc.add_heading('全书五大板块的逻辑关系', level=2)
doc.add_paragraph(
    '第一部分"证明"是地基——教你如何严格地推理；'
    '第二部分"结构"是语言——教你用数学对象（数、图、关系）来建模计算问题；'
    '第三部分"计数"是工具——教你精确地数"有多少种可能"；'
    '第四部分"概率"是升级——在不确定的世界中做推理；'
    '第五部分"递推"是综合——用递推关系把前面所有工具串联起来分析算法。'
    '五大部分环环相扣：没有证明能力就无法严格论证图论定理，没有计数能力就无法计算概率，没有概率基础就无法分析随机算法。'
)

doc.add_heading('核心哲学', level=2)
doc.add_paragraph(
    '本书的一个贯穿始终的哲学是：直觉不可靠，证明才可靠。书中反复用反例说明"看起来显然"的命题可能是错的，'
    '而"看起来不对"的命题可能是对的。例如，欧拉曾猜想对于n>2，n次幂的和至少需要n个n次幂才能凑成另一个n次幂，'
    '这个猜想看起来很自然，但被反例推翻了。而"每个地图只需4种颜色"看起来太强了，却是对的。'
    '这种对直觉的警惕和对证明的执着，是本书最想培养的思维习惯。'
)

# ================================================================
# PART I: PROOFS
# ================================================================
doc.add_heading('第一部分：证明（Proofs）', level=1)
doc.add_paragraph(
    '本部分共7章（第1-7章），是全书的基础。计算机科学区别于其他工程学科的一个核心特征是：'
    '我们能够（也必须）证明程序和系统的正确性，而不仅仅是测试。测试只能发现bug，证明才能保证没有bug。'
    'Intel在1990年代的浮点除法bug造成了数亿美元损失，这推动了芯片验证中形式化证明的广泛应用。'
    '本部分的目标就是让你掌握证明的"语法"和"修辞"——如何构造一个合法的证明，如何让它清晰有力。'
)

# ---- Chapter 1 ----
doc.add_heading('第1章 什么是证明（What is a Proof）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '这是全书的起点。本章要回答一个根本问题：在数学中，"真理"是如何确立的？'
    '不同领域有不同的"真理标准"——法庭靠证据和陪审团，科学靠实验，商业靠权威，统计靠显著性。'
    '数学有自己独特的标准：从公理出发，通过逻辑推导链到达结论。本章让你理解这个标准，并掌握最基本的证明模式。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '识别命题（Proposition）——知道什么样的语句可以被证明或证伪', '命题意识：')
add_bullet(doc, '理解公理→逻辑推导→定理的证明链条', '公理化思维：')
add_bullet(doc, '掌握证明蕴含、当且仅当、反证法、分类讨论的标准模板', '证明构造能力：')
add_bullet(doc, '区分"看起来对"和"严格对"，警惕直觉陷阱', '反直觉意识：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '1. 命题（Proposition）', '：一个非真即假的判断语句。例如"2+3=5"是真命题，"2+3=6"是假命题。')
add_bold_para(doc, '2. 谓词（Predicate）', '：真值依赖于变量的命题。如P(n)="n是完全平方数"，P(4)为真，P(5)为假。谓词是后续逻辑公式和量词的基础。')
add_bold_para(doc, '3. 公理（Axiom）', '：不加证明即接受为真的基本假设。全书采用"高中数学常识"作为公理基础，而非最底层的ZFC集合论公理——因为用ZFC证明2+2=4需要超过20000步！')
add_bold_para(doc, '4. 逻辑推导规则（Inference Rules）', '：最基本的是假言推理（Modus Ponens）：从P和P→Q推出Q。推导规则必须是可靠的（sound），即前提为真时结论必为真。')
add_bold_para(doc, '5. 定理、引理、推论', '：定理是重要的真命题，引理是为证明定理服务的小命题，推论是从定理几步就能推出的结论。但界限并非绝对——有些引理比定理还重要。')

doc.add_paragraph('')
add_bold_para(doc, '证明蕴含（Implication）的两种方法：', '')
add_bullet(doc, '假设P为真，推导Q为真。这是最直接的方法。', '方法一（直接证明）：')
add_bullet(doc, '证明逆否命题¬Q→¬P，因为P→Q等价于¬Q→¬P。例如证明"若r是无理数则√r也是无理数"，直接证很绕，但证逆否"若√r是有理数则r是有理数"就非常简洁。', '方法二（逆否证明）：')

add_bold_para(doc, '证明当且仅当（If and Only If）的两种方法：', '')
add_bullet(doc, '分别证明P→Q和Q→P', '方法一：双向证明')
add_bullet(doc, '构造P≡R₁≡R₂≡...≡Q的等价链', '方法二：等价链')

add_bold_para(doc, '反证法（Proof by Contradiction）：', '假设命题为假，推出矛盾。本章通过经典例子"√2是无理数"来展示反证法的威力。')
add_bold_para(doc, '分类讨论（Proof by Cases）：', '将证明分成若干情况分别证明。经典例子：6个人中必有3人互相认识或3人互相不认识（拉姆齐理论的入门例子）。')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '本章表面上是教你"证明是什么"，但更深层的是在培养一种思维习惯：不要相信直觉，要相信逻辑。'
    '书中特意列举了四色定理、费马大定理、哥德巴赫猜想等例子——四色定理看起来"不太可能对"却是对的，'
    '欧拉猜想看起来"显然对"却是错的，哥德巴赫猜想至今不知道对不对。这些例子都在说同一件事：'
    '在数学中，"看起来对"不等于"确实对"，只有证明才能给出确定的答案。'
)

# ---- Chapter 2 ----
doc.add_heading('第2章 良序原理（Well Ordering Principle）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '良序原理是全书第一个真正的证明工具。它看似简单——"每个非空自然数集都有最小元素"——'
    '但这个朴素的陈述蕴含着强大的证明力量。本章教你用良序原理来证明关于自然数的命题，'
    '同时为第5章的归纳法奠定理论基础（良序原理与归纳法等价）。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '假设存在反例→取最小反例→推出矛盾', '最小反例法：')
add_bullet(doc, '将"对所有n成立"的命题转化为"存在反例则矛盾"的论证', '无限到有限的归约：')
add_bullet(doc, '理解良序原理是归纳法的理论基础', '为归纳法铺路：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '良序原理（Well Ordering Principle）', '：每个非空自然数集合都有一个最小元素。这是关于自然数的一个基本性质，不需要证明——我们把它当作公理接受。')

add_bold_para(doc, '良序证明的标准模板：', '')
doc.add_paragraph('1. 用反证法：假设存在某个自然数n使命题P(n)为假')
doc.add_paragraph('2. 定义反例集合C = {n ∈ N | P(n)为假}')
doc.add_paragraph('3. 由良序原理，C有最小元素m')
doc.add_paragraph('4. 证明m不可能是最小反例（通常通过证明P(m)为真或构造更小反例）')
doc.add_paragraph('5. 得出矛盾，因此C为空集，命题对所有n成立')

add_bold_para(doc, '经典应用：', '证明求和公式 1+2+3+...+n = n(n+1)/2。假设存在使公式不成立的n，取最小反例m，则对m-1公式成立，由此推出对m也成立——矛盾。')

add_bold_para(doc, '良序原理的适用范围', '：良序原理只适用于自然数（非负整数）。对于整数集、有理数集等不存在良序性（例如负整数没有最小元），因此良序证明只能用于自然数上的命题。')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '良序原理的深层意义在于：它提供了一种将"无限验证"转化为"有限矛盾"的方法。'
    '要证明"对所有自然数n，P(n)成立"，你不可能逐个验证无限个n，'
    '但良序原理让你只需处理"最小反例"这一个情况。这种从无限到有限的归约，是数学推理的核心技巧之一。'
)

# ---- Chapter 3 ----
doc.add_heading('第3章 逻辑公式（Logical Formulas）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '将日常语言中的推理形式化、精确化。在计算机科学中，逻辑是硬件设计（布尔电路）、'
    '程序验证（霍尔逻辑）、数据库查询（SQL的数学基础）、人工智能（知识表示）的共同基础。'
    '本章让你掌握命题逻辑和谓词逻辑的精确语言，学会用符号来表达和验证推理。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '将自然语言推理转化为命题逻辑公式', '形式化表达能力：')
add_bullet(doc, '用真值表验证逻辑等价和推理有效性', '真值表分析能力：')
add_bullet(doc, '正确使用∀和∃量词，避免量词混淆的常见错误', '量词操作能力：')
add_bullet(doc, '理解SAT问题的意义——这是计算机科学中最重要的NP完全问题', '计算复杂性意识：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '命题逻辑的连接词：', '')
add_bullet(doc, '合取P∧Q（且）、析取P∨Q（或）、否定¬P（非）', '基本连接词：')
add_bullet(doc, 'P→Q等价于¬P∨Q——这是理解蕴含的关键', '蕴含P→Q：')
add_bullet(doc, 'P↔Q = (P→Q)∧(Q→P)', '双条件P↔Q：')

add_bold_para(doc, '逻辑等价与基本定律：', '')
add_bullet(doc, '¬(P∧Q) ≡ ¬P∨¬Q，¬(P∨Q) ≡ ¬P∧¬Q', '德摩根定律：')
add_bullet(doc, 'P→Q ≡ ¬Q→¬P（逆否命题）', '蕴含的逆否：')
add_bullet(doc, 'P∧(Q∨R) ≡ (P∧Q)∨(P∧R)', '分配律：')

add_bold_para(doc, '谓词逻辑与量词：', '')
add_bullet(doc, '∀x.P(x)表示"对所有x，P(x)为真"', '全称量词∀：')
add_bullet(doc, '∃x.P(x)表示"存在某个x使P(x)为真"', '存在量词∃：')
add_bullet(doc, '¬∀x.P(x) ≡ ∃x.¬P(x)，¬∃x.P(x) ≡ ∀x.¬P(x)——量词否定时翻转', '量词的否定：')
add_bullet(doc, '∀x∃y和∃y∀x含义完全不同——量词顺序不能随意交换', '量词顺序的重要性：')

add_bold_para(doc, '可满足性问题（SAT）：', '给定一个命题公式，是否存在一组真值赋值使公式为真？SAT是第一个被证明为NP完全的问题（Cook-Levin定理），'
    '是计算复杂性理论的核心。虽然本章不深入复杂性理论，但理解SAT是理解"P vs NP"问题的起点。')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '本章最重要的教训是：自然语言是模糊的，逻辑语言是精确的。'
    '"如果你努力学习，你就会通过考试"——这句话在日常语言中似乎很清楚，'
    '但在逻辑中，"努力学习→通过考试"在你不努力学习时自动为真（空真），这违反直觉但逻辑上正确。'
    '理解这种自然语言与形式逻辑之间的鸿沟，是避免推理错误的关键。'
)

# ---- Chapter 4 ----
doc.add_heading('第4章 数学数据类型（Mathematical Data Types）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '介绍离散数学的三大基本数据类型：集合、函数、关系。这些是所有高级数学结构和计算机科学中数据建模的"原子"。'
    '数据库是关系的集合，程序是函数，类型系统依赖于集合论。本章为后续所有结构性的讨论提供精确的数学语言。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '集合运算的熟练计算与严格证明', '集合论基础：')
add_bullet(doc, '判断和证明函数的单射性、满射性、双射性', '函数性质分析：')
add_bullet(doc, '识别等价关系、偏序关系，理解它们与划分、排序的联系', '关系分类能力：')
add_bullet(doc, '利用双射建立一一对应来计数', '双射计数法：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '集合（Sets）：', '')
add_bullet(doc, '子集、真子集、幂集P(A)（A的所有子集构成的集合，|P(A)|=2^|A|）', '')
add_bullet(doc, '并A∪B、交A∩B、差A-B、补Ā', '')
add_bullet(doc, '笛卡尔积A×B = {(a,b) | a∈A, b∈B}，|A×B| = |A|·|B|', '')
add_bullet(doc, '集合构造记号：{x ∈ A | P(x)}表示A中满足P的元素集合', '')

add_bold_para(doc, '函数（Functions）：', '')
add_bullet(doc, 'f: A→B是满射（surjection）若每个b∈B都有a∈A使f(a)=b', '满射：')
add_bullet(doc, 'f是单射（injection）若f(a₁)=f(a₂)则a₁=a₂', '单射：')
add_bullet(doc, 'f是双射（bijection）当且仅当既是单射又是满射', '双射：')
add_bullet(doc, '双射原理：若存在A到B的双射，则|A|=|B|——这是计数的基本工具', '双射计数：')

add_bold_para(doc, '二元关系（Binary Relations）：', '')
add_bullet(doc, '自反性：∀a, aRa', '')
add_bullet(doc, '对称性：aRb → bRa', '')
add_bullet(doc, '反对称性：aRb ∧ bRa → a=b', '')
add_bullet(doc, '传递性：aRb ∧ bRc → aRc', '')
add_bullet(doc, '等价关系 = 自反+对称+传递，等价关系与集合划分一一对应', '')
add_bullet(doc, '偏序关系 = 自反+反对称+传递，偏序集是后续讨论拓扑排序和格论的基础', '')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '本章看似在罗列定义，实际上在建立一种"类型意识"：面对一个数学对象，你要能快速判断它属于什么类型、具有什么性质。'
    '这种能力在后续章节中无处不在——证明图论定理时需要判断关系性质，分析算法时需要判断函数性质，计数时需要判断双射存在性。'
)

# ---- Chapter 5 ----
doc.add_heading('第5章 归纳法（Induction）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '归纳法是离散数学中最重要的证明方法，没有之一。如果说良序原理是归纳法的理论基础，'
    '那么归纳法就是它的实践形态。在计算机科学中，归纳法与递归是一体两面：'
    '每个递归程序的正确性都需要归纳法来证明，每个归纳定义都对应一个递归算法。'
    '本章系统讲解普通归纳法和强归纳法，并展示它们在程序验证中的应用。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '能独立完成"基础步+归纳步"的标准归纳证明', '归纳证明构造能力：')
add_bullet(doc, '能根据问题特点选择普通归纳法或强归纳法', '归纳策略选择：')
add_bullet(doc, '能识别归纳证明中的常见错误（如归纳假设使用不当）', '归纳陷阱识别：')
add_bullet(doc, '用归纳法证明递归程序和循环的不变式', '程序正确性证明：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '普通归纳法（Ordinary Induction）：', '')
doc.add_paragraph('证明对所有n≥0，P(n)成立：')
doc.add_paragraph('  基础步（Base Case）：证明P(0)')
doc.add_paragraph('  归纳步（Inductive Step）：假设P(n)（归纳假设），证明P(n+1）')

add_bold_para(doc, '强归纳法（Strong Induction）：', '')
doc.add_paragraph('归纳假设不再是"P(n)成立"，而是"P(0), P(1), ..., P(n)都成立"。')
doc.add_paragraph('强归纳法在证明关于整数的命题时特别有用，例如：任何大于1的整数都可以分解为素数的乘积。')
doc.add_paragraph('重要事实：普通归纳法、强归纳法、良序原理三者等价——任何一个都可以推出另外两个。')

add_bold_para(doc, '归纳证明的常见错误：', '')
add_bullet(doc, '归纳步中使用了比归纳假设更强的结论，但该结论并未被证明', '归纳假设使用不当：')
add_bullet(doc, '基础步遗漏n=0或n=1的情况，导致归纳链断裂', '基础步遗漏：')
add_bullet(doc, '经典反例："所有马的颜色相同"的"证明"——错误出在从n=1到n=2的归纳步', '经典反例：')

add_bold_para(doc, '归纳法与程序验证：', '')
doc.add_paragraph('循环不变式（Loop Invariant）是归纳法在程序验证中的直接应用：')
doc.add_paragraph('  基础步：循环开始前不变式成立')
doc.add_paragraph('  归纳步：如果循环开始时不变式成立，则循环体执行后不变式仍然成立')
doc.add_paragraph('  终止：循环结束时不变式给出程序的正确性')
doc.add_paragraph('这种方法被Intel等公司用于验证芯片设计的正确性。')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '归纳法的深层意义在于：它是一种处理无穷的有限方法。你不需要验证无限个情况，'
    '只需要验证一个基础情况和一个"传递规则"。这与程序中的循环和递归完美对应——'
    '循环体执行一次对应归纳步一次，循环的终止条件对应基础步。'
    '理解归纳法，就是理解了"如何用有限的推理把握无限的对象"。'
)

# ---- Chapter 6 ----
doc.add_heading('第6章 递归数据类型（Recursive Data Types）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '将归纳思想从自然数推广到更一般的递归结构。在计算机科学中，链表、树、表达式、'
    '字符串等都是递归定义的——它们由基础情况和构造规则定义。本章教你对这类结构进行归纳证明（结构归纳法），'
    '这是函数式编程和编译器设计中的核心推理工具。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '用"基础情况+构造规则"定义数据类型', '递归定义能力：')
add_bullet(doc, '对递归定义的结构进行归纳证明', '结构归纳法：')
add_bullet(doc, '理解递归数据类型与上下文无关文法的对应关系', '文法与递归的联系：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '递归定义的模式：', '')
doc.add_paragraph('  基础情况：定义最简单的元素')
doc.add_paragraph('  构造规则：从已有元素构造新元素的规则')
doc.add_paragraph('例如，匹配括号字符串的递归定义：')
doc.add_paragraph('  基础：空串ε是匹配的')
doc.add_paragraph('  规则1：若s是匹配的，则(s)是匹配的')
doc.add_paragraph('  规则2：若s和t是匹配的，则st是匹配的')

add_bold_para(doc, '结构归纳法（Structural Induction）：', '')
doc.add_paragraph('  基础步：证明命题对基础情况成立')
doc.add_paragraph('  归纳步：假设命题对已有的结构成立，证明通过构造规则得到的新结构也满足命题')
doc.add_paragraph('结构归纳法比自然数归纳法更一般——自然数归纳法是结构归纳法的特例。')

add_bold_para(doc, '递归函数与计算：', '')
doc.add_paragraph('每个递归定义的数据类型都自然对应一类递归函数。例如，字符串长度、列表求和、树的深度等。')
doc.add_paragraph('递归函数的计算可以通过结构归纳法来证明其正确性。')

doc.add_heading('▎本章的深层意图', level=3)
doc.add_paragraph(
    '本章的深层意图是让你理解：归纳不仅是一种证明技术，更是一种思维方式——'
    '从简单情况出发，通过规则逐步构建复杂结构。这与计算机科学中"自底向上"的设计哲学完全一致。'
    '函数式编程中的代数数据类型（ADT）、编译器中的抽象语法树（AST）、'
    '数据库中的递归查询，都是这种思维方式的体现。'
)

# ---- Chapter 7 ----
doc.add_heading('第7章 无限集合（Infinite Sets）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '从有限世界进入无限世界。本章回答一个深刻的问题：无穷也有大小之分吗？'
    '康托尔的回答是：是的。有些无穷比另一些无穷更大。这不仅是一个数学惊奇，'
    '更是理解计算理论中可判定性与可计算性的基础——停机问题的不可判定性本质上就是'
    '对角线论证法在计算领域的应用。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '判断集合是否可数，构造可数集的枚举', '可数性判断：')
add_bullet(doc, '掌握康托尔对角线论证法——这是数学中最优美的证明之一', '对角线论证法：')
add_bullet(doc, '理解"无穷也有等级"这一反直觉事实', '无穷层次意识：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '可数集（Countable Sets）：', '')
add_bullet(doc, '集合S是可数的当且仅当存在从N到S的双射（或满射）', '定义：')
add_bullet(doc, '自然数N、整数Z、有理数Q都是可数的', '经典例子：')
add_bullet(doc, '有理数可数的证明：将正有理数排列成无穷矩阵，沿对角线枚举', '有理数可数的证明：')

add_bold_para(doc, '不可数集与对角线论证法：', '')
add_bullet(doc, '实数集R是不可数的', '')
add_bullet(doc, '康托尔的证明：假设[0,1]中的实数可枚举为r₁,r₂,...，构造实数d使其第i位与rᵢ的第i位不同，则d不在枚举中——矛盾', '对角线论证：')
add_bullet(doc, '对角线论证法是数学中最重要的证明方法之一，在逻辑学（哥德尔不完备定理）和计算理论（停机问题）中都有核心应用', '深远影响：')

add_bold_para(doc, "幂集定理（Cantor's Theorem）：", '')
add_bullet(doc, '对任何集合A，|P(A)| > |A|', '')
add_bullet(doc, '证明同样使用对角线论证法', '')
add_bullet(doc, '推论：不存在"最大"的无穷——N, P(N), P(P(N)), ... 严格递增', '')

add_bold_para(doc, '与计算理论的联系：', '')
doc.add_paragraph(
    '停机问题（给定程序P和输入x，P(x)是否停机？）是不可判定的。'
    '其证明本质上就是对角线论证法在计算领域的应用：假设存在判定停机的程序H，'
    '构造程序D使得D(P)的行为与H(P,P)相反，则D(D)产生矛盾。'
    '理解了本章的对角线论证法，你就理解了为什么有些问题是计算机永远无法解决的。'
)

# ================================================================
# PART II: STRUCTURES
# ================================================================
doc.add_heading('第二部分：结构（Structures）', level=1)
doc.add_paragraph(
    '本部分共5章（第8-12章），介绍离散数学的核心结构：数论、有向图、通信网络、无向图和平面图。'
    '这些结构是建模实际计算问题的数学语言。本部分的训练目标是：'
    '让你学会将实际问题抽象为数学结构，然后用结构本身的性质来解决问题。'
)

# ---- Chapter 8 ----
doc.add_heading('第8章 数论（Number Theory）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '数论是"数学女王"，也是现代密码学的数学基础。本章从最基本的整除性出发，'
    '逐步构建模运算理论，最终达到RSA公钥加密系统——这是纯数学应用于计算机科学最成功的案例之一。'
    '通过本章，你将理解网上银行、HTTPS、数字签名背后的数学原理。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '模运算的计算与推理——这是密码学的基本语言', '模运算能力：')
add_bullet(doc, '理解欧几里得算法为何正确、为何高效', '算法设计与分析能力：')
add_bullet(doc, '理解RSA加密系统的每一步数学原理', '密码学理解力：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '整除性与GCD：', '')
add_bullet(doc, 'a | b 表示a整除b，即b = ka对某个整数k', '')
add_bullet(doc, 'gcd(a,b)：a和b的最大公约数', '')

add_bold_para(doc, '欧几里得算法（Euclidean Algorithm）：', '')
doc.add_paragraph('基于gcd(a,b) = gcd(b, a mod b)，递归计算直到b=0。')
doc.add_paragraph('时间复杂度：O(log(min(a,b)))——每两次递归调用，参数至少减半。')
doc.add_paragraph('扩展欧几里得算法：同时求出gcd(a,b)和满足gcd(a,b)=sa+tb的整数s,t（贝祖定理）。')

add_bold_para(doc, '模运算：', '')
add_bullet(doc, 'a ≡ b (mod m) 当且仅当 m | (a-b)', '同余的定义：')
add_bullet(doc, '同余关系是等价关系，模m的等价类称为剩余类', '等价关系：')
add_bullet(doc, "加法、减法、乘法在同余下保持：若a≡a'且b≡b'，则a+b≡a'+b'，ab≡a'b'", '运算性质：')
add_bullet(doc, '乘法逆元：a在模m下有逆元当且仅当gcd(a,m)=1', '逆元：')

add_bold_para(doc, '费马小定理与欧拉定理：', '')
add_bullet(doc, '费马小定理：若p是素数且gcd(a,p)=1，则a^(p-1) ≡ 1 (mod p)', '')
add_bullet(doc, '欧拉定理：若gcd(a,m)=1，则a^φ(m) ≡ 1 (mod m)，其中φ(m)是欧拉函数（1到m中与m互素的整数个数）', '')
add_bullet(doc, '费马小定理是欧拉定理的特例（p是素数时φ(p)=p-1）', '')

add_bold_para(doc, 'RSA加密系统：', '')
doc.add_paragraph('密钥生成：选两个大素数p,q，计算n=pq和φ(n)=(p-1)(q-1)，选e使gcd(e,φ(n))=1，计算d=e⁻¹ (mod φ(n))')
doc.add_paragraph('公钥：(e,n)，私钥：d')
doc.add_paragraph('加密：c = m^e (mod n)，解密：m = c^d (mod n)')
doc.add_paragraph('正确性保证：由欧拉定理保证m^ed ≡ m (mod n)')
doc.add_paragraph('安全性基础：已知n和e，求d需要知道φ(n)，而求φ(n)需要分解n=pq，大数分解目前没有高效算法')

# ---- Chapter 9 ----
doc.add_heading('第9章 有向图（Directed Graphs）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '有向图是建模状态机、依赖关系、网络流等计算问题的核心结构。'
    '本章建立图论的基本语言和证明方法，重点是有向无环图（DAG）和拓扑排序——'
    '这在任务调度、编译器设计、数据分析中都有直接应用。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '将实际问题（任务依赖、状态转移）抽象为有向图', '图论建模能力：')
add_bullet(doc, '判断DAG、进行拓扑排序', 'DAG与拓扑排序：')
add_bullet(doc, '分析路径、连通性、强连通分量', '图性质分析：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '有向图基本概念：', '')
add_bullet(doc, '有向图G = (V, E)，V是顶点集，E是有向边集', '')
add_bullet(doc, '有向路径：顶点序列v₀→v₁→...→vₖ', '')
add_bullet(doc, '有向环：起点和终点相同的有向路径', '')

add_bold_para(doc, '有向无环图（DAG）与拓扑排序：', '')
add_bullet(doc, 'DAG：不含有向环的有向图', '')
add_bullet(doc, '拓扑排序：将DAG的顶点排成线性序列，使每条边的起点在终点之前', '')
add_bullet(doc, 'DAG存在拓扑排序当且仅当图无环', '')
add_bullet(doc, '应用：课程先修关系、任务调度、Makefile依赖分析', '')

add_bold_para(doc, '强连通分量：', '')
add_bullet(doc, '顶点u和v强连通当且仅当u可达v且v可达u', '')
add_bullet(doc, '强连通分量是极大的强连通子图', '')
add_bullet(doc, '任何有向图都可以分解为强连通分量的DAG', '')

# ---- Chapter 10 ----
doc.add_heading('第10章 通信网络（Communication Networks）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '本章是图论的应用篇，展示数学如何指导网络工程设计。通过分析特定网络拓扑（二叉树、蝶形网络、Benes网络）的性能，'
    '让你理解工程设计中的数学权衡：交换节点数、网络直径、拥塞之间的制约关系。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '用图论语言描述和评估网络拓扑', '网络拓扑分析：')
add_bullet(doc, '理解设计中不同指标之间的权衡', '工程权衡意识：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '交换网络的性能指标：', '')
add_bullet(doc, '交换节点数：网络成本', '')
add_bullet(doc, '网络直径：最短路径的最大长度——决定延迟', '')
add_bullet(doc, '拥塞（Congestion）：最繁忙链路上的最大路径数——决定吞吐量', '')

add_bold_para(doc, '经典网络拓扑：', '')
add_bullet(doc, '完全二叉树：直径O(log N)，但根节点拥塞严重', '')
add_bullet(doc, '蝶形网络（Butterfly）：直径O(log N)，拥塞O(√N)', '')
add_bullet(doc, 'Benes网络：无阻塞网络，任意排列都可实现——但节点数更多', '')

# ---- Chapter 11 ----
doc.add_heading('第11章 简单图（Simple Graphs）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '无向图是建模社交网络、通信网络、分子结构等的基础工具。'
    '本章系统介绍无向图的性质和经典定理，包括握手定理、连通性、树、二部图、'
    '欧拉路径、图着色等。这些概念在算法设计和网络分析中无处不在。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '分析无向图的基本性质（度、连通性、着色等）', '图性质分析：')
add_bullet(doc, '识别二部图并利用二部图性质', '二部图应用：')
add_bullet(doc, '理解图着色问题的困难性和近似算法', '图着色能力：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '握手定理（Handshaking Lemma）：', '')
doc.add_paragraph('无向图中所有顶点度数之和等于2倍边数：Σ deg(v) = 2|E|。'
    '推论：度数为奇数的顶点个数必为偶数。这个定理虽然简单，但在图论证明中经常使用。')

add_bold_para(doc, '连通性与树：', '')
add_bullet(doc, '连通图：任意两顶点间有路径', '')
add_bullet(doc, '树：连通无环图，边数 = 顶点数 - 1', '')
add_bullet(doc, '生成树：连通图的极小连通子图', '')
add_bullet(doc, '最小生成树：边权之和最小的生成树（Kruskal/Prim算法）', '')

add_bold_para(doc, '二部图（Bipartite Graphs）：', '')
add_bullet(doc, '顶点集可分为两部分，所有边连接不同部分的顶点', '')
add_bullet(doc, '等价条件：图是二部图当且仅当不含奇数长度的环', '')
add_bullet(doc, '匹配问题：二部图中的最大匹配可用匈牙利算法求解', '')

add_bold_para(doc, '欧拉路径与欧拉回路：', '')
add_bullet(doc, '欧拉路径：经过每条边恰好一次的路径', '')
add_bullet(doc, '欧拉回路：起点和终点相同的欧拉路径', '')
add_bullet(doc, '欧拉回路存在当且仅当图连通且每个顶点度数为偶数', '')
add_bullet(doc, '欧拉路径存在当且仅当图连通且恰有0个或2个奇度顶点', '')

add_bold_para(doc, '图着色：', '')
add_bullet(doc, '色数χ(G)：给顶点着色使相邻顶点颜色不同的最少颜色数', '')
add_bullet(doc, '二部图的色数为2', '')
add_bullet(doc, '图着色是NP难问题——没有已知的多项式时间算法', '')

# ---- Chapter 12 ----
doc.add_heading('第12章 平面图（Planar Graphs）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '平面图理论在电路板设计、地图着色等领域有直接应用。'
    '本章通过欧拉公式和五色定理，展示图论中"结构决定性质"的深刻思想，'
    '以及如何用简单的公式推出强大的结论。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '判断图是否为平面图', '平面图判断：')
add_bullet(doc, '用欧拉公式进行图论推导', '欧拉公式应用：')
add_bullet(doc, '理解"不可能性"证明的方法论', '不可能性证明：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, "欧拉公式（Euler's Formula）：", '')
doc.add_paragraph('对连通平面图：V - E + F = 2，其中V是顶点数，E是边数，F是面数（包括外部面）。')
doc.add_paragraph('这是数学中最优美的公式之一，与多面体的欧拉公式V-E+F=2是同一回事。')

add_bold_para(doc, '平面图的必要条件：', '')
add_bullet(doc, '若G是简单平面图（V≥3），则E ≤ 3V - 6', '')
add_bullet(doc, '若G不含三角形，则E ≤ 2V - 4', '')
add_bullet(doc, '推论：K₅和K₃,₃不是平面图', '')

add_bold_para(doc, 'Kuratowski定理：', '')
doc.add_paragraph('图是平面图当且仅当它不包含K₅或K₃,₃的细分（subdivision）。'
    '这给出了平面图的一个完全刻画。')

add_bold_para(doc, '五色定理：', '')
doc.add_paragraph('平面图的色数不超过5。证明使用欧拉公式和归纳法，是"结构+归纳"的经典组合。')
doc.add_paragraph('四色定理：平面图的色数不超过4——证明需要计算机辅助，是第一个重要的计算机辅助证明。')

# ================================================================
# PART III: COUNTING
# ================================================================
doc.add_heading('第三部分：计数（Counting）', level=1)
doc.add_paragraph(
    '本部分共3章（第13-15章），介绍计数技术和渐近分析。'
    '计数是算法分析的核心——你需要知道有多少种可能、算法需要多少步骤、'
    '某种情况发生的概率是多少。本部分从求和公式出发，经过系统化的计数规则，'
    '最终到达生成函数——计数问题的"万能工具"。'
)

# ---- Chapter 13 ----
doc.add_heading('第13章 求和与渐近（Sums and Asymptotics）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '算法分析的核心技能——估计求和式的增长速度。当你分析一个算法的时间复杂度时，'
    '本质上就是在做求和与渐近分析。本章教你从精确求和到渐近估计的完整工具链。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '等差、等比、调和级数等常见求和', '精确求和：')
add_bullet(doc, '大O、Ω、Θ记号的准确定义和使用', '渐近记号：')
add_bullet(doc, '用积分近似求和，估计增长速度', '积分近似：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '常见求和公式：', '')
add_bullet(doc, '等差级数：1+2+...+n = n(n+1)/2', '')
add_bullet(doc, '等比级数：1+r+r²+...+rⁿ = (rⁿ⁺¹-1)/(r-1)', '')
add_bullet(doc, '调和级数：H_n = 1+1/2+1/3+...+1/n ≈ ln n + γ（γ≈0.5772是欧拉常数）', '')

add_bold_para(doc, '渐近记号：', '')
add_bullet(doc, 'f = O(g)：存在c和n₀，n≥n₀时f(n)≤c·g(n)——上界', '')
add_bullet(doc, 'f = Ω(g)：存在c和n₀，n≥n₀时f(n)≥c·g(n)——下界', '')
add_bullet(doc, 'f = Θ(g)：f = O(g)且f = Ω(g)——紧界', '')
add_bullet(doc, 'f = o(g)：lim f(n)/g(n) = 0——严格小于', '')

add_bold_para(doc, 'Stirling公式：', 'n! ≈ (n/e)^n · √(2πn)。这是估计阶乘和组合数的核心工具。')

# ---- Chapter 14 ----
doc.add_heading('第14章 基数规则（Cardinality Rules）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '系统化计数方法。从最简单的乘法规则到强大的容斥原理，本章教你一套完整的计数工具箱。'
    '这些工具直接应用于概率计算、算法分析中的情况枚举、以及组合优化问题。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '乘法规则和加法规则的灵活运用', '基本计数规则：')
add_bullet(doc, '排列组合的快速计算', '排列组合：')
add_bullet(doc, '处理重叠计数的系统方法', '容斥原理：')
add_bullet(doc, '将计数问题转化为双射问题', '双射计数：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '乘法规则：', '若|A|=a，|B|=b，则|A×B|=a·b。用于分步计数：每步有若干选择，总选择数等于各步选择数之积。')

add_bold_para(doc, '加法规则：', '若A和B不交，则|A∪B|=|A|+|B|。用于分类计数：将情况分成互不重叠的类，总数等于各类之和。')

add_bold_para(doc, '排列与组合：', '')
add_bullet(doc, '排列：从n个元素中选k个有序排列，P(n,k) = n!/(n-k)!', '')
add_bullet(doc, '组合：从n个元素中选k个无序子集，C(n,k) = n!/(k!(n-k)!)', '')
add_bullet(doc, '组合数的性质：C(n,k) = C(n,n-k)，C(n,k) = C(n-1,k-1)+C(n-1,k)（帕斯卡法则）', '')

add_bold_para(doc, '容斥原理（Inclusion-Exclusion）：', '')
doc.add_paragraph('两个集合：|A₁∪A₂| = |A₁|+|A₂|-|A₁∩A₂|')
doc.add_paragraph('n个集合：|A₁∪...∪Aₙ| = Σ|Aᵢ| - Σ|Aᵢ∩Aⱼ| + Σ|Aᵢ∩Aⱼ∩Aₖ| - ...')
doc.add_paragraph('应用：错排问题（n个元素的排列中没有任何元素在原位的排列数）')

add_bold_para(doc, '鸽巢原理（Pigeonhole Principle）：', '')
doc.add_paragraph('若n+1个物品放入n个盒子，则至少有一个盒子有至少2个物品。')
doc.add_paragraph('推广形式：若N个物品放入k个盒子，则至少有一个盒子有至少⌈N/k⌉个物品。')
doc.add_paragraph('应用：在任意5个整数中，必有3个之和是3的倍数。')

# ---- Chapter 15 ----
doc.add_heading('第15章 生成函数（Generating Functions）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '将计数问题代数化。生成函数是计数中最优雅、最强大的工具——它将复杂的组合问题转化为幂级数的代数运算。'
    '许多看似困难的计数问题，在生成函数框架下变得系统可解。'
    '生成函数也是求解递推关系的重要工具，连接了计数和递推两个主题。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '将计数问题转化为生成函数', '问题转化能力：')
add_bullet(doc, '幂级数的加法、乘法（卷积）、求导等运算', '代数运算能力：')
add_bullet(doc, '从生成函数提取系数来得到计数结果', '系数提取技巧：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '普通生成函数（OGF）：', '')
doc.add_paragraph('序列a₀,a₁,a₂,...的生成函数是G(x) = a₀ + a₁x + a₂x² + ...')
doc.add_paragraph('生成函数的系数就是我们要计数的对象，xⁿ的系数就是aₙ。')

add_bold_para(doc, '常见序列的生成函数：', '')
add_bullet(doc, '1,1,1,... → 1/(1-x)', '')
add_bullet(doc, '1,2,3,... → 1/(1-x)²', '')
add_bullet(doc, '1,k,k²,... → 1/(1-kx)', '')
add_bullet(doc, 'C(n,0),C(n,1),...,C(n,n) → (1+x)ⁿ（二项式定理）', '')

add_bold_para(doc, '生成函数的运算：', '')
add_bullet(doc, '加法：对应系数相加——对应计数的分类', '')
add_bullet(doc, '乘法（卷积）：(G·H)ₙ = Σaᵢbₙ₋ᵢ——对应计数的分步', '')
add_bullet(doc, '求导：G'(x)对应序列a₁,2a₂,3a₃,...', '')

add_bold_para(doc, '用生成函数解递推：', '')
doc.add_paragraph('以斐波那契数列为例：F(n)=F(n-1)+F(n-2)，F(0)=0，F(1)=1')
doc.add_paragraph('生成函数G(x) = x/(1-x-x²)，通过部分分式分解可得F(n) = (φⁿ-ψⁿ)/√5（Binet公式）')

# ================================================================
# PART IV: PROBABILITY
# ================================================================
doc.add_heading('第四部分：概率（Probability）', level=1)
doc.add_paragraph(
    '本部分共5章（第16-20章），是现代计算机科学中最重要的数学工具之一。'
    '从机器学习到密码学，从网络分析到随机算法，概率论无处不在。'
    '本部分从概率的基本定义出发，逐步构建条件概率、随机变量、期望和方差的理论，'
    '最终到达集中不等式和随机游走——这些是分析随机算法和机器学习理论的核心工具。'
)

# ---- Chapter 16 ----
doc.add_heading('第16章 事件与概率（Events and Probability）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '建立概率论的严格数学框架。从样本空间和事件出发，定义概率测度。'
    '本章强调概率不是"猜测"，而是有严格公理基础的数学理论。'
)

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '概率空间的三要素：', '')
add_bullet(doc, '样本空间Ω：所有可能结果的集合', '')
add_bullet(doc, '事件：样本空间的子集', '')
add_bullet(doc, '概率测度P：满足三条公理——非负性、规范性(P(Ω)=1)、可加性（互斥事件之并的概率等于概率之和）', '')

add_bold_para(doc, '等可能模型：', '若Ω有n个等可能结果，则P(A) = |A|/n。这是最直观的概率模型，将概率问题转化为计数问题。')

# ---- Chapter 17 ----
doc.add_heading('第17章 条件概率（Conditional Probability）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '条件概率是概率论中最重要的概念。贝叶斯定理是机器学习和统计推断的数学基础。'
    '本章特别强调条件概率中的直觉陷阱——人的直觉在条件概率问题上经常出错。'
)

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '条件概率：', 'P(A|B) = P(A∩B)/P(B)，表示"已知B发生时A发生的概率"。')

add_bold_para(doc, '贝叶斯定理：', 'P(A|B) = P(B|A)·P(A)/P(B)')
doc.add_paragraph('这是"逆概率"问题的核心：从观察到的结果B推断原因A的概率。')
doc.add_paragraph('应用：垃圾邮件过滤（朴素贝叶斯）、医疗诊断、机器学习分类器。')

add_bold_para(doc, '独立性：', 'P(A∩B) = P(A)·P(B)等价于P(A|B)=P(A)，即B的发生不影响A的概率。')
doc.add_paragraph('注意：互斥≠独立！事实上，若P(A)>0且P(B)>0，则A和B互斥意味着它们不独立。')

add_bold_para(doc, '经典直觉陷阱：', '')
doc.add_paragraph('检察官谬误：P(证据|无罪)很小不等于P(无罪|证据)很小。')
doc.add_paragraph('辛普森悖论：在分组数据中观察到的趋势，在合并数据中可能反转。')

# ---- Chapter 18 ----
doc.add_heading('第18章 随机变量与期望（Random Variables & Expectation）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '从事件概率过渡到数值分析。随机变量是量化随机现象的工具，'
    '期望是分析随机算法平均性能的核心指标。本章最重要的工具是期望的线性性——'
    '它让你无需知道随机变量的分布，就能计算复杂随机变量的期望。'
)

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '随机变量：', '从样本空间到实数的函数。不是"随机"的也不是"变量"——它是一个确定的函数，只是输入是随机的。')

add_bold_para(doc, '期望（Expected Value）：', 'E[X] = Σ x·P(X=x)。期望是随机变量的"加权平均"。')

add_bold_para(doc, '期望的线性性（Linearity of Expectation）：', '')
doc.add_paragraph('E[X₁+X₂+...+Xₙ] = E[X₁]+E[X₂]+...+E[Xₙ]')
doc.add_paragraph('关键：不需要X₁,...,Xₙ相互独立！这是期望线性性最强大的地方。')
doc.add_paragraph('经典应用：帽子检查问题——n个人随机取帽子，期望有1个人拿到自己的帽子（无论n多大）。')

add_bold_para(doc, '方差（Variance）：', '')
doc.add_paragraph('Var(X) = E[(X-E[X])²] = E[X²] - (E[X])²')
doc.add_paragraph('方差衡量随机变量偏离期望的程度，是分析算法稳定性（最坏情况与平均情况的差距）的关键指标。')

add_bold_para(doc, '常见分布：', '')
add_bullet(doc, '伯努利分布：P(X=1)=p, P(X=0)=1-p，E[X]=p', '')
add_bullet(doc, '二项分布：n次独立伯努利试验的成功次数，E[X]=np', '')
add_bullet(doc, '几何分布：首次成功所需的试验次数，E[X]=1/p', '')

# ---- Chapter 19 ----
doc.add_heading('第19章 偏离均值（Deviation from the Mean）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '只知道期望是不够的——你还需要知道随机变量有多大的概率偏离期望。'
    '集中不等式（Concentration Inequalities）给出了这种概率的上界，'
    '是分析随机算法、机器学习泛化误差、网络可靠性的核心工具。'
)

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, "马尔可夫不等式（Markov's Inequality）：", '')
doc.add_paragraph('P(X ≥ c) ≤ E[X]/c（X非负时）')
doc.add_paragraph('只需知道期望就能给出尾部概率的上界，但界限较松。')

add_bold_para(doc, "切比雪夫不等式（Chebyshev's Inequality）：", '')
doc.add_paragraph('P(|X-E[X]| ≥ c) ≤ Var(X)/c²')
doc.add_paragraph('利用方差给出更紧的界限。应用：民意调查的样本量估计。')

add_bold_para(doc, '切尔诺夫界（Chernoff Bound）：', '')
doc.add_paragraph('对独立伯努利随机变量之和X=X₁+...+Xₙ：')
doc.add_paragraph('P(X ≥ (1+δ)μ) ≤ (e^δ/(1+δ)^(1+δ))^μ，其中μ=E[X]')
doc.add_paragraph('切尔诺夫界是指数级的尾部界限——比切比雪夫的1/c²衰减快得多。')
doc.add_paragraph('应用：负载均衡（n个任务分配到n台机器，最大负载为O(ln n/ln ln n)）、随机算法的错误概率分析。')

# ---- Chapter 20 ----
doc.add_heading('第20章 随机游走（Random Walks）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '随机游走是概率论与图论的交叉主题，在物理学、金融学和计算机科学中都有重要应用。'
    '本章重点介绍PageRank算法——Google搜索引擎的核心——它本质上是图上的随机游走。'
)

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '随机游走基本概念：', '')
doc.add_paragraph('从某个顶点出发，每步随机选择一条边走到邻居。核心问题：')
add_bullet(doc, '期望多少步回到起点？（返回时间）', '')
add_bullet(doc, '期望多少步访问所有顶点？（覆盖时间）', '')
add_bullet(doc, '长期来看在每个顶点停留的时间比例是多少？（平稳分布）', '')

add_bold_para(doc, 'PageRank算法：', '')
doc.add_paragraph('网页排名 = 图上随机游走的平稳分布。')
doc.add_paragraph('核心思想：一个网页重要当且仅当很多重要的网页链接到它。')
doc.add_paragraph('数学模型：以概率α随机点击链接，以概率1-α跳转到随机页面（避免陷入死胡同）。')
doc.add_paragraph('平稳分布的存在性和唯一性由马尔可夫链理论保证。')

# ================================================================
# PART V: RECURRENCES
# ================================================================
doc.add_heading('第五部分：递推（Recurrences）', level=1)
doc.add_paragraph(
    '本部分只有1章（第21章），但它是全书的综合应用。递推关系是将前面所有工具串联起来的纽带：'
    '分治算法产生递推关系，递推关系需要计数和求和来求解，'
    '随机递推需要概率工具来分析。本章教你建立和求解递推关系的系统方法。'
)

# ---- Chapter 21 ----
doc.add_heading('第21章 递推关系（Recurrences）', level=2)

doc.add_heading('▎章节目的', level=3)
doc.add_paragraph(
    '递推关系是分析递归算法和分治算法的核心工具。当你写出一个递归程序时，'
    '它的时间复杂度自然表示为递推关系。本章教你三种求解递推的方法：'
    '替换法、展开法、和主定理。其中主定理是分析分治算法的标准工具。'
)

doc.add_heading('▎训练的核心能力', level=3)
add_bullet(doc, '从递归算法建立递推关系', '递推建立：')
add_bullet(doc, '用替换法、展开法、主定理求解递推', '递推求解：')
add_bullet(doc, '快速判断分治算法的时间复杂度', '主定理应用：')

doc.add_heading('▎核心知识点详解', level=3)

add_bold_para(doc, '替换法（Substitution Method）：', '')
doc.add_paragraph('猜一个解，用归纳法验证。关键在于如何猜——通常先展开前几项找规律。')

add_bold_para(doc, '展开法（Plug-and-Chug）：', '')
doc.add_paragraph('将递推关系反复展开，直到出现已知项，然后求和。')
doc.add_paragraph('例如：T(n) = T(n-1) + n → T(n) = T(0) + 1+2+...+n = O(n²)')

add_bold_para(doc, '主定理（Master Theorem）：', '')
doc.add_paragraph('对递推 T(n) = aT(n/b) + O(nᵈ)：')
doc.add_paragraph('  情况1：d < log_b(a) → T(n) = O(n^(log_b(a)))')
doc.add_paragraph('  情况2：d = log_b(a) → T(n) = O(nᵈ log n)')
doc.add_paragraph('  情况3：d > log_b(a) → T(n) = O(nᵈ)')
doc.add_paragraph('直觉：a是子问题数，n/b是子问题大小，nᵈ是合并代价。')
doc.add_paragraph('  如果子问题增长快于合并（情况1），递归叶节点主导')
doc.add_paragraph('  如果两者平衡（情况2），每层贡献相同，乘以层数log n')
doc.add_paragraph('  如果合并增长快于子问题（情况3），根节点主导')

add_bold_para(doc, '经典应用：', '')
add_bullet(doc, '归并排序：T(n)=2T(n/2)+O(n) → 情况2 → O(n log n)', '')
add_bullet(doc, '二分搜索：T(n)=T(n/2)+O(1) → 情况2(d=0, log₂2=1>0) → O(log n)', '')
add_bullet(doc, 'Strassen矩阵乘法：T(n)=7T(n/2)+O(n²) → 情况1 → O(n^log₂7)≈O(n^2.81)', '')

# ================================================================
# SUMMARY TABLE & ADVICE
# ================================================================
doc.add_heading('全书能力体系总览', level=1)

table = doc.add_table(rows=7, cols=4, style='Light Grid Accent 1')
table.cell(0,0).text = '部分'
table.cell(0,1).text = '核心数学工具'
table.cell(0,2).text = '训练的核心能力'
table.cell(0,3).text = '计算机科学应用'

row1 = table.rows[1].cells
row1[0].text = 'I 证明'
row1[1].text = '归纳法、反证法、逻辑公式、良序原理'
row1[2].text = '逻辑推理、证明构造、将直觉转化为严格论证'
row1[3].text = '程序正确性验证、形式化方法、芯片验证'

row2 = table.rows[2].cells
row2[0].text = 'II 结构'
row2[1].text = '数论、图论、关系、网络拓扑'
row2[2].text = '结构化建模、图论分析、模运算推理'
row2[3].text = '密码学(RSA)、网络设计、调度优化、社交网络分析'

row3 = table.rows[3].cells
row3[0].text = 'III 计数'
row3[1].text = '排列组合、容斥原理、生成函数、渐近分析'
row3[2].text = '精确计数、渐近估计、代数化计数'
row3[3].text = '算法复杂度分析、组合优化、概率计算'

row4 = table.rows[4].cells
row4[0].text = 'IV 概率'
row4[1].text = '贝叶斯定理、期望线性性、集中不等式'
row4[2].text = '概率建模、随机变量分析、偏差估计'
row4[3].text = '机器学习、随机算法、PageRank、可靠性分析'

row5 = table.rows[5].cells
row5[0].text = 'V 递推'
row5[1].text = '主定理、替换法、展开法'
row5[2].text = '递推关系建立与求解、分治分析'
row5[3].text = '分治算法分析、递归算法复杂度'

row6 = table.rows[6].cells
row6[0].text = '综合'
row6[1].text = '以上所有工具的组合'
row6[2].text = '数学建模+严格推理+算法分析的综合能力'
row6[3].text = '独立解决CS研究中的数学问题'

doc.add_paragraph('')

doc.add_heading('学习建议', level=1)

doc.add_paragraph(
    '1. 证明部分是地基中的地基。归纳法和逻辑推理如果不扎实，后续所有内容都会吃力。'
    '建议：每学一种证明模板，至少做5道题，直到模板变成肌肉记忆。', style='List Number')

doc.add_paragraph(
    '2. 结构部分的核心是"建模"——学会将实际问题翻译成图论/数论语言，这比记住定理更重要。'
    '建议：每学一种图结构，想一个实际应用场景，自己画出图并分析。', style='List Number')

doc.add_paragraph(
    '3. 计数部分的核心是"系统化"——不要靠"数数"，要用规则和公式。'
    '容斥原理和生成函数是最强大的两个工具，必须反复练习。'
    '建议：遇到计数问题，先尝试用乘法规则/容斥原理系统化求解，而不是凭直觉数。', style='List Number')

doc.add_paragraph(
    '4. 概率部分的核心是"期望线性性"——这个技巧在算法分析中出现频率极高，必须熟练掌握。'
    '同时要警惕条件概率中的直觉陷阱，贝叶斯定理必须能熟练运用。'
    '建议：用期望线性性重做所有期望计算题，体会它的威力。', style='List Number')

doc.add_paragraph(
    '5. 递推部分的核心是"主定理"——它是分析分治算法的标准工具，必须能快速判断三种情况。'
    '建议：对常见的分治算法（归并排序、快速排序、Strassen等）都用主定理分析一遍。', style='List Number')

doc.add_paragraph(
    '6. 全书最重要的思维方式：不要相信直觉，要相信证明。每当你觉得"显然"的时候，'
    '问自己：我能严格证明吗？如果不行，你的"显然"可能就是错的。', style='List Number')

# Save
doc.save('c:/Users/27063/Desktop/AI   文件夹/计算机科学数学_全书重点总结.docx')
print("Word文档已生成！")
