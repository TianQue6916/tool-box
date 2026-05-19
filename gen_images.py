
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
# removed unused import
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import os

output_dir = r"c:\Users\27063\Desktop\AI   文件夹"
img_dir = os.path.join(output_dir, "summary_images")
os.makedirs(img_dir, exist_ok=True)

# ============ Generate Images ============
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# Image 1: Well Ordering Proof Template
fig, ax = plt.subplots(1, 1, figsize=(8, 4))
ax.set_xlim(0, 10)
ax.set_ylim(0, 6)
ax.axis('off')
ax.set_title("Well Ordering Proof Template (良序证明模板)", fontsize=14, fontweight='bold')

steps = [
    "1. Assume \u2203 counterexample n (假设存在反例n)",
    "2. Let C = {n \u2208 N | P(n) is false} (定义反例集合C)",
    "3. By WOP, C has minimum element m (由良序原理, C有最小元m)",
    "4. Show m cannot be the minimum (证明m不可能是最小反例)",
    "5. Contradiction! \u2234 P(n) holds for all n (矛盾! \u2234 P(n)对所有n成立)"
]
for i, step in enumerate(steps):
    y = 5 - i * 1.1
    ax.text(0.5, y, step, fontsize=11, va='center',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='lightyellow', edgecolor='gray'))
    if i < 4:
        ax.annotate('', xy=(0.5, y-0.5), xytext=(0.5, y-0.1),
                    arrowprops=dict(arrowstyle='->', color='blue', lw=2))

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "well_ordering.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 2: Induction Template
fig, ax = plt.subplots(1, 1, figsize=(8, 3.5))
ax.set_xlim(0, 10)
ax.set_ylim(0, 5)
ax.axis('off')
ax.set_title("Induction Proof Template (归纳法证明模板)", fontsize=14, fontweight='bold')

boxes = [
    (5, 4.2, "Base Case (基础步): Prove P(0)", 'lightgreen'),
    (5, 2.8, "Inductive Step (归纳步): Assume P(n) \u2192 Prove P(n+1)", 'lightyellow'),
    (5, 1.4, "Conclusion (结论): \u2234 P(n) holds for all n \u2265 0", 'lightsalmon'),
]
for x, y, text, color in boxes:
    ax.text(x, y, text, fontsize=11, ha='center', va='center',
            bbox=dict(boxstyle='round,pad=0.4', facecolor=color, edgecolor='black', lw=1.5))

ax.annotate('', xy=(5, 3.3), xytext=(5, 3.7), arrowprops=dict(arrowstyle='->', lw=2))
ax.annotate('', xy=(5, 1.9), xytext=(5, 2.3), arrowprops=dict(arrowstyle='->', lw=2))

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "induction.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 3: RSA Flow
fig, ax = plt.subplots(1, 1, figsize=(9, 5))
ax.set_xlim(0, 12)
ax.set_ylim(0, 8)
ax.axis('off')
ax.set_title("RSA Encryption System (RSA加密系统)", fontsize=14, fontweight='bold')

# Key Generation
ax.text(1, 7, "Key Generation (密钥生成)", fontsize=11, fontweight='bold', color='darkblue')
kg_steps = ["Choose primes p, q (选素数p,q)", "n = pq, \u03c6(n) = (p-1)(q-1)",
            "Choose e: gcd(e,\u03c6(n))=1", "Compute d = e\u207b\u00b9 mod \u03c6(n)",
            "Public Key (公钥): (e, n)", "Private Key (私钥): d"]
for i, s in enumerate(kg_steps):
    ax.text(1, 6.2 - i*0.7, s, fontsize=9, va='center')

# Encryption
ax.text(7, 7, "Encryption (加密)", fontsize=11, fontweight='bold', color='darkred')
ax.text(7, 6.2, "c = m^e mod n", fontsize=11, va='center',
        bbox=dict(boxstyle='round', facecolor='mistyrose', edgecolor='red'))

# Decryption
ax.text(7, 4.5, "Decryption (解密)", fontsize=11, fontweight='bold', color='darkgreen')
ax.text(7, 3.7, "m = c^d mod n", fontsize=11, va='center',
        bbox=dict(boxstyle='round', facecolor='honeydew', edgecolor='green'))

# Correctness
ax.text(7, 2, "Correctness (正确性)", fontsize=11, fontweight='bold', color='purple')
ax.text(7, 1.2, "m^ed \u2261 m (mod n)", fontsize=11, va='center',
        bbox=dict(boxstyle='round', facecolor='lavender', edgecolor='purple'))
ax.text(7, 0.5, "By Euler's Theorem (欧拉定理)", fontsize=9, va='center', style='italic')

# Arrows
ax.annotate('', xy=(6.5, 6.2), xytext=(5.5, 4.5), arrowprops=dict(arrowstyle='->', lw=1.5, color='red'))
ax.annotate('', xy=(6.5, 3.7), xytext=(5.5, 2.5), arrowprops=dict(arrowstyle='->', lw=1.5, color='green'))

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "rsa.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 4: Euler Formula
fig, ax = plt.subplots(1, 1, figsize=(6, 5))
ax.set_xlim(-2, 2)
ax.set_ylim(-2, 2)
ax.axis('off')
ax.set_title("Euler's Formula (欧拉公式): V - E + F = 2", fontsize=13, fontweight='bold')

# Draw a simple planar graph (tetrahedron projection)
import matplotlib.patches as patches
triangle = plt.Polygon([(-1, -0.8), (1, -0.8), (0, 1)], fill=True, facecolor='lightyellow', edgecolor='black', lw=2)
ax.add_patch(triangle)
ax.plot([0, 0], [1, -0.8], 'k-', lw=2)
ax.plot([0, -1], [1, -0.8], 'k--', lw=1.5, color='gray')
ax.plot([0, 1], [1, -0.8], 'k--', lw=1.5, color='gray')

# Vertices
for (x, y, label) in [(-1, -0.8, 'v1'), (1, -0.8, 'v2'), (0, 1, 'v3'), (0, -0.8, 'v4')]:
    ax.plot(x, y, 'ko', markersize=8)
    ax.text(x, y-0.2, label, ha='center', fontsize=9)

ax.text(0, -1.6, "V=4, E=6, F=4  \u2192  4-6+4=2 \u2713", ha='center', fontsize=11,
        bbox=dict(boxstyle='round', facecolor='lightgreen', edgecolor='green'))

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "euler.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 5: Master Theorem
fig, ax = plt.subplots(1, 1, figsize=(9, 5))
ax.set_xlim(0, 10)
ax.set_ylim(0, 7)
ax.axis('off')
ax.set_title("Master Theorem (主定理): T(n) = aT(n/b) + O(n^d)", fontsize=13, fontweight='bold')

cases = [
    (5, 5.8, "Case 1 (情况1): d < log_b(a)", "T(n) = O(n^{log_b(a)})", 
     "Leaves dominate (叶节点主导)", 'lightyellow'),
    (5, 3.8, "Case 2 (情况2): d = log_b(a)", "T(n) = O(n^d log n)", 
     "Balanced (每层均衡)", 'lightgreen'),
    (5, 1.8, "Case 3 (情况3): d > log_b(a)", "T(n) = O(n^d)", 
     "Root dominates (根节点主导)", 'lightsalmon'),
]

for x, y, title, formula, desc, color in cases:
    ax.text(x, y+0.5, title, fontsize=11, ha='center', fontweight='bold', color='darkblue')
    ax.text(x, y, formula, fontsize=12, ha='center',
            bbox=dict(boxstyle='round,pad=0.3', facecolor=color, edgecolor='black'))
    ax.text(x, y-0.5, desc, fontsize=9, ha='center', style='italic', color='gray')

# Examples
ax.text(5, 0.5, "Examples (示例): Merge Sort (归并排序) 2T(n/2)+O(n) \u2192 Case 2 \u2192 O(n log n)  |  "
        "Binary Search (二分搜索) T(n/2)+O(1) \u2192 O(log n)", 
        fontsize=8, ha='center', va='center',
        bbox=dict(boxstyle='round', facecolor='lavender', edgecolor='purple'))

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "master.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 6: Bayesian Theorem
fig, ax = plt.subplots(1, 1, figsize=(7, 4))
ax.set_xlim(0, 8)
ax.set_ylim(0, 5)
ax.axis('off')
ax.set_title("Bayes' Theorem (贝叶斯定理)", fontsize=14, fontweight='bold')

ax.text(4, 3.5, "P(A|B) = P(B|A) \u00b7 P(A) / P(B)", fontsize=16, ha='center', va='center',
        bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', edgecolor='black', lw=2))

ax.text(2, 1.5, "P(A|B): Posterior (后验概率)", fontsize=10, ha='center', color='darkblue')
ax.text(6, 1.5, "P(B|A): Likelihood (似然)", fontsize=10, ha='center', color='darkred')
ax.text(2, 0.7, "P(A): Prior (先验概率)", fontsize=10, ha='center', color='darkgreen')
ax.text(6, 0.7, "P(B): Evidence (证据)", fontsize=10, ha='center', color='purple')

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "bayes.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 7: Concentration Inequalities comparison
fig, ax = plt.subplots(1, 1, figsize=(8, 4))
x = np.linspace(0.01, 5, 200)
markov = 1/x
chebyshev = 1/x**2
chernoff = np.exp(-x**2/2)

ax.plot(x, markov, 'r-', lw=2, label="Markov (马尔可夫): O(1/c)")
ax.plot(x, chebyshev, 'b-', lw=2, label="Chebyshev (切比雪夫): O(1/c\u00b2)")
ax.plot(x, chernoff, 'g-', lw=2, label="Chernoff (切尔诺夫): O(e^{-c\u00b2})")
ax.set_xlabel("Deviation c (偏离程度)", fontsize=11)
ax.set_ylabel("Upper Bound (上界)", fontsize=11)
ax.set_title("Concentration Inequalities Comparison (集中不等式比较)", fontsize=13, fontweight='bold')
ax.legend(fontsize=10)
ax.set_ylim(0, 2)
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "concentration.png"), dpi=150, bbox_inches='tight')
plt.close()

# Image 8: PageRank
fig, ax = plt.subplots(1, 1, figsize=(7, 5))
ax.set_xlim(-0.5, 3.5)
ax.set_ylim(-0.5, 3.5)
ax.axis('off')
ax.set_title("PageRank: Random Walk on Graph (图上的随机游走)", fontsize=13, fontweight='bold')

nodes = [(0, 2, 'A\n0.33'), (1.5, 3, 'B\n0.20'), (3, 2, 'C\n0.10'),
         (0, 0.5, 'D\n0.10'), (1.5, 0, 'E\n0.07'), (3, 0.5, 'F\n0.20')]
edges = [(0,1), (1,2), (2,5), (5,3), (3,0), (0,3), (1,4), (4,5), (2,1)]

for i, (x, y, label) in enumerate(nodes):
    circle = plt.Circle((x, y), 0.35, facecolor='lightskyblue', edgecolor='black', lw=2)
    ax.add_patch(circle)
    ax.text(x, y, label, ha='center', va='center', fontsize=9, fontweight='bold')

for i, j in edges:
    x1, y1 = nodes[i][0], nodes[i][1]
    x2, y2 = nodes[j][0], nodes[j][1]
    dx, dy = x2-x1, y2-y1
    dist = np.sqrt(dx**2+dy**2)
    ax.annotate('', xy=(x2-dx/dist*0.4, y2-dy/dist*0.4), 
                xytext=(x1+dx/dist*0.4, y1+dy/dist*0.4),
                arrowprops=dict(arrowstyle='->', lw=1.5, color='darkblue'))

ax.text(1.5, -0.8, "Steady Distribution (平稳分布) = PageRank Scores", 
        ha='center', fontsize=10, style='italic', color='gray')

plt.tight_layout()
plt.savefig(os.path.join(img_dir, "pagerank.png"), dpi=150, bbox_inches='tight')
plt.close()

print("All images generated successfully!")
