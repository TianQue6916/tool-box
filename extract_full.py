
import PyPDF2
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

# Extract more pages for each chapter to get thorough content
# Format: (label, start_page, end_page)
chapter_ranges = [
    ("Introduction", 12, 15),
    ("Ch1 What is a Proof", 16, 37),
    ("Ch2 Well Ordering", 38, 51),
    ("Ch3 Logical Formulas", 52, 91),
    ("Ch4 Math Data Types", 92, 125),
    ("Ch5 Induction", 126, 183),
    ("Ch6 Recursive Data Types", 184, 215),
    ("Ch7 Infinite Sets", 216, 251),
    ("Ch8 Number Theory", 254, 327),
    ("Ch9 Directed Graphs", 328, 383),
    ("Ch10 Communication Networks", 384, 403),
    ("Ch11 Simple Graphs", 404, 483),
    ("Ch12 Planar Graphs", 484, 511),
    ("Ch13 Sums and Asymptotics", 514, 561),
    ("Ch14 Cardinality Rules", 562, 637),
    ("Ch15 Generating Functions", 638, 675),
    ("Ch16 Events and Probability", 678, 707),
    ("Ch17 Conditional Probability", 708, 749),
    ("Ch18 Random Variables", 750, 799),
    ("Ch19 Deviation from Mean", 800, 849),
    ("Ch20 Random Walks", 850, 875),
    ("Ch21 Recurrences", 878, 910),
]

for label, start, end in chapter_ranges:
    print(f"\n{'='*60}")
    print(f"=== {label} (PDF pages {start}-{end}) ===")
    print(f"{'='*60}")
    # Extract first 3 pages of each chapter for key content
    for pg in range(start, min(start+4, end+1)):
        idx = pg - 1
        if idx < len(reader.pages):
            text = reader.pages[idx].extract_text()
            if text:
                print(f"\n--- Page {pg} ---")
                print(text[:2000])
