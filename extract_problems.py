
import PyPDF2

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

# Extract problem sections from each chapter
# Problems typically appear at the end of each chapter section
chapter_problem_pages = [
    ("Ch1", 30, 37),
    ("Ch2", 46, 51),
    ("Ch3", 80, 91),
    ("Ch4", 116, 125),
    ("Ch5", 170, 183),
    ("Ch6", 208, 215),
    ("Ch7", 244, 251),
    ("Ch8", 316, 327),
    ("Ch9", 374, 383),
    ("Ch10", 398, 403),
    ("Ch11", 472, 483),
    ("Ch12", 506, 511),
    ("Ch13", 554, 561),
    ("Ch14", 628, 637),
    ("Ch15", 668, 675),
    ("Ch16", 700, 707),
    ("Ch17", 742, 749),
    ("Ch18", 792, 799),
    ("Ch19", 842, 849),
    ("Ch20", 870, 875),
    ("Ch21", 904, 910),
]

with open("c:/Users/27063/Desktop/AI   文件夹/problems_text.txt", "w", encoding="utf-8") as f:
    for label, start, end in chapter_problem_pages:
        f.write(f"\n{'='*80}\n")
        f.write(f"=== {label} Problems (PDF pages {start}-{end}) ===\n")
        f.write(f"{'='*80}\n")
        for pg in range(start, end+1):
            idx = pg - 1
            if idx < len(reader.pages):
                text = reader.pages[idx].extract_text()
                if text:
                    f.write(f"\n--- Page {pg} ---\n")
                    f.write(text)

print("Done! Problems extracted.")
