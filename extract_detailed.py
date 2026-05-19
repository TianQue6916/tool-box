
import PyPDF2

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

# Extract key pages for each chapter - more pages for thorough content
chapter_ranges = [
    ("Ch1", 16, 37),
    ("Ch2", 38, 51),
    ("Ch3", 52, 91),
    ("Ch4", 92, 125),
    ("Ch5", 126, 183),
    ("Ch6", 184, 215),
    ("Ch7", 216, 251),
    ("Ch8", 254, 327),
    ("Ch9", 328, 383),
    ("Ch10", 384, 403),
    ("Ch11", 404, 483),
    ("Ch12", 484, 511),
    ("Ch13", 514, 561),
    ("Ch14", 562, 637),
    ("Ch15", 638, 675),
    ("Ch16", 678, 707),
    ("Ch17", 708, 749),
    ("Ch18", 750, 799),
    ("Ch19", 800, 849),
    ("Ch20", 850, 875),
    ("Ch21", 878, 910),
]

for label, start, end in chapter_ranges:
    print(f"\n{'='*80}")
    print(f"=== {label} (PDF pages {start}-{end}) ===")
    print(f"{'='*80}")
    # Extract first 6 pages of each chapter
    for pg in range(start, min(start+7, end+1)):
        idx = pg - 1
        if idx < len(reader.pages):
            text = reader.pages[idx].extract_text()
            if text:
                print(f"\n--- Page {pg} ---")
                print(text[:3000])
