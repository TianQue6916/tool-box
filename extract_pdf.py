import PyPDF2

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

# Extract key pages: Part introductions and first pages of each chapter
# The PDF has ~11 pages of front matter before page 1 of the book
# So book page N corresponds to PDF page N+11 approximately

key_pages = {
    "Part I Proofs Intro": 12,
    "Ch1 What is a Proof": 16,
    "Ch2 Well Ordering": 38,
    "Ch3 Logical Formulas": 52,
    "Ch4 Math Data Types": 92,
    "Ch5 Induction": 126,
    "Ch6 Recursive Data Types": 184,
    "Ch7 Infinite Sets": 216,
    "Part II Structures Intro": 252,
    "Ch8 Number Theory": 254,
    "Ch9 Directed Graphs": 328,
    "Ch10 Communication Networks": 384,
    "Ch11 Simple Graphs": 404,
    "Ch12 Planar Graphs": 484,
    "Part III Counting Intro": 512,
    "Ch13 Sums and Asymptotics": 514,
    "Ch14 Cardinality Rules": 562,
    "Ch15 Generating Functions": 638,
    "Part IV Probability Intro": 676,
    "Ch16 Events and Probability": 678,
    "Ch17 Conditional Probability": 708,
    "Ch18 Random Variables": 750,
    "Ch19 Deviation from Mean": 800,
    "Ch20 Random Walks": 850,
    "Part V Recurrences Intro": 876,
    "Ch21 Recurrences": 878,
}

for label, pg in key_pages.items():
    idx = pg - 1
    if idx < len(reader.pages):
        text = reader.pages[idx].extract_text()
        if text:
            print("\n=== " + label + " (PDF page " + str(pg) + ") ===")
            print(text[:1200])

