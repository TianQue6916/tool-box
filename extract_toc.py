
import PyPDF2

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

# Extract Table of Contents pages (usually at the beginning)
# Try pages 4-14 which likely contain TOC
for pg in range(4, 15):
    idx = pg - 1
    if idx < len(reader.pages):
        text = reader.pages[idx].extract_text()
        if text and ('Contents' in text or 'Chapter' in text or '1 What' in text):
            print(f"--- Page {pg} ---")
            print(text[:3000])
            print()
