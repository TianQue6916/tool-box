
import PyPDF2

reader = PyPDF2.PdfReader("c:/Users/27063/Desktop/AI   文件夹/计算机离散数学   MIT   计算机科学数学 已打印.pdf")

for pg in [6, 7, 8]:
    idx = pg - 1
    if idx < len(reader.pages):
        text = reader.pages[idx].extract_text()
        if text:
            print(f"--- Page {pg} ---")
            print(text[:3000])
            print()
