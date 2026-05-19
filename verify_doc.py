
from docx import Document

doc = Document(r"c:\Users\27063\Desktop\AI   文件夹\计算机科学数学_全书重点总结_v2.docx")
print(f"Paragraphs: {len(doc.paragraphs)}")
print(f"Tables: {len(doc.tables)}")
print(f"Images: {len(doc.inline_shapes)}")

headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
print(f"Headings: {len(headings)}")
for h in headings:
    print(f"  {h[:80]}")
