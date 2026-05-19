
import re

path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Remove Chinese quotes that cause issues inside Python strings
content = content.replace('“', '').replace('”', '')

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Fixed Chinese quotes in add_problems.py")
