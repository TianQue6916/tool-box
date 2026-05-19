
import re

path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Strategy: replace ALL double-quoted strings that contain Chinese characters
# with single-quoted strings, removing inner quote marks

# Find all string literals that use double quotes and contain Chinese
# Replace inner ASCII double quotes used as Chinese quotation marks with nothing
lines = content.split('\n')
new_lines = []

for line in lines:
    # Skip lines that are purely structural (no Chinese text in quotes)
    if not any('\u4e00' <= c <= '\u9fff' for c in line):
        new_lines.append(line)
        continue

    # For lines with Chinese, we need to be careful
    # Simple approach: if a line has an odd number of unescaped double quotes,
    # it likely has a problem. Fix by converting outer quotes to single quotes
    # and removing inner quotes

    # Count double quotes not preceded by backslash
    dq_count = len(re.findall(r'(?<!\\)"', line))

    if dq_count > 4 and dq_count % 2 != 0:
        # Problematic line - convert to single quotes and remove inner quotes
        # First, identify the Python string boundaries
        # Simple fix: replace the entire line's double quotes with single quotes
        # for the string delimiters, and remove inner quotes

        # Actually, let's just find Chinese text between double quotes
        # and replace the inner double quotes
        new_line = line
        # Replace patterns like: "text"word"text" -> 'text word text'
        # by converting all double-quoted strings to single-quoted
        new_line = re.sub(r'"([^"]*[一-鿿][^"]*)"', r"'\1'", new_line)
        new_lines.append(new_line)
    else:
        new_lines.append(line)

result = '\n'.join(new_lines)

with open(path, 'w', encoding='utf-8') as f:
    f.write(result)

print("Applied comprehensive quote fix")

# Verify compilation
import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("File compiles successfully!")
except py_compile.PyCompileError as e:
    print(f"Still has error: {e}")
