
path = r"c:\Users\27063\Desktop\AI   文件夹\add_supplement_v2.py"
with open(path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find all lines with potential quote issues
# Check for lines where double quotes appear inside double-quoted strings
import re

problem_lines = []
for i, line in enumerate(lines, 1):
    stripped = line.strip()
    # Skip comments and empty lines
    if not stripped or stripped.startswith('#'):
        continue
    # Count quotes - if odd number, there might be an issue
    # Check for Chinese curly quotes
    if '\u201c' in line or '\u201d' in line:
        problem_lines.append((i, line.rstrip()[:100]))
    # Check for lines with more than 2 double-quote pairs that might conflict
    count = line.count('"')
    if count > 4 and count % 2 != 0:
        problem_lines.append((i, f"Odd quotes ({count}): {line.rstrip()[:80]}"))

if problem_lines:
    print("Potential problem lines:")
    for num, text in problem_lines:
        print(f"  Line {num}: {text}")
else:
    print("No obvious quote issues found")

# Also try to compile the file
import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("File compiles successfully!")
except py_compile.PyCompileError as e:
    print(f"Compile error: {e}")
