
import re

path = r"c:\Users\27063\Desktop\AI   文件夹\add_supplement.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find all lines that are just a standalone double quote (possibly with whitespace)
# These represent raw newlines inside strings that need to be converted to \n
lines = content.split('\n')
new_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    stripped = line.strip()
    if stripped == '"' and i > 0:
        # This is a continuation quote - merge with previous line
        prev = new_lines[-1]
        if prev.rstrip().endswith('"'):
            new_lines[-1] = prev.rstrip()[:-1] + '\\n"'
        else:
            new_lines.append(line)
    else:
        new_lines.append(line)
    i += 1

result = '\n'.join(new_lines)

with open(path, 'w', encoding='utf-8') as f:
    f.write(result)

print("Fixed multiline strings")

import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("File compiles successfully!")
except py_compile.PyCompileError as e:
    print(f"Still has error: {e}")
