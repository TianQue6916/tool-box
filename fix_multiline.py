
path = r"c:\Users\27063\Desktop\AI   文件夹\add_supplement.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix: replace actual newlines inside string concatenation that use 
"
# Pattern: a line ending with text (no \n) followed by "
" on next line
# This means the string has an unescaped newline

import re

# Replace patterns where a string has a raw newline followed by " on the next line
# e.g.  "some text
"
    "more text"
# should be: "some text\n"
    "more text"

lines = content.split('
')
new_lines = []
i = 0
while i < len(lines):
    line = lines[i]
    # Check if this line is a standalone quote on a string continuation
    # Pattern: line is just '"
' or starts with whitespace then '"'
    stripped = line.strip()
    if stripped == '"' and i > 0:
        # This is a continuation quote - merge with previous line
        prev = new_lines[-1]
        # Add \n to the previous line and skip this line
        if prev.endswith('"'):
            new_lines[-1] = prev[:-1] + '\n"'
        else:
            new_lines.append(line)
    else:
        new_lines.append(line)
    i += 1

result = '
'.join(new_lines)

with open(path, 'w', encoding='utf-8') as f:
    f.write(result)

print("Fixed multiline strings")

# Verify
import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("File compiles successfully!")
except py_compile.PyCompileError as e:
    print(f"Still has error: {e}")
