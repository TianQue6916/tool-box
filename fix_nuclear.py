
path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Nuclear option: replace ALL double-quoted strings with single-quoted strings
# This is safe because our strings don't contain single quotes internally
import re

# Match double-quoted strings and convert to single-quoted
# But we need to handle multi-line strings carefully
# Strategy: replace " at string boundaries with '

# Simple approach: just replace every " with ' in the entire file
# But that would break comments and other things
# Better: only in lines that have Chinese characters

lines = content.split('\n')
new_lines = []
for line in lines:
    has_chinese = any('\u4e00' <= c <= '\u9fff' for c in line)
    if has_chinese:
        # Replace all double quotes with single quotes in this line
        line = line.replace('"', "'")
    new_lines.append(line)

result = '\n'.join(new_lines)

with open(path, 'w', encoding='utf-8') as f:
    f.write(result)

print("Replaced all double quotes with single quotes in Chinese lines")

# Verify
import py_compile
try:
    py_compile.compile(path, doraise=True)
    print("File compiles successfully!")
except py_compile.PyCompileError as e:
    print(f"Still has error: {e}")
