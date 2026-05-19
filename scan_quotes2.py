
path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find lines where double quotes might be nested incorrectly
# Strategy: for each line, try to parse it
import ast
import re

problematic = []
for i, line in enumerate(lines, 1):
    stripped = line.strip()
    if not stripped or stripped.startswith('#') or stripped.startswith('doc.') and stripped.endswith(')'):
        continue
    # Look for patterns like: "text"word"text" where inner quotes are not escaped
    # This happens when Chinese text uses " as quotation marks inside Python strings
    # Find all " positions
    in_string = False
    string_char = None
    quote_positions = []
    for j, ch in enumerate(stripped):
        if ch in ['"', "'"]:
            if not in_string:
                in_string = True
                string_char = ch
                quote_positions.append(j)
            elif ch == string_char:
                in_string = False
                quote_positions.append(j)
                string_char = None
    # If we end still in a string, something is wrong
    if in_string:
        problematic.append((i, stripped[:100]))

if problematic:
    print("Lines with potential quote nesting issues:")
    for num, text in problematic:
        print(f"  Line {num}: {text}")
else:
    print("No issues found by manual scan")
