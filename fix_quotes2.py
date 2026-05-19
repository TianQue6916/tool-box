
path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Fix lines that have Chinese quotes inside Python strings
# Replace problematic characters
fixed_lines = []
for line in lines:
    # Replace fullwidth left/right double quotation marks
    line = line.replace('\u201c', '').replace('\u201d', '')
    # Also try the actual bytes
    line = line.replace('\xe2\x80\x9c', '').replace('\xe2\x80\x9d', '')
    fixed_lines.append(line)

with open(path, 'w', encoding='utf-8') as f:
    f.writelines(fixed_lines)

# Verify
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Check for remaining problematic chars
for i, line in enumerate(content.split('\n'), 1):
    if '\u201c' in line or '\u201d' in line:
        print(f"Line {i} still has Chinese quotes: {line[:80]}")

print("Done fixing quotes")
