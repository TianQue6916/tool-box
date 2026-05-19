
path = r"c:\Users\27063\Desktop\AI   文件夹\add_problems.py"
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find all lines with Chinese-style quotes (curly double quotes used inside Python strings)
# These are the problematic characters: “ " and ” "
lines = content.split('\n')
for i, line in enumerate(lines, 1):
    # Check for any character that could be a curly/smart quote
    for j, ch in enumerate(line):
        code = ord(ch)
        if code in [8220, 8221, 8216, 8217]:  # " " ' '
            print(f"Line {i}, col {j}: char U+{code:04X} ({ch}) context: ...{line[max(0,j-10):j+10]}...")
