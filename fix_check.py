import re

with open(r'c:\Users\leahliu\WorkBuddy\20260525175100\budai_planner.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Find all $() usages
matches = re.findall(r'\$\([\'"]?(\w+)[\'"]?\)', content)
print('All $() calls:', matches)
print('Total count:', len(matches))

# Check for syntax errors around line 474
lines = content.split('\n')
print(f'Total lines: {len(lines)}')
for i in range(470, 480):
    print(f'{i+1}: {repr(lines[i])}')
