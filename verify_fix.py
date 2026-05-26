import re

with open(r'c:\Users\leahliu\WorkBuddy\20260525175100\budai_planner.html', 'r', encoding='utf-8') as f:
    content = f.read()

matches = re.findall(r'\$\([\'"]?(\w+)[\'"]?\)', content)
if matches:
    print('STILL HAS $() calls:', matches)
else:
    print('OK - No more $() calls')

# Check line 474 area
lines = content.split('\n')
for i in range(473, 477):
    print(f'{i+1}: {lines[i][:100]}')
