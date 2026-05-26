import re

with open(r'c:\Users\leahliu\WorkBuddy\20260525175100\budai_planner.html', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Replace all $(id) with document.getElementById(id)
# Pattern: $(id) or $('id')
def replace_dollar(match):
    id_val = match.group(1).strip('"').strip("'")
    return "document.getElementById('" + id_val + "')"

content = re.sub(r'\$\([\'"]?(\w+)[\'"]?\)', replace_dollar, content)

# Fix 2: Fix line 474 - remove stray backslash before single quote in template literal
content = content.replace("html||\\'<tr>", "html||'<tr>")

# Write fixed version
with open(r'c:\Users\leahliu\WorkBuddy\20260525175100\budai_planner.html', 'w', encoding='utf-8') as f:
    f.write(content)

print('Fixed! All $() replaced with document.getElementById()')
print('Template literal escape also fixed')
