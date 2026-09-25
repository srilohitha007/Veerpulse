import os

app_file = r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app\templates\welfare_officer_personnel_overview.html'

with open(app_file, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('placeholder="Example: VP1043"', 'placeholder="Example: ID1234"')
content = content.replace('placeholder="Example: VP1042"', 'placeholder="Example: ID1234"')

with open(app_file, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done template replacements.")
