import os
import re

app_dir = r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app'
templates_dir = os.path.join(app_dir, 'templates')

def check_templates():
    missing = []
    
    # Check render_template in __init__.py
    with open(os.path.join(app_dir, '__init__.py'), 'r', encoding='utf-8') as f:
        content = f.read()
        
    for m in re.finditer(r'render_template\(\s*[\"\']([^\"\']+)[\"\']', content):
        tpl = m.group(1)
        path = os.path.join(templates_dir, tpl)
        if not os.path.exists(path):
            missing.append(f'__init__.py -> {tpl}')
            
    # Check includes in all templates
    for root, _, files in os.walk(templates_dir):
        for file in files:
            if file.endswith('.html'):
                file_path = os.path.join(root, file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                for m in re.finditer(r'\{%\s*include\s+[\"\']([^\"\']+)[\"\']\s*%\}', content):
                    tpl = m.group(1)
                    path = os.path.join(templates_dir, tpl)
                    if not os.path.exists(path):
                        missing.append(f'{file} -> {tpl}')
                        
    for m in missing:
        print(f'MISSING: {m}')

if __name__ == '__main__':
    check_templates()
