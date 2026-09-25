import sys

def extract_func(func_name):
    with open(r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app\__init__.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    in_func = False
    indent = 0
    for i, line in enumerate(lines):
        if line.strip().startswith(f'def {func_name}('):
            in_func = True
            indent = len(line) - len(line.lstrip())
            print(f"--- Line {i+1} ---")
            print(line, end='')
            continue
        
        if in_func:
            if line.strip() != '' and (len(line) - len(line.lstrip())) <= indent:
                break
            print(line, end='')

if __name__ == '__main__':
    extract_func('welfare_officer_case_review')
