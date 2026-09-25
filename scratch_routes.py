import re

def extract_routes():
    with open(r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app\__init__.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Simple regex to find route and function name
    matches = re.finditer(r'@app\.route\(\s*[\"\']([^\"\']+)[\"\'].*?\n\s*def\s+([a-zA-Z0-9_]+)\(', content, re.DOTALL)
    for m in matches:
        print(f"{m.group(2)}: {m.group(1)}")

if __name__ == '__main__':
    extract_routes()
