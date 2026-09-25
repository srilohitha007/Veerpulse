import ast
import astor

with open(r'c:\Users\SRI LOHITHA\OneDrive\Desktop\veerpulse\app\__init__.py', 'r', encoding='utf-8') as f:
    source = f.read()

tree = ast.parse(source)

for node in ast.walk(tree):
    if isinstance(node, ast.FunctionDef) and node.name == 'welfare_officer_case_review':
        print(astor.to_source(node))
