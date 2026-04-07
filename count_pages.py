import re, sys, io
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

filepath = r'C:\Users\bahaf\.claude\projects\D--PREPA-BTS-MEC\5155ef4a-bb6e-4561-9000-da2227c49864\tool-results\toolu_01Mrov8eU8J5ewfDC5xCmYiH.json'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# The file has JSON-escaped strings. Unescape properly
inner = content.replace('\\"', '"')

# Find all page_index values
pattern = r'"page_index":(\d+)'
indices = [int(m.group(1)) for m in re.finditer(pattern, inner)]

if indices:
    c = Counter(indices)
    for page in sorted(c.keys()):
        if page >= 23:
            print(f'Page {page}: {c[page]} elements')
    print(f'Max page in file: {max(indices)}')
else:
    print('No page indices found')
    # Debug: show escaping around page_index
    idx = content.find('page_index')
    if idx >= 0:
        print(f'Raw around page_index: {repr(content[idx-5:idx+30])}')
