import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

filepath = r'C:\Users\bahaf\.claude\projects\D--PREPA-BTS-MEC\5155ef4a-bb6e-4561-9000-da2227c49864\tool-results\toolu_01Mrov8eU8J5ewfDC5xCmYiH.json'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print(f"File size: {len(content)}")

# Find all page indices
indices = re.findall(r'page_index\\\\":(\\d+)', content)
unique = sorted(set(int(x) for x in indices))
print(f"Page indices: {unique}")
print(f"Max page: {max(unique) if unique else 'none'}")
print()

# Unescape and extract elements for pages >= 29
inner = content.replace('\\"', '"')
pattern = r'"page_index":(\d+),"regions":\[(.*?)\],"element_id":"([^"]+)"'
matches = list(re.finditer(pattern, inner))
print(f"Total elements: {len(matches)}")

for m in matches:
    pi = int(m.group(1))
    if pi >= 29:
        eid = m.group(3)
        regions = m.group(2)
        texts = re.findall(r'"text":"((?:[^"\\]|\\.)*)"', regions)
        combined = ' '.join(texts)[:80].replace('\n', ' ')
        print(f'P{pi} | {eid} | {combined}')
