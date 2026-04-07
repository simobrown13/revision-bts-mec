import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

filepath = r'C:\Users\bahaf\.claude\projects\D--PREPA-BTS-MEC\5155ef4a-bb6e-4561-9000-da2227c49864\tool-results\toolu_015WR11WwboQKghmR3VwrDaT.json'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

inner = content.replace('\\"', '"')
pattern = r'"page_index":(\d+),"regions":\[(.*?)\],"element_id":"([^"]+)"'
matches = list(re.finditer(pattern, inner))

target_pages = [9, 10, 20, 25, 26, 27, 28, 29]
for m in matches:
    pi = int(m.group(1))
    if pi in target_pages:
        eid = m.group(3)
        regions = m.group(2)
        texts = re.findall(r'"text":"((?:[^"\\]|\\.)*)"', regions)
        combined = ' '.join(texts)[:120].replace('\n', ' ')
        print(f'P{pi} | {eid} | {combined}')
