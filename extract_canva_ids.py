import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

filepath = r'C:\Users\bahaf\.claude\projects\D--PREPA-BTS-MEC\5155ef4a-bb6e-4561-9000-da2227c49864\tool-results\toolu_01RgLiC5knYvXyx3Day3rTH7.json'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# The JSON has escaped quotes as \\" in the inner JSON string
# We need to find page_index followed by element_id

# First, let's unescape the inner JSON
# Find the start of the inner JSON string
start = content.find('{"transaction"')
if start < 0:
    start = content.find('{\\\"transaction')

# Extract and unescape
inner = content[start:]
# Remove the JSON string escaping (replace \\" with " and \\\\ with \\)
inner = inner.replace('\\"', '"')

# Now parse with regex - find blocks between successive page_index
# Pattern: "page_index":N,...,"element_id":"XXX"
pattern = r'"page_index":(\d+),"regions":\[(.*?)\],"element_id":"([^"]+)"'
matches = list(re.finditer(pattern, inner))

print(f"Total elements found: {len(matches)}")
print()

for m in matches:
    pi = int(m.group(1))
    if pi >= 23:
        eid = m.group(3)
        regions = m.group(2)
        # Extract text values from regions
        texts = re.findall(r'"text":"((?:[^"\\]|\\.)*)"', regions)
        combined = ' '.join(texts)[:80].replace('\n', ' ')
        print(f'P{pi} | {eid} | {combined}')
