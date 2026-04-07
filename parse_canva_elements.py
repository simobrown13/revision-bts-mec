# -*- coding: utf-8 -*-
import json, sys, io, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

JSON_PATH = 'C:/Users/bahaf/.claude/projects/D--PREPA-BTS-MEC/5155ef4a-bb6e-4561-9000-da2227c49864/tool-results/toolu_015WR11WwboQKghmR3VwrDaT.json'

def extract_richtexts_regex(raw_text):
    # Use a simple approach: split on element boundaries
    elements = []
    # Find all page_index...element_id blocks
    # Pattern: {"page_index":N,"regions":[...],"element_id":"..."}
    # Split by the start pattern
    parts = re.split(r'\{"page_index":', raw_text)
    for part in parts[1:]:  # skip first empty/header part
        # Extract page_index
        pi_match = re.match(r'(\d+)', part)
        if not pi_match:
            continue
        page_index = int(pi_match.group(1))
        # Extract element_id
        eid_match = re.search(r'"element_id":"([^"]+)"', part)
        if not eid_match:
            continue
        element_id = eid_match.group(1)
        # Extract all text values from regions
        # Simple: find all "text":"..." pairs
        text_parts = []
        for tm in re.finditer(r'"text":"((?:[^"]*)*)"', part):
            t = tm.group(1)
            # Unescape JSON string escapes
            bs = chr(92)
            t = t.replace(bs + 'n', chr(10))
            t = t.replace(bs + 't', chr(9))
            t = t.replace(bs + '"', '"')
            t = t.replace(bs + bs, bs)
            text_parts.append(t)
        combined = ''.join(text_parts)
        elements.append(dict(page_index=page_index, element_id=element_id, text=combined))
    return elements

def main():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        wrapper = json.load(f)

    raw_text = ''
    for item in wrapper:
        if item.get('type') == 'text':
            raw_text += item.get('text', '')

    # Try direct JSON parse first
    data = None
    try:
        data = json.loads(raw_text, strict=False)
    except json.JSONDecodeError:
        pass

    elements = []
    truncated = False

    if data and isinstance(data, dict) and 'richtexts' in data:
        for rt in data['richtexts']:
            page_index = rt.get('page_index', -1)
            element_id = rt.get('element_id', 'unknown')
            regions = rt.get('regions', [])
            combined_text = ''
            for region in regions:
                combined_text += region.get('text', '')
            elements.append(dict(page_index=page_index, element_id=element_id, text=combined_text))
    else:
        truncated = True
        print('NOTE: JSON may be truncated. Using regex extraction.')
        print()
        elements = extract_richtexts_regex(raw_text)

    if not elements:
        print('ERROR: No richtext elements found!')
        return

    pages = {}
    for el in elements:
        pi = el['page_index']
        if pi not in pages:
            pages[pi] = []
        pages[pi].append(el)

    total_elements = len(elements)
    total_pages = len(pages)
    sep = '=' * 120

    print(sep)
    print('CANVA RICHTEXT ELEMENTS - FULL EXTRACTION')
    if truncated:
        print('(Data may be incomplete due to truncation)')
    print(sep)
    print()

    for page_num in sorted(pages.keys()):
        page_elements = pages[page_num]
        print(sep)
        msg = 'PAGE %d (%d elements)' % (page_num, len(page_elements))
        print(msg)
        print(sep)
        for i, el in enumerate(page_elements, 1):
            text_preview = el['text'].replace(chr(10), ' | ').strip()
            if len(text_preview) > 100:
                text_preview = text_preview[:100] + '...'
            eid = el['element_id']
            print('  [%02d] %s' % (i, eid))
            print('       Text: %s' % text_preview)
            print()
        print()

    print(sep)
    print('SUMMARY')
    print(sep)
    print('  Total pages:    %d' % total_pages)
    print('  Page range:     %d - %d' % (min(pages.keys()), max(pages.keys())))
    print('  Total elements: %d' % total_elements)
    print()
    print('  Elements per page:')
    for page_num in sorted(pages.keys()):
        count = len(pages[page_num])
        bar = '#' * count
        print('    Page %3d: %3d elements  %s' % (page_num, count, bar))
    print()

    print(sep)
    print('LONG TEXT ELEMENTS (>200 chars) - FULL CONTENT')
    print(sep)
    long_count = 0
    for page_num in sorted(pages.keys()):
        for el in pages[page_num]:
            if len(el['text']) > 200:
                long_count += 1
                eid = el['element_id']
                tlen = len(el['text'])
                print()
                print('--- Page %d | %s (%d chars) ---' % (page_num, eid, tlen))
                print(el['text'])
                print()
    if long_count == 0:
        print('  (none found)')
    print()
    print('  Total long elements: %d' % long_count)

if __name__ == '__main__':
    main()
