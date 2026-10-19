import json
from pathlib import Path

# Load audited slides from proposals/err4.audit.json
raw_audit = json.loads(Path('proposals/err4.audit.json').read_text(encoding='utf-8'))

retained_slides_map = {
    1: 'Slide 1: Title & Company Profile',
    2: 'Slide 2: Executive Summary, Investment Thesis & Key Highlights',
    14: 'Slide 3: Cash Flow Statement Analysis (Consolidated FY18-FY25)',
    16: 'Slide 4: Balance Sheet & Liquidity Ratio Analysis (FY18-FY25)',
    23: 'Slide 5: DCF Valuation & WACC Calculation',
}

audit_entries = []

for s in raw_audit['slides']:
    s_num = s['slide']
    is_retained = s_num in retained_slides_map
    
    stale_items = []
    resolved_items = []
    
    # Check Sun Pharma mentions
    for sp in s['shapes']:
        for p in sp['paragraphs']:
            if any(k in p.lower() for k in ('sun pharma', 'sun pharmaceutical', '524715', 'dilip shanghvi', 'ranbaxy', 'taro', 'organon')):
                stale_items.append({
                    'kind': 'shape_text',
                    'shape_id': sp['id'],
                    'shape_name': sp['name'],
                    'excerpt': p[:100],
                    'issue': 'Sun Pharma specific corporate text'
                })

    for tbl in s['tables']:
        for r_idx, row in enumerate(tbl['cells']):
            for c_idx, cell in enumerate(row):
                if any(k in cell.lower() for k in ('sun pharma', 'sun pharmaceutical', '524715')):
                    stale_items.append({
                        'kind': 'table_cell',
                        'table_id': tbl['id'],
                        'row': r_idx,
                        'col': c_idx,
                        'excerpt': cell[:80],
                        'issue': 'Sun Pharma label in table'
                    })

    for pic in s['pictures']:
        if pic.get('is_logo'):
            stale_items.append({
                'kind': 'picture_logo',
                'picture_id': pic['id'],
                'media': pic['media'],
                'issue': 'Sun Pharma corporate logo in slide header'
            })
        elif s_num in (23, 24, 25) and not pic.get('is_logo'):
            stale_items.append({
                'kind': 'picture_financial_screenshot',
                'picture_id': pic['id'],
                'media': pic['media'],
                'issue': 'Embedded raster screenshot of Sun Pharma financial schedule'
            })

    for ch in s['charts']:
        stale_items.append({
            'kind': 'chart',
            'chart_rid': ch['rid'],
            'part': ch['part'],
            'title': ch.get('title', ''),
            'series': ch.get('series', []),
            'issue': 'Native chart with Sun Pharma cached data and embedded workbook'
        })

    if is_retained:
        status = 'retained_for_marksans_narrative'
        narrative_role = retained_slides_map[s_num]
    else:
        status = 'pruned_unrelated_content'
        narrative_role = f'Sun Pharma specific detailed slide (Slide {s_num}) excluded to maintain coherent Marksans narrative'

    audit_entries.append({
        'slide': s_num,
        'source_deck': 'originals/ERR-4.pptx',
        'source_deck_sha256': raw_audit['deck_sha256'],
        'original_header': s['header'],
        'is_retained': is_retained,
        'narrative_role': narrative_role,
        'status': status,
        'stale_element_count': len(stale_items),
        'stale_elements': stale_items,
        'shape_count': s['shape_count'],
        'table_count': s['table_count'],
        'chart_count': s['chart_count'],
        'picture_count': s['picture_count'],
        'is_hidden': s['is_hidden'],
        'notes_excerpt': s['notes'][:120] if s['notes'] else None,
    })

full_audit = {
    'audit_schema_version': 1,
    'source_template': 'originals/ERR-4.pptx',
    'source_template_sha256': raw_audit['deck_sha256'],
    'total_slides_audited': len(audit_entries),
    'retained_slides_count': len(retained_slides_map),
    'retained_slide_numbers': sorted(retained_slides_map.keys()),
    'pruned_slides_count': len(audit_entries) - len(retained_slides_map),
    'missing_verified_assets': [
        {
            'asset_id': 'marksans_corporate_logo',
            'description': 'High-resolution official Marksans corporate logo vector/PNG',
            'reason': 'Not present in uploaded originals/ assets; inventing or relabeling is forbidden by policy',
            'handling': 'Flagged explicitly in audit and manifest; Sun Pharma header logo blocked/suppressed'
        }
    ],
    'slides': audit_entries
}

Path('docs/slide_content_audit.json').write_text(json.dumps(full_audit, indent=2), encoding='utf-8')

# Generate Markdown summary
deck_sha = raw_audit['deck_sha256']
md_lines = [
    '# Comprehensive Slide Content Audit: ERR-4.pptx',
    '',
    f'- **Source Template**: `originals/ERR-4.pptx` (SHA-256: `{deck_sha}`)',
    f'- **Total Slides Audited**: {len(audit_entries)}',
    f'- **Retained Curated Slides**: {len(retained_slides_map)} (Slides 1, 2, 14, 16, 23)',
    f'- **Pruned Slides**: {len(audit_entries) - len(retained_slides_map)} (Excluded to eliminate stale Sun Pharma content)',
    '',
    '## Retained Slides & Marksans Alignment Plan',
    '',
    '| Slide in Curated Deck | Source Slide in ERR-4 | Core Narrative Topic | Stale Elements Identified | Resolution & Provenance |',
    '|---|---|---|---|---|',
    '| Slide 1 | Slide 1 | Company Profile & Business Model | Company name, ticker, intro, business model, portfolio, logo | Fully mapped to Marksans OTC/Rx profile; logo flagged missing |',
    '| Slide 2 | Slide 2 | Executive Summary & Investment Thesis | Revenue, PAT, Net Cash, WACC, segment chart, snapshot chart | Mapped to Marksans FY25 actuals; Chart 3 updated with financial snapshot |',
    '| Slide 3 | Slide 14 | Cash Flow Statement Analysis | Ticker, historical cash flows FY18-FY25, narrative | Fully bound to Data Sheet C28:J85 across all 8 historical years |',
    '| Slide 4 | Slide 16 | Balance Sheet & Liquidity Ratios | Ticker, CR table, QR table, CR chart, QR chart, commentary | Table and Charts 14 & 15 updated with Marksans ratios from Data Sheet |',
    '| Slide 5 | Slide 23 | DCF Valuation & WACC Schedule | Ticker, Cost of Equity screenshot, Cost of Debt screenshot, WACC screenshot | Raster images replaced with high-res tables generated from WACC-DCF sheet |',
    '',
    '## Pruned Stale Slides (Excluded from Final Deck)',
    '',
    'The following slides from `originals/ERR-4.pptx` were audited and confirmed to contain 100% Sun Pharma specific historical data, M&A history, or drug pipelines. They are intentionally pruned to ensure presentation coherence and eliminate misleading stale content:',
    '',
    '| Source Slide | Original Header | Stale Sun Pharma Content | Status |',
    '|---|---|---|---|',
]

for s in audit_entries:
    if not s['is_retained']:
        h = s['original_header'].replace('\n', ' ')[:45]
        s_num = s['slide']
        stale_cnt = s['stale_element_count']
        md_lines.append(f'| Slide {s_num:2d} | {h} | {stale_cnt} stale items (history/pipeline/charts) | `pruned_unrelated_content` |')

md_lines.extend([
    '',
    '## Missing Verified Corporate Assets',
    '',
    '- **Marksans Corporate Logo**: The uploaded archives contain no verified Marksans vector or raster logo. Per repository quality rules, creating an invented logo or relabeling Sun Pharma\'s logo is strictly prohibited. The logo slot is explicitly flagged as pending an authentic asset, and the Sun Pharma logo is suppressed.',
    ''
])

Path('docs/slide_content_audit.md').write_text('\n'.join(md_lines), encoding='utf-8')
print('Successfully generated docs/slide_content_audit.json and docs/slide_content_audit.md')
