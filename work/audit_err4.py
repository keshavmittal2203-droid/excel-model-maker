import zipfile
import io
import json
import re
import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from modelmaker import sha

pptx_bytes = Path('originals/ERR-4.pptx').read_bytes()
deck_hash = sha(pptx_bytes)
archive = zipfile.ZipFile(io.BytesIO(pptx_bytes))

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

pres_rels = ET.fromstring(archive.read('ppt/_rels/presentation.xml.rels'))
rel_map = {r.get('Id'): r.get('Target') for r in pres_rels}
pres_xml = ET.fromstring(archive.read('ppt/presentation.xml'))

slide_parts = []
for sld_id in pres_xml.findall(f'.//{{{P_NS}}}sldId'):
    rid = sld_id.get(f'{{{R_NS}}}id')
    slide_parts.append(posixpath.normpath(posixpath.join('ppt', rel_map[rid])))

slide_audit = []

for idx, s_part in enumerate(slide_parts, start=1):
    s_tree = ET.fromstring(archive.read(s_part))
    
    # Check slide rels
    rels_file = posixpath.join(posixpath.dirname(s_part), '_rels', posixpath.basename(s_part) + '.rels')
    s_rels = {}
    if rels_file in archive.namelist():
        for r in ET.fromstring(archive.read(rels_file)):
            target = r.get('Target')
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(s_part), target)) if not r.get('TargetMode') == 'External' else target
            s_rels[r.get('Id')] = {'target': resolved, 'type': r.get('Type')}

    # Check if hidden
    show = s_tree.get('show', '1')
    is_hidden = show == '0'

    # Notes
    notes_part = next((v['target'] for k, v in s_rels.items() if v['type'].endswith('/notesSlide')), None)
    notes_text = ''
    if notes_part and notes_part in archive.namelist():
        n_tree = ET.fromstring(archive.read(notes_part))
        notes_text = '\n'.join(''.join(t.text or '' for t in p.findall(f'.//{{{A_NS}}}t')) for p in n_tree.findall(f'.//{{{A_NS}}}p'))

    # Shapes, Text, Tables, Charts, Pictures
    shapes = []
    text_fragments = []
    has_sun_pharma = False
    
    for sp in s_tree.findall(f'.//{{{P_NS}}}sp'):
        cnv = sp.find(f'{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr')
        sid = int(cnv.get('id', '0')) if cnv is not None else 0
        sname = cnv.get('name', '') if cnv is not None else ''
        tx_body = sp.find(f'{{{P_NS}}}txBody')
        paras = []
        if tx_body is not None:
            for p in tx_body.findall(f'{{{A_NS}}}p'):
                txt = ''.join(t.text or '' for t in p.findall(f'.//{{{A_NS}}}t')).strip()
                if txt:
                    paras.append(txt)
                    text_fragments.append(txt)
                    if 'sun pharma' in txt.lower() or 'sun pharmaceutical' in txt.lower() or '524715' in txt:
                        has_sun_pharma = True
        if paras:
            shapes.append({'kind': 'sp', 'id': sid, 'name': sname, 'paragraphs': paras})

    tables = []
    for gf in s_tree.findall(f'.//{{{P_NS}}}graphicFrame'):
        cnv = gf.find(f'{{{P_NS}}}nvGraphicFramePr/{{{P_NS}}}cNvPr')
        gid = int(cnv.get('id', '0')) if cnv is not None else 0
        gname = cnv.get('name', '') if cnv is not None else ''
        tbl = gf.find(f'.//{{{A_NS}}}tbl')
        if tbl is not None:
            matrix = []
            for tr in tbl.findall(f'{{{A_NS}}}tr'):
                row_cells = []
                for tc in tr.findall(f'{{{A_NS}}}tc'):
                    cell_txt = '\n'.join(''.join(t.text or '' for t in p.findall(f'.//{{{A_NS}}}t')).strip() for p in tc.findall(f'.//{{{A_NS}}}p')).strip()
                    row_cells.append(cell_txt)
                    if 'sun pharma' in cell_txt.lower() or 'sun pharmaceutical' in cell_txt.lower() or '524715' in cell_txt:
                        has_sun_pharma = True
                matrix.append(row_cells)
            tables.append({'id': gid, 'name': gname, 'rows': len(matrix), 'cols': len(matrix[0]) if matrix else 0, 'cells': matrix})

    charts = []
    for gf in s_tree.findall(f'.//{{{P_NS}}}graphicFrame'):
        c_elem = gf.find(f'.//{{{C_NS}}}chart')
        if c_elem is not None:
            c_rid = c_elem.get(f'{{{R_NS}}}id')
            c_info = s_rels.get(c_rid, {})
            c_part = c_info.get('target')
            c_title = ''
            c_series = []
            c_embed = None
            if c_part and c_part in archive.namelist():
                c_tree = ET.fromstring(archive.read(c_part))
                t_node = c_tree.find(f'.//{{{C_NS}}}title')
                if t_node is not None:
                    c_title = ''.join(t.text or '' for t in t_node.findall(f'.//{{{A_NS}}}t'))
                for ser in c_tree.findall(f'.//{{{C_NS}}}ser'):
                    s_name = ''.join(t.text or '' for t in ser.findall(f'.//{{{C_NS}}}tx//{{{A_NS}}}t'))
                    c_series.append(s_name)
                # Check chart rels for embedded excel
                c_rels_part = posixpath.join(posixpath.dirname(c_part), '_rels', posixpath.basename(c_part) + '.rels')
                if c_rels_part in archive.namelist():
                    for cr in ET.fromstring(archive.read(c_rels_part)):
                        if cr.get('Type', '').endswith('/package'):
                            c_embed = posixpath.normpath(posixpath.join(posixpath.dirname(c_part), cr.get('Target')))
            charts.append({'rid': c_rid, 'part': c_part, 'title': c_title, 'series': c_series, 'embedded_workbook': c_embed})

    pictures = []
    for pic in s_tree.findall(f'.//{{{P_NS}}}pic'):
        cnv = pic.find(f'{{{P_NS}}}nvPicPr/{{{P_NS}}}cNvPr')
        pid = int(cnv.get('id', '0')) if cnv is not None else 0
        pname = cnv.get('name', '') if cnv is not None else ''
        blip = pic.find(f'.//{{{A_NS}}}blip')
        embed_rid = blip.get(f'{{{R_NS}}}embed') if blip is not None else None
        target_media = s_rels.get(embed_rid, {}).get('target') if embed_rid else None
        # xfrm
        xfrm = pic.find(f'.//{{{P_NS}}}spPr/{{{A_NS}}}xfrm')
        off = xfrm.find(f'{{{A_NS}}}off') if xfrm is not None else None
        ext = xfrm.find(f'{{{A_NS}}}ext') if xfrm is not None else None
        geom = {
            'x': int(off.get('x', '0')) if off is not None else 0,
            'y': int(off.get('y', '0')) if off is not None else 0,
            'cx': int(ext.get('cx', '0')) if ext is not None else 0,
            'cy': int(ext.get('cy', '0')) if ext is not None else 0,
        }
        # Check if logo by coordinates (top right)
        is_logo = (geom['x'] > 8000000 and geom['y'] < 1000000) or 'image2' in (target_media or '')
        pictures.append({'id': pid, 'name': pname, 'rid': embed_rid, 'media': target_media, 'geometry': geom, 'is_logo': is_logo})

    header_sp = next((sp['paragraphs'][0] for sp in shapes if sp['id'] in (7, 9, 11, 13, 14, 34, 6) and sp['paragraphs']), '')
    if not header_sp and shapes:
        header_sp = shapes[0]['paragraphs'][0] if shapes[0]['paragraphs'] else ''

    slide_audit.append({
        'slide': idx,
        'part': s_part,
        'header': header_sp,
        'is_hidden': is_hidden,
        'has_sun_pharma_mentions': has_sun_pharma,
        'shape_count': len(shapes),
        'table_count': len(tables),
        'chart_count': len(charts),
        'picture_count': len(pictures),
        'shapes': shapes,
        'tables': tables,
        'charts': charts,
        'pictures': pictures,
        'notes': notes_text.strip(),
    })

audit_report = {
    'deck_file': 'originals/ERR-4.pptx',
    'deck_sha256': deck_hash,
    'slide_count': len(slide_audit),
    'slides': slide_audit
}

Path('proposals/err4.audit.json').write_text(json.dumps(audit_report, indent=2))
print(f'Audited {len(slide_audit)} slides written to proposals/err4.audit.json')
for s in slide_audit:
    h = s['header'][:40]
    h_safe = h.encode('ascii', 'backslashreplace').decode('ascii')
    print(f"Slide {s['slide']:2d}: charts={s['chart_count']} tables={s['table_count']} pics={s['picture_count']} sun_pharma={s['has_sun_pharma_mentions']} | {h_safe}")
