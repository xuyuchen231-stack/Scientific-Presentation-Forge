"""Portable scientific deck ingestion, native authoring and structural inspection."""
import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import shutil
import sys


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write(path, value):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def validate(spec, base):
    errors = []
    sources = spec.get('sources', [])
    ids = [s.get('id') for s in sources]
    if len(set(ids)) != len(ids) or any(not i for i in ids):
        errors.append('Source IDs must be nonempty and unique')
    ledger = {s.get('id'): s for s in sources}
    if not spec.get('title') or not spec.get('slides'):
        errors.append('A deck title and at least one slide are required')
    if re.search(r'\[(MISSING|UNVERIFIED|TODO)\b', json.dumps(spec), re.I):
        errors.append('Unresolved content marker in specification')
    for n, slide in enumerate(spec.get('slides', []), 1):
        tag = f'Slide {n}'
        for key in ('title', 'role', 'notes', 'seconds', 'elements', 'citations'):
            if key not in slide:
                errors.append(f'{tag}: missing {key}')
        if not slide.get('title') or not slide.get('elements'):
            errors.append(f'{tag}: empty title/content')
        if not isinstance(slide.get('seconds'), (float, int)) or slide.get('seconds', 0) <= 0:
            errors.append(f'{tag}: positive duration required')
        cited = slide.get('citations', [])
        if slide.get('role') == 'result' and not cited:
            errors.append(f'{tag}: result requires evidence')
        for sid in cited:
            s = ledger.get(sid, {})
            if not all(s.get(k) for k in ('citation', 'reference', 'locator')) or s.get('verified') is not True:
                errors.append(f'{tag}: unverified/incomplete source {sid}')
        for e in slide.get('elements', []):
            box = e.get('box', [])
            if len(box) != 4 or not all(isinstance(v, (int, float)) and math.isfinite(v) for v in box):
                errors.append(f'{tag}: invalid element box')
                continue
            x, y, w, h = box
            if x < 0 or y < 0 or w <= 0 or h <= 0 or x+w > 13.334 or y+h > 7.501:
                errors.append(f'{tag}: element outside canvas')
            kind = e.get('kind')
            if kind not in ('text', 'table', 'chart', 'diagram', 'image', 'equation'):
                errors.append(f'{tag}: unsupported kind {kind}')
            if kind in ('chart', 'image') and not e.get('source_ids'):
                errors.append(f'{tag}: {kind} requires source_ids')
            if any(sid not in cited for sid in e.get('source_ids', [])):
                errors.append(f'{tag}: element evidence missing from visible citations')
            if kind == 'image' and not (base / e.get('path', '')).is_file():
                errors.append(f'{tag}: missing image')
            if kind == 'chart':
                xs = e.get('x', [])
                values = list(xs)
                series = e.get('series', [])
                if not values or not series:
                    errors.append(f'{tag}: empty chart')
                for series_item in series:
                    ys = series_item.get('values', [])
                    if len(ys) != len(xs):
                        errors.append(f'{tag}: unequal x/y lengths')
                    values = values + ys
                if any(not isinstance(v, (int, float)) or isinstance(v, bool) or not math.isfinite(v) for v in values):
                    errors.append(f'{tag}: non-finite or non-numeric chart data')
                if not e.get('x_label') or not e.get('y_label'):
                    errors.append(f'{tag}: chart axis labels required')
            if kind == 'table':
                rows = e.get('rows', [])
                if not rows or not rows[0] or any(len(r) != len(rows[0]) for r in rows):
                    errors.append(f'{tag}: empty or ragged table')
            if kind == 'equation':
                if e.get('omml_path'):
                    if not (base / e['omml_path']).is_file() or not e.get('latex'):
                        errors.append(f'{tag}: equation requires OMML file and LaTeX source')
                elif not e.get('text'):
                    errors.append(f'{tag}: empty equation')
            if kind == 'diagram':
                nodes = e.get('nodes', [])
                node_ids = [v.get('id') for v in nodes]
                if not nodes or len(set(node_ids)) != len(node_ids) or any(not v for v in node_ids):
                    errors.append(f'{tag}: invalid diagram nodes')
                for v in nodes:
                    b = v.get('box', [])
                    if len(b) != 4 or not all(isinstance(z, (int, float)) and math.isfinite(z) for z in b):
                        errors.append(f'{tag}: invalid node geometry')
                    elif b[0] < 0 or b[1] < 0 or min(b[2:]) <= 0 or b[0]+b[2] > w or b[1]+b[3] > h:
                        errors.append(f'{tag}: node outside diagram box')
                for edge in e.get('edges', []):
                    if edge.get('from') not in node_ids or edge.get('to') not in node_ids:
                        errors.append(f'{tag}: unknown edge endpoint')
    return errors


def build(spec_path, out):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.chart.data import XyChartData
    from pptx.oxml.xmlchemy import OxmlElement
    from lxml import etree
    spec = read(spec_path)
    base = Path(spec_path).resolve().parent
    errors = validate(spec, base)
    if errors:
        raise ValueError('\n'.join(errors))
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333333), Inches(7.5)
    font = spec.get('font', 'Times New Roman')
    ledger = {s['id']: s for s in spec.get('sources', [])}
    def text(slide, value, box, size=22, bold=False, color='17324D'):
        shape = slide.shapes.add_textbox(*(Inches(v) for v in box))
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.04)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        for i, line in enumerate(str(value).split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.name, p.font.size, p.font.bold = font, Pt(size), bold
            p.font.color.rgb = RGBColor.from_string(color)
        return shape
    for number, s in enumerate(spec['slides'], 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        text(slide, s['title'], [.5, .35, 12.3, .8], 30, True)
        for e in s['elements']:
            box = e['box']
            x, y, w, h = box
            kind = e['kind']
            if kind == 'text':
                text(slide, e.get('text', ''), box, e.get('size', 22), e.get('bold', False), e.get('color', '17324D'))
            elif kind == 'image':
                from PIL import Image
                path = base / e['path']
                with Image.open(path) as img:
                    iw, ih = img.size
                scale = min(w/iw, h/ih)
                pic = slide.shapes.add_picture(str(path), Inches(x+(w-iw*scale)/2), Inches(y+(h-ih*scale)/2), width=Inches(iw*scale), height=Inches(ih*scale))
                pic._element.nvPicPr.cNvPr.set('descr', e.get('description', ''))
            elif kind == 'table':
                rows = e['rows']
                table = slide.shapes.add_table(len(rows), len(rows[0]), *(Inches(v) for v in box)).table
                for i, row in enumerate(rows):
                    for j, value in enumerate(row):
                        cell = table.cell(i,j)
                        cell.text = str(value)
                        for p in cell.text_frame.paragraphs:
                            p.font.name, p.font.size = font, Pt(e.get('size', 18))
                            p.font.bold = i == 0
            elif kind == 'chart':
                data = XyChartData()
                for series in e['series']:
                    d = data.add_series(series['name'])
                    for xv, yv in zip(e['x'], series['values']):
                        d.add_data_point(xv, yv)
                chart = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES, *(Inches(v) for v in box), data).chart
                chart.has_legend = len(e['series']) > 1
                if chart.has_legend:
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                    chart.legend.font.size = Pt(16)
                for axis, label in ((chart.category_axis, e['x_label']), (chart.value_axis, e['y_label'])):
                    axis.has_title = True
                    axis.axis_title.text_frame.text = label
                    axis.tick_labels.font.size = Pt(16)
                    for p in axis.axis_title.text_frame.paragraphs:
                        p.font.name, p.font.size = font, Pt(18)
            elif kind == 'diagram':
                nodes = {}
                for node in e['nodes']:
                    nx, ny, nw, nh = node['box']
                    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x+nx), Inches(y+ny), Inches(nw), Inches(nh))
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = RGBColor.from_string('EDF3F7')
                    shape.line.color.rgb = RGBColor.from_string('17324D')
                    shape.text = node['label']
                    for p in shape.text_frame.paragraphs:
                        p.font.name, p.font.size = font, Pt(18)
                        p.font.color.rgb = RGBColor.from_string('17324D')
                    nodes[node['id']] = shape
                for edge in e.get('edges', []):
                    a, b = nodes[edge['from']], nodes[edge['to']]
                    # Connect native shape connection sites rather than covering node interiors.
                    rightward = b.left >= a.left + a.width
                    downward = b.top >= a.top + a.height
                    if rightward:
                        ax, ay, bx, by, ai, bi = a.left+a.width, a.top+a.height//2, b.left, b.top+b.height//2, 3, 1
                    elif downward:
                        ax, ay, bx, by, ai, bi = a.left+a.width//2, a.top+a.height, b.left+b.width//2, b.top, 2, 0
                    else:
                        ax, ay, bx, by, ai, bi = a.left, a.top+a.height//2, b.left+b.width, b.top+b.height//2, 1, 3
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, ax, ay, bx, by)
                    line.begin_connect(a, ai)
                    line.end_connect(b, bi)
                    tail = OxmlElement('a:tailEnd')
                    tail.set('type', 'triangle')
                    line._element.spPr.get_or_add_ln().append(tail)
                    if edge.get('label'):
                        text(slide, edge['label'], [(ax+bx)/2/914400-.6, (ay+by)/2/914400-.3, 1.2, .3], 14)
            elif kind == 'equation':
                if e.get('omml_path'):
                    root = etree.fromstring((base/e['omml_path']).read_bytes(), etree.XMLParser(resolve_entities=False, no_network=True))
                    if root.tag != '{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara':
                        raise ValueError('OMML root must be m:oMathPara')
                    shape = text(slide, '', box)
                    para = shape.text_frame.paragraphs[0]._p
                    wrapper = etree.Element('{http://schemas.microsoft.com/office/drawing/2010/main}m', nsmap={'a14':'http://schemas.microsoft.com/office/drawing/2010/main'})
                    wrapper.append(root)
                    para.append(wrapper)
                else:
                    text(slide, e['text'], box, e.get('size', 26))
        text(slide, ' | '.join(ledger[v]['citation'] for v in s.get('citations', [])), [.5, 6.98, 11.6, .3], 12, color='4B5563')
        text(slide, str(number), [12.3, 6.98, .5, .3], 12)
        slide.notes_slide.notes_text_frame.text = s.get('notes', '')
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return {'pptx': str(Path(out).resolve()), 'slides': len(prs.slides), 'sha256': sha(out)}


def audit(spec_path, pptx_path):
    from pptx import Presentation
    spec = read(spec_path)
    errors = validate(spec, Path(spec_path).resolve().parent)
    prs = Presentation(pptx_path)
    warnings = []
    if len(prs.slides) != len(spec['slides']):
        errors.append('PPTX/spec slide count mismatch')
    for i, slide in enumerate(prs.slides, 1):
        if i <= len(spec['slides']):
            expected = spec['slides'][i-1]
            texts = '\n'.join(s.text for s in slide.shapes if s.has_text_frame)
            if expected['title'] not in texts:
                errors.append(f'Slide {i}: title differs from specification')
            if slide.notes_slide.notes_text_frame.text != expected.get('notes', ''):
                errors.append(f'Slide {i}: notes differ from specification')
            ledger = {s['id']:s for s in spec.get('sources', [])}
            for sid in expected.get('citations', []):
                if ledger.get(sid, {}).get('citation', '\x00') not in texts:
                    errors.append(f'Slide {i}: visible citation missing for {sid}')
        for shape in slide.shapes:
            if shape.left < -9144 or shape.top < -9144 or shape.left+shape.width > prs.slide_width+9144 or shape.top+shape.height > prs.slide_height+9144:
                errors.append(f'Slide {i}: shape outside canvas')
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if run.font.size and run.font.size.pt < 12:
                            warnings.append(f'Slide {i}: text below 12 pt')
    return {'passed': not errors, 'errors': errors, 'warnings': warnings, 'pptx_sha256': sha(pptx_path), 'scope': 'Structural only; scientific truth and rendered layout require inspection.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('doctor')
    for cmd in ('ingest', 'build', 'audit'):
        p = sub.add_parser(cmd)
        p.add_argument('input')
        p.add_argument('output')
        if cmd == 'audit':
            p.add_argument('--out', required=True)
    args = parser.parse_args()
    if args.command == 'doctor':
        print(json.dumps({'python': sys.version.split()[0], 'libraries': {m: importlib.util.find_spec(m) is not None for m in ('pptx','fitz','PIL')}, 'soffice': shutil.which('soffice')}, indent=2))
    elif args.command == 'ingest':
        import fitz
        with fitz.open(args.input) as doc:
            pages = [{'page': i+1, 'text': page.get_text(), 'images': [{'bbox': list(info['bbox']), 'width': info['width'], 'height': info['height']} for info in page.get_image_info()]} for i,page in enumerate(doc)]
        write(args.output, {'source': str(Path(args.input).resolve()), 'sha256': sha(args.input), 'pages': pages})
    elif args.command == 'build':
        print(json.dumps(build(args.input, args.output), indent=2))
    elif args.command == 'audit':
        result = audit(args.input, args.output)
        write(args.out, result)
        print(json.dumps(result, indent=2))
        return 0 if result['passed'] else 1
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except (ValueError, KeyError, OSError) as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(1)
