"""Render final decks and verify a hash-bound per-slide inspection record."""
import argparse
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
from forge import sha, write, read


def render(pptx, out, pdf=None):
    import fitz
    from PIL import Image, ImageOps, ImageDraw
    from pptx import Presentation
    out = Path(out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    pptx = Path(pptx).resolve()
    if not pdf:
        soffice = shutil.which('soffice')
        if not soffice:
            raise ValueError('LibreOffice unavailable; export this PPTX to PDF and supply --pdf. Visual review is incomplete.')
        with tempfile.TemporaryDirectory(prefix='forge-render-') as tmp:
            profile = (Path(tmp)/'profile').as_uri()
            subprocess.run([soffice, f'-env:UserInstallation={profile}', '--headless', '--convert-to', 'pdf', '--outdir', tmp, str(pptx)], check=True, timeout=120)
            generated = Path(tmp)/(pptx.stem+'.pdf')
            if not generated.is_file():
                raise ValueError('Renderer did not produce PDF')
            shutil.copy2(generated, out/'rendered.pdf')
        pdf = out/'rendered.pdf'
    with fitz.open(pdf) as doc:
        if len(doc) != len(Presentation(pptx).slides):
            raise ValueError('PDF/PPTX slide count differs')
        report = {'pptx_sha256': sha(pptx), 'pdf_sha256': sha(pdf), 'pdf_origin': 'Supplied export: verify that it was exported from this exact PPTX.' if Path(pdf).resolve() != out/'rendered.pdf' else 'LibreOffice export of final PPTX', 'slides': []}
        tiles = []
        for n,page in enumerate(doc, 1):
            dest = out/f'slide-{n:03}.png'
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(dest)
            report['slides'].append({'number': n, 'image': dest.name, 'sha256': sha(dest), 'checked': False, 'findings': [], 'resolved': False})
            with Image.open(dest) as im:
                tile = Image.new('RGB',(480, 300),'#eeeeee')
                tile.paste(ImageOps.contain(im,(480,270)),(0,0))
                ImageDraw.Draw(tile).text((10,280),str(n),fill='black')
                tiles.append(tile)
        columns = min(3, len(tiles))
        montage = Image.new('RGB', (480*columns,300*((len(tiles)+columns-1)//columns)), 'white')
        for i,tile in enumerate(tiles):
            montage.paste(tile,((i%columns)*480,(i//columns)*300))
        montage.save(out/'montage.png')
    write(out/'review.json',report)
    return {'rendered':len(report['slides']), 'review':str(out/'review.json'), 'status':'Awaiting actual per-slide inspection'}


def check(pptx, out):
    from pptx import Presentation
    out = Path(out)
    report = read(out/'review.json')
    errors = []
    if report['pptx_sha256'] != sha(pptx):
        errors.append('PPTX changed since rendering; re-render and inspect')
    slides = report.get('slides', [])
    count = len(Presentation(pptx).slides)
    if [s['number'] for s in slides] != list(range(1,count+1)):
        errors.append('Inspection coverage differs from final slide sequence')
    for s in slides:
        image = out/s['image']
        if not image.is_file() or sha(image) != s['sha256']:
            errors.append(f"Slide {s['number']}: preview missing or changed")
        if s.get('checked') is not True or s.get('resolved') is not True or not isinstance(s.get('findings'),list):
            errors.append(f"Slide {s['number']}: review incomplete")
    return {'passed':not errors, 'errors':errors}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('pptx')
    p.add_argument('out')
    p.add_argument('--pdf')
    p.add_argument('--check', action='store_true')
    a = p.parse_args()
    try:
        r = check(a.pptx,a.out) if a.check else render(a.pptx,a.out,a.pdf)
        print(json.dumps(r,indent=2))
        return 1 if r.get('passed') is False else 0
    except (ValueError, OSError, KeyError, subprocess.SubprocessError) as exc:
        print(f'ERROR: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
