"""Run isolated synthetic software fixtures; these are not research evidence."""
import copy
import json
from pathlib import Path
import tempfile
import zipfile
from forge import build, audit, validate, write, read
from render_review import render, check


def main():
    import fitz
    from PIL import Image
    from pptx import Presentation
    with tempfile.TemporaryDirectory(prefix='forge-selftest-') as tmp:
        root = Path(tmp)
        Image.new('RGB', (300,150), 'navy').save(root/'fixture.png')
        (root/'equation.xml').write_text('<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x = 1</m:t></m:r></m:oMath></m:oMathPara>',encoding='utf-8')
        source = {'id':'fixture','citation':'Synthetic software fixture','reference':'Synthetic self-test; not experimental data','locator':'selftest.py','verified':True}
        spec = {'title':'Tooling fixture','sources':[source],'slides':[
            {'title':'Fixture tests native elements','role':'result','seconds':60,'notes':'Synthetic values test software only.','citations':['fixture'],'elements':[
                {'kind':'chart','box':[.6,1.4,6,4.7],'x':[0,1,4], 'series':[{'name':'A','values':[1,2,3]},{'name':'B','values':[2,3,4]}],'x_label':'Fixture time (s)','y_label':'Fixture value (a.u.)','source_ids':['fixture']},
                {'kind':'image','box':[7,1.4,2,.8],'path':'fixture.png','source_ids':['fixture']},
                {'kind':'table','box':[7,2.5,4,1.2],'rows':[['Type','Value'],['Fixture','1']]},
                {'kind':'equation','box':[7,4,4,.5],'omml_path':'equation.xml','latex':'x=1'},
                {'kind':'equation','box':[7,5,4,.5],'text':'dX/dt = μX'}]},
            {'title':'Fixture tests a model diagram','role':'method','seconds':60,'notes':'Two native nodes.','citations':[],'elements':[
                {'kind':'diagram','box':[.7,1.7,11,4], 'nodes':[{'id':'input','label':'Measured input','box':[0,1,3,1]},{'id':'model','label':'Model','box':[6,1,3,1]}], 'edges':[{'from':'input','to':'model','label':'u(t)'}]}]}]}
        spec_path, pptx = root/'deck.json', root/'deck.pptx'
        write(spec_path,spec)
        build(spec_path,pptx)
        assert audit(spec_path,pptx)['passed']
        prs = Presentation(pptx)
        assert prs.slides[0].shapes[1].chart.series[0].values == (1.0,2.0,3.0)
        with zipfile.ZipFile(pptx) as z:
            assert 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx' in z.namelist()
            assert b'oMath' in z.read('ppt/slides/slide1.xml')
            assert b'cNvCxnSpPr' in z.read('ppt/slides/slide2.xml')
        for mutate in ('citation','nan','length','overflow','node','placeholder'):
            broken = copy.deepcopy(spec)
            chart = broken['slides'][0]['elements'][0]
            if mutate == 'citation': broken['sources'][0]['verified'] = False
            if mutate == 'nan': chart['series'][0]['values'][0] = float('nan')
            if mutate == 'length': chart['series'][1]['values'].pop()
            if mutate == 'overflow': chart['box'][0] = 20
            if mutate == 'node': broken['slides'][1]['elements'][0]['edges'][0]['to'] = 'absent'
            if mutate == 'placeholder': broken['slides'][0]['title'] = '[MISSING: result]'
            assert validate(broken,root), mutate
        # A deliberately made PDF exercises only the renderer/review machinery.
        # It is not a rendering of the PPTX and cannot validate PowerPoint appearance.
        pdf = root/'fixture.pdf'
        with fitz.open() as doc:
            for n in range(2):
                page = doc.new_page(width=960,height=540)
                page.insert_text((40,80),f'Synthetic renderer fixture {n+1}')
            doc.save(pdf)
        render(pptx,root/'review',pdf)
        assert not check(pptx,root/'review')['passed']
        report = read(root/'review'/'review.json')
        for slide in report['slides']:
            slide.update(checked=True,resolved=True)
        write(root/'review'/'review.json',report)
        assert check(pptx,root/'review')['passed']
        with open(pptx,'ab') as f: f.write(b'changed')
        assert not check(pptx,root/'review')['passed']
        print(json.dumps({'passed':True,'checks':['native chart with two series and irregular x','embedded editable workbook','native OMML','table/image/notes','connected editable diagram','six invalid-input gates','PDF preview generation','incomplete review rejection','stale PPTX review rejection'], 'limits':'Fixtures do not establish actual PPTX visual rendering or scientific truth.'},indent=2))


if __name__ == '__main__':
    main()
