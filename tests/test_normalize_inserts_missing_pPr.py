import os
import shutil
from pptx_builder import _normalize_slide_paragraph_pPr


def make_presentation_xml(slide_count=1):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'>
            '<p:sldIdLst>' + ''.join([f'<p:sldId id="{100+i}" r:id="rId{i+1}"/>' for i in range(slide_count)]) + '</p:sldIdLst></p:presentation>')


def test_inserts_missing_pPr_for_repaired_paragraphs(tmp_path):
    rep_dir = '/tmp/pptx_wed_check/ex_repaired'
    try:
        # setup repaired extraction with presentation and slide
        os.makedirs(os.path.join(rep_dir, 'ppt', 'slides'), exist_ok=True)
        rep_pres = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        rep_pres += '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        rep_pres += '<p:sldIdLst><p:sldId id="1" r:id="rId1"/></p:sldIdLst>\n</p:presentation>'
        open(os.path.join(rep_dir, 'ppt', 'presentation.xml'), 'wb').write(rep_pres.encode('utf-8'))

        # repaired slide has pPr present in paragraphs
        rep_slide = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        rep_slide += '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        rep_slide += '<p:cSld><p:spTree>\n'
        rep_slide += '<p:sp><p:txBody><a:bodyPr/><a:lstStyle/>'
        rep_slide += '<a:p><a:r><a:t>First</a:t></a:r></a:p>'
        rep_slide += '<a:p><a:pPr/><a:r><a:t>Second</a:t></a:r></a:p>'
        rep_slide += '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
        open(os.path.join(rep_dir, 'ppt', 'slides', 'slide1.xml'), 'wb').write(rep_slide.encode('utf-8'))

        # our files mapping has a slide lacking pPr in second paragraph
        files = {}
        our_pres = rep_pres
        files['ppt/presentation.xml'] = our_pres.encode('utf-8')
        our_slide = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        our_slide += '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
        our_slide += '<p:cSld><p:spTree>\n'
        our_slide += '<p:sp><p:txBody><a:bodyPr/><a:lstStyle/>'
        our_slide += '<a:p><a:r><a:t>First</a:t></a:r></a:p>'
        our_slide += '<a:p><a:r><a:t>Second</a:t></a:r></a:p>'
        our_slide += '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'
        files['ppt/slides/slide1.xml'] = our_slide.encode('utf-8')

        # run normalization helper
        new_files = _normalize_slide_paragraph_pPr(files)
        out_slide = new_files['ppt/slides/slide1.xml'].decode('utf-8')

        # assert that a pPr element was inserted in second paragraph (namespace prefix may vary)
        import xml.etree.ElementTree as ET
        root = ET.fromstring(out_slide)
        assert any(el.tag.split('}', 1)[-1] == 'pPr' for el in root.iter())
    finally:
        try:
            shutil.rmtree('/tmp/pptx_wed_check')
        except Exception:
            pass
