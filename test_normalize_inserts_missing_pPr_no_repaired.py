import os
import shutil
import xml.etree.ElementTree as ET

from pptx_builder import _normalize_slide_paragraph_pPr


def test_inserts_missing_pPr_when_no_repaired(tmp_path):
    # ensure any debug repaired extraction is absent
    shutil.rmtree('/tmp/pptx_wed_check', ignore_errors=True)

    # minimal presentation with one slide
    rep_pres = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    rep_pres += '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
    rep_pres += '<p:sldIdLst><p:sldId id="1" r:id="rId1"/></p:sldIdLst>\n</p:presentation>'

    # slide: first paragraph already has a pPr, second does not
    our_slide = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    our_slide += '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n'
    our_slide += '<p:cSld><p:spTree>\n'
    our_slide += '<p:sp><p:txBody><a:bodyPr/><a:lstStyle/>'
    our_slide += '<a:p><a:pPr/><a:r><a:t>First</a:t></a:r></a:p>'
    our_slide += '<a:p><a:r><a:t>Second</a:t></a:r></a:p>'
    our_slide += '</p:txBody></p:sp></p:spTree></p:cSld></p:sld>'

    files = {}
    files['ppt/presentation.xml'] = rep_pres.encode('utf-8')
    files['ppt/slides/slide1.xml'] = our_slide.encode('utf-8')

    new_files = _normalize_slide_paragraph_pPr(files)
    out_slide = new_files['ppt/slides/slide1.xml'].decode('utf-8')

    root = ET.fromstring(out_slide)
    # find txBody by localname to avoid prefix/namespace serialization differences in tests
    tx = None
    for el in root.iter():
        if el.tag.split('}', 1)[-1] == 'txBody':
            tx = el
            break
    assert tx is not None, 'txBody not found in normalized slide XML'

    pars = [p for p in tx.iter() if p.tag.split('}', 1)[-1] == 'p']

    assert len(pars) == 2
    assert any(ch.tag.split('}', 1)[-1] == 'pPr' for ch in list(pars[0]))
    assert any(ch.tag.split('}', 1)[-1] == 'pPr' for ch in list(pars[1]))
