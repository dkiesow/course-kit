import os
import zipfile
import xml.etree.ElementTree as ET
import pytest
import shutil


def test_presentation_rels_match_repaired_if_present(tmp_path):
    repaired_rels = '/tmp/pptx_wed_check/ex_repaired/ppt/_rels/presentation.xml.rels'
    if not os.path.exists(repaired_rels):
        pytest.skip('repaired rels artifact not present')

    p = '/tmp/pptx_wed_check/test_norm10_adj4.pptx'
    if not os.path.exists(p):
        pytest.skip('test artifact not present')
    cp = tmp_path / 'x.pptx'
    shutil.copy2(p, cp)

    from pptx_builder import normalize_pptx
    normalize_pptx(str(cp))

    with zipfile.ZipFile(str(cp)) as zf:
        ours = zf.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
    repaired_text = open(repaired_rels, 'rb').read().decode('utf-8')

    ours_root = ET.fromstring(ours)
    repaired_root = ET.fromstring(repaired_text)

    ours_items = [(r.attrib.get('Target'), r.attrib.get('Id')) for r in list(ours_root)]
    repaired_items = [(r.attrib.get('Target'), r.attrib.get('Id')) for r in list(repaired_root)]
    assert repaired_items == ours_items
