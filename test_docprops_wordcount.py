import os
import zipfile
import xml.etree.ElementTree as ET
import pytest
from pptx_builder import _compute_layout_master_overlap_subtract, normalize_pptx


def test_overlap_subtract_simple():
    files = {}
    # layout has 8 occurrences of 'alpha'
    files['ppt/slideLayouts/slideLayoutA.xml'] = b'<cSld name="MyLayout"><a:t>alpha alpha alpha alpha alpha alpha alpha alpha</a:t></cSld>'
    # master has 12 occurrences of 'alpha'
    files['ppt/slideMasters/slideMaster1.xml'] = b'<root><a:t>alpha alpha alpha alpha alpha alpha alpha alpha alpha alpha alpha alpha</a:t></root>'

    used_layout_names = {'MyLayout'}
    used_master_files = {'ppt/slideMasters/slideMaster1.xml'}

    # overlap_occurrences = min(8,12) = 8 -> subtract = 8 // 4 = 2
    assert _compute_layout_master_overlap_subtract(files, used_layout_names, used_master_files) == 2


def test_normalize_failing_if_present(tmp_path):
    p = '/tmp/pptx_wed_check/test_norm10_adj4.pptx'
    if not os.path.exists(p):
        pytest.skip('test artifact not present')
    cp = tmp_path / 'x.pptx'
    import shutil
    shutil.copy2(p, cp)
    normalize_pptx(str(cp))
    with zipfile.ZipFile(str(cp)) as zf:
        app_xml = zf.read('docProps/app.xml').decode('utf-8')
        tree = ET.fromstring(app_xml)
        ns = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        words = tree.find(f'{{{ns}}}Words')
        assert words is not None
        assert int(words.text) == 752, f'expected 752 words, got {words.text}'