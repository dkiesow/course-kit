import os
import zipfile
import xml.etree.ElementTree as ET
import pytest
import shutil


def test_titles_of_parts_head_is_aptos(tmp_path):
    p = '/tmp/pptx_wed_check/test_norm10_adj4.pptx'
    if not os.path.exists(p):
        pytest.skip('test artifact not present')
    cp = tmp_path / 'x.pptx'
    shutil.copy2(p, cp)

    # Normalize and inspect
    from pptx_builder import normalize_pptx
    normalize_pptx(str(cp))

    with zipfile.ZipFile(str(cp)) as zf:
        app_xml = zf.read('docProps/app.xml').decode('utf-8')
        tree = ET.fromstring(app_xml)
        ns = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        vt_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

        tops = tree.find(f'{{{ns}}}TitlesOfParts')
        assert tops is not None

        # find the first lpstr in the vector
        vec = None
        for c in tops:
            if c.tag.endswith('vector'):
                vec = c
                break
        assert vec is not None

        # collect all font lpstr entries (these appear first in the vector)
        fonts = [e.text.strip() for e in vec.findall('{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr') if e.text and e.text.strip()]
        low = [f.lower() for f in fonts]
        assert 'aptos' in low
        assert 'arial' in low
        assert 'helvetica neue' in low
        assert low.index('aptos') < low.index('arial') < low.index('helvetica neue')

        # Ensure 'PowerPoint Presentation' appears before the first real slide title (e.g., 'Entreprenuership and Intrapreneurship')
        all_lpstr = [e.text.strip() for e in vec.findall('{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpstr') if e.text and e.text.strip()]
        lower_all = [x.lower() for x in all_lpstr]
        if 'entreprenuership and intrapreneurship' in lower_all:
            assert 'powerpoint presentation' in lower_all
            assert lower_all.index('powerpoint presentation') < lower_all.index('entreprenuership and intrapreneurship')
