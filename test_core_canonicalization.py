import os
import zipfile
import xml.etree.ElementTree as ET
import pytest
import shutil


def test_core_revision_increments_when_app_changed(tmp_path):
    p = '/tmp/pptx_wed_check/test_norm10_adj4.pptx'
    if not os.path.exists(p):
        pytest.skip('test artifact not present')
    cp = tmp_path / 'x.pptx'
    shutil.copy2(p, cp)

    # read current core revision
    with zipfile.ZipFile(str(cp)) as zf:
        core_xml = zf.read('docProps/core.xml').decode('utf-8')
        tree = ET.fromstring(core_xml)
        cp_ns = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        rev = tree.find(f'{{{cp_ns}}}revision')
        old_rev = int(rev.text) if rev is not None and rev.text and rev.text.isdigit() else 0

    from pptx_builder import normalize_pptx
    normalize_pptx(str(cp))

    with zipfile.ZipFile(str(cp)) as zf:
        core_xml = zf.read('docProps/core.xml').decode('utf-8')
        tree = ET.fromstring(core_xml)
        rev = tree.find(f'{{{cp_ns}}}revision')
        assert rev is not None
        # If a repaired copy exists in the debug folder, match it exactly; otherwise assert +1 behavior
        repaired_core = '/tmp/pptx_wed_check/ex_repaired/docProps/core.xml'
        if os.path.exists(repaired_core):
            rtxt = open(repaired_core, 'rb').read().decode('utf-8')
            rroot = ET.fromstring(rtxt)
            rrev = rroot.find(f'{{{cp_ns}}}revision')
            if rrev is not None and rrev.text and rrev.text.strip().isdigit():
                assert int(rev.text) == int(rrev.text.strip())
            else:
                assert int(rev.text) == old_rev + 1
        else:
            assert int(rev.text) == old_rev + 1
