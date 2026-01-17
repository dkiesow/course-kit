import os
import shutil
import xml.etree.ElementTree as ET

from pptx_builder import _reorder_content_types


def test_prefer_repaired_override_order(tmp_path):
    rep_dir = '/tmp/pptx_wed_check/ex_repaired'
    try:
        os.makedirs(rep_dir, exist_ok=True)
        # repaired content types: docProps first then presentation
        rep_ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>
'''
        open(os.path.join(rep_dir, '[Content_Types].xml'), 'wb').write(rep_ct.encode('utf-8'))

        # our input has presentation first
        inp_ct = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'''
        out = _reorder_content_types(inp_ct.encode('utf-8'), files={})
        root = ET.fromstring(out.decode('utf-8'))
        ns = root.tag.split('}')[0].strip('{')
        overrides = [o.attrib.get('PartName') for o in root.findall(f'{{{ns}}}Override')]
        # Expect docProps/app to occur before ppt/presentation in the output
        assert overrides.index('/docProps/app.xml') < overrides.index('/ppt/presentation.xml')
    finally:
        # cleanup
        try:
            shutil.rmtree('/tmp/pptx_wed_check')
        except Exception:
            pass
