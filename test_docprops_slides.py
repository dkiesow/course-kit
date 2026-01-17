import json
import os
import zipfile
from pptx import Presentation
from pptx_builder import build_pptx_from_slides


def test_docprops_slides_count(tmp_path):
    # Create 5 simple bullet slides
    bullets = json.dumps(["Bullet 1"])
    slides_data = []
    for i in range(5):
        slides_data.append((None, f"Headline {i+1}", None, bullets, None, None, None, False, False, False, False, 'bullets'))

    output_file = str(tmp_path / 'docprops_test.pptx')

    success = build_pptx_from_slides(
        slides_data=slides_data,
        output_path=output_file,
        template_path='templates/4734_template.potx',
        pptx_layouts_map=json.load(open('pptx_layouts.json'))
    )

    assert success
    assert os.path.exists(output_file)

    # python-pptx should load the file with the expected number of slides
    prs = Presentation(output_file)
    assert len(prs.slides) == 5

    # And docProps/app.xml Slides element should equal the number of slides
    with zipfile.ZipFile(output_file, 'r') as zf:
        assert 'docProps/app.xml' in zf.namelist()
        app_xml = zf.read('docProps/app.xml').decode('utf-8')
        import xml.etree.ElementTree as ET
        tree = ET.fromstring(app_xml)
        slides_el = tree.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Slides')
        assert slides_el is not None
        assert slides_el.text == '5'