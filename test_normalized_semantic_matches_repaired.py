import os
import shutil
import tempfile
import subprocess
import pytest


def test_normalized_semantic_matches_repaired_if_present():
    repaired_dir = '/tmp/pptx_wed_check/ex_repaired'
    artifact = '/tmp/pptx_wed_check/test_norm10_adj4.pptx'
    if not os.path.exists(repaired_dir) or not os.path.exists(artifact):
        pytest.skip('repaired artifact or test artifact not present')

    temp_pptx = '/tmp/pptx_wed_check/test_norm10_adj4_normalized_for_test.pptx'
    shutil.copy2(artifact, temp_pptx)

    from pptx_builder import normalize_pptx
    normalize_pptx(temp_pptx)

    # extract normalized pptx
    outdir = tempfile.mkdtemp(prefix='ex_test_norm_')
    subprocess.run(['unzip', '-q', temp_pptx, '-d', outdir], check=True)

    # run the XML comparator tool to find first semantic difference
    import tools.find_first_xml_diff as F
    rc = F.main(outdir, repaired_dir)
    assert rc == 0
