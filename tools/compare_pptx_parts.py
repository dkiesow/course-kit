#!/usr/bin/env python3
"""Compare key PPTX parts structurally (docProps/app.xml, presentation rels, presentation.xml)"""
import sys
import os
import xml.etree.ElementTree as ET


def read_xml(path):
    try:
        with open(path, 'rb') as f:
            data = f.read()
        text = data.decode('utf-8')
        root = ET.fromstring(text)
        return root
    except Exception as e:
        print(f"ERROR reading/parsing {path}: {e}")
        return None


def qname_strip(tag):
    # returns (ns, localname)
    if tag is None:
        return (None, None)
    if tag.startswith('{'):
        ns, local = tag[1:].split('}', 1)
        return (ns, local)
    return (None, tag)


def find_child_text(root, localname, ns=None):
    for c in root:
        ns_c, local_c = qname_strip(c.tag)
        if local_c == localname and (ns is None or ns_c == ns):
            return (c.text or '').strip(), c
    return None, None


def parse_app(app_root):
    ns, _ = qname_strip(app_root.tag)
    vt_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

    slides_text, _ = find_child_text(app_root, 'Slides', ns)
    words_text, _ = find_child_text(app_root, 'Words', ns)
    paras_text, _ = find_child_text(app_root, 'Paragraphs', ns)

    # TitlesOfParts vector
    tops = None
    tops_list = []
    for c in app_root:
        ns_c, local_c = qname_strip(c.tag)
        if local_c == 'TitlesOfParts':
            # find vt:vector
            for cc in c:
                ns_cc, local_cc = qname_strip(cc.tag)
                if local_cc == 'vector':
                    for lp in cc:
                        ns_lp, local_lp = qname_strip(lp.tag)
                        if local_lp == 'lpstr':
                            tops_list.append((lp.text or '').strip())
    # HeadingPairs: extract pair counts
    hp_counts = {}
    for c in app_root:
        ns_c, local_c = qname_strip(c.tag)
        if local_c == 'HeadingPairs':
            for cc in c:
                ns_cc, local_cc = qname_strip(cc.tag)
                if local_cc == 'vector':
                    children = list(cc)
                    # children are variants alternating lpstr and i4
                    for i in range(0, len(children), 2):
                        lp = children[i]
                        i4 = children[i+1] if i+1 < len(children) else None
                        lp_text = ''
                        i4_text = ''
                        for x in lp:
                            ns_x, local_x = qname_strip(x.tag)
                            if local_x == 'lpstr':
                                lp_text = (x.text or '').strip()
                        if i4 is not None:
                            for y in i4:
                                ns_y, local_y = qname_strip(y.tag)
                                if local_y == 'i4':
                                    i4_text = (y.text or '').strip()
                        if lp_text:
                            hp_counts[lp_text] = i4_text
    return {
        'slides': slides_text,
        'words': words_text,
        'paragraphs': paras_text,
        'titles_of_parts': tops_list,
        'heading_pairs': hp_counts,
    }


def parse_presentation_slides(pres_root):
    # return ordered list of r:ids referenced in sldId list
    ns, _ = qname_strip(pres_root.tag)
    p_ns = ns
    sld_ids = []
    # find all sldId children
    for el in pres_root.findall('.//'):
        ns_el, local_el = qname_strip(el.tag)
        if local_el == 'sldId':
            rid = el.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or el.attrib.get('r:id') or el.attrib.get('id')
            sld_ids.append(rid)
    return sld_ids


def parse_rels_mapping(rels_root):
    mapping = {}
    for r in rels_root:
        ns_r, local_r = qname_strip(r.tag)
        if local_r == 'Relationship':
            rid = r.attrib.get('Id') or r.attrib.get('Id')
            rtype = r.attrib.get('Type')
            targ = r.attrib.get('Target')
            mapping[rid] = (rtype, targ)
    return mapping


def main(dir_a, dir_b):
    # docProps/app.xml
    app_a = read_xml(os.path.join(dir_a, 'docProps', 'app.xml'))
    app_b = read_xml(os.path.join(dir_b, 'docProps', 'app.xml'))
    if app_a is None or app_b is None:
        print('Skipping docProps/app.xml comparison (missing)')
    else:
        pa = parse_app(app_a)
        pb = parse_app(app_b)
        print('docProps/app.xml differences:')
        for key in ['slides', 'words', 'paragraphs']:
            if pa.get(key) != pb.get(key):
                print(f"  - {key}: ours={pa.get(key)!r} repaired={pb.get(key)!r}")
        # compare heading_pairs keys and values
        for k in sorted(set(list(pa.get('heading_pairs', {}).keys()) + list(pb.get('heading_pairs', {}).keys()))):
            av = pa['heading_pairs'].get(k)
            bv = pb['heading_pairs'].get(k)
            if av != bv:
                print(f"  - HeadingPairs[{k!r}]: ours={av!r} repaired={bv!r}")
        # TitlesOfParts list long-differ
        if pa.get('titles_of_parts') != pb.get('titles_of_parts'):
            print(f"  - TitlesOfParts length: ours={len(pa.get('titles_of_parts'))} repaired={len(pb.get('titles_of_parts'))}")
            # show head entries
            print('    ours titles head:', pa.get('titles_of_parts')[:8])
            print('    repaired titles head:', pb.get('titles_of_parts')[:8])

    # presentation.xml slides order vs mapping in rels
    pres_a = read_xml(os.path.join(dir_a, 'ppt', 'presentation.xml'))
    pres_b = read_xml(os.path.join(dir_b, 'ppt', 'presentation.xml'))

    rels_a = read_xml(os.path.join(dir_a, 'ppt', '_rels', 'presentation.xml.rels'))
    rels_b = read_xml(os.path.join(dir_b, 'ppt', '_rels', 'presentation.xml.rels'))

    if pres_a is None or pres_b is None or rels_a is None or rels_b is None:
        print('Skipping presentation rels comparison (missing)')
    else:
        sld_ids_a = parse_presentation_slides(pres_a)
        sld_ids_b = parse_presentation_slides(pres_b)
        map_a = parse_rels_mapping(rels_a)
        map_b = parse_rels_mapping(rels_b)
        # convert sld ids to targets sequence
        targets_a = [map_a.get(rid, (None, None))[1] for rid in sld_ids_a]
        targets_b = [map_b.get(rid, (None, None))[1] for rid in sld_ids_b]
        print('\nppt/presentation.xml slide targets sequences (first 10):')
        print('  ours:', targets_a[:10])
        print('  repaired:', targets_b[:10])
        if targets_a != targets_b:
            print('  -> slide order/target list differs!')
        else:
            print('  -> slide target lists match')

        # Check that set of slide targets are the same
        set_a = set([t for t in targets_a if t])
        set_b = set([t for t in targets_b if t])
        if set_a != set_b:
            print('  -> set of slide targets differs')
            print('    only in ours:', sorted(list(set_a - set_b))[:10])
            print('    only in repaired:', sorted(list(set_b - set_a))[:10])
        else:
            print('  -> slide target sets match (same files, possibly different rId numbering)')

    # check top-level rels (_rels/.rels)
    rels_root_a = read_xml(os.path.join(dir_a, '_rels', '.rels'))
    rels_root_b = read_xml(os.path.join(dir_b, '_rels', '.rels'))
    if rels_root_a is None or rels_root_b is None:
        print('Skipping top-level _rels/.rels comparison')
    else:
        ma = parse_rels_mapping(rels_root_a)
        mb = parse_rels_mapping(rels_root_b)
        if ma != mb:
            print('\n_rels/.rels mappings differ:')
            print('  ours:', ma)
            print('  repaired:', mb)
        else:
            print('\n_rels/.rels mappings match')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: compare_pptx_parts.py <dirA> <dirB>')
        sys.exit(2)
    dirA = sys.argv[1]
    dirB = sys.argv[2]
    main(dirA, dirB)
