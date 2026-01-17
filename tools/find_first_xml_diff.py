#!/usr/bin/env python3
"""Compare XML files in two directories using ElementTree and report the first semantic difference.

Usage: python3 tools/find_first_xml_diff.py <dir_a> <dir_b>
"""
import sys
import os
import xml.etree.ElementTree as ET


def list_files(root):
    out = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            rel = os.path.relpath(os.path.join(dirpath, fn), root)
            out.append(rel)
    return sorted(out)


def read_xml(path):
    try:
        text = open(path, 'rb').read()
    except Exception:
        return None
    try:
        # Parse with ElementTree
        root = ET.fromstring(text)
        return root
    except Exception:
        return None


def tag_localname(tag):
    if tag is None:
        return ''
    t = tag
    if t.startswith('{'):
        return t.split('}', 1)[1]
    return t


def sorted_attrib_list(elem):
    return sorted([(k, (v or '').strip()) for k, v in elem.attrib.items()])


def compare_elements(a, b, path='/'):
    # Compare tags (namespace-aware)
    if a.tag != b.tag:
        return f"Tag mismatch at {path}: {a.tag!r} != {b.tag!r}", a, b
    # Compare text
    a_text = (a.text or '').strip()
    b_text = (b.text or '').strip()
    if a_text != b_text:
        return f"Text mismatch at {path}: {a_text!r} != {b_text!r}", a, b
    # Compare attributes (sorted)
    a_attribs = sorted_attrib_list(a)
    b_attribs = sorted_attrib_list(b)
    if a_attribs != b_attribs:
        return f"Attributes mismatch at {path}: {a_attribs} != {b_attribs}", a, b
    # Compare children count
    a_children = [c for c in list(a) if isinstance(c.tag, str)]
    b_children = [c for c in list(b) if isinstance(c.tag, str)]
    if len(a_children) != len(b_children):
        return f"Different children count at {path}: {len(a_children)} != {len(b_children)}", a, b
    # Compare children in order
    for i, (ac, bc) in enumerate(zip(a_children, b_children)):
        child_name = tag_localname(ac.tag)
        subpath = f"{path}/{child_name}[{i}]"
        diff = compare_elements(ac, bc, subpath)
        if diff:
            return diff
    return None


def compare_files(path_a, path_b):
    # If both are XML, parse and compare
    a_xml = read_xml(path_a)
    b_xml = read_xml(path_b)
    if a_xml is None or b_xml is None:
        # fallback to raw bytes compare
        a_bytes = open(path_a, 'rb').read()
        b_bytes = open(path_b, 'rb').read()
        if a_bytes != b_bytes:
            return f"Binary files differ", a_bytes[:200], b_bytes[:200]
        return None
    # Compare element trees
    diff = compare_elements(a_xml, b_xml, path='')
    if diff:
        msg, a_elem, b_elem = diff
        return msg, ET.tostring(a_elem, encoding='unicode'), ET.tostring(b_elem, encoding='unicode')
    return None


def main(dir_a, dir_b):
    files_a = set(list_files(dir_a))
    files_b = set(list_files(dir_b))
    common = sorted(files_a & files_b)
    only_a = sorted(files_a - files_b)
    only_b = sorted(files_b - files_a)
    if only_a:
        print(f"Files only in {dir_a}: {only_a[:10]}{('...' if len(only_a)>10 else '')}")
    if only_b:
        print(f"Files only in {dir_b}: {only_b[:10]}{('...' if len(only_b)>10 else '')}")

    for rel in common:
        pa = os.path.join(dir_a, rel)
        pb = os.path.join(dir_b, rel)
        res = compare_files(pa, pb)
        if res is not None:
            print(f"\nFirst difference in file: {rel}")
            if isinstance(res, tuple) and len(res) == 3:
                msg, a_snip, b_snip = res
                print(msg)
                print('\n--- OURS snippet ---')
                print(a_snip)
                print('\n--- REPAIRED snippet ---')
                print(b_snip)
            else:
                print(res)
            return 1
    print("No differences found among common files (after basic XML normalization)")
    return 0


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    dir_a = sys.argv[1]
    dir_b = sys.argv[2]
    sys.exit(main(dir_a, dir_b))
