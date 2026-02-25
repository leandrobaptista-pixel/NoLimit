#!/usr/bin/env python3
import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INDEX = ROOT / 'index.html'
VIEWER = ROOT / 'viewer.html'
SCRIPT = ROOT / 'assets' / 'script.js'
MANIFEST = ROOT / 'assets' / 'gallery-manifest.json'


def read(path: Path) -> str:
    return path.read_text(encoding='utf-8', errors='replace')


def check_html_refs(path: Path):
    text = read(path)
    refs = re.findall(r'(?:href|src)="([^"]+)"', text)
    ids = set(re.findall(r'id="([^"]+)"', text))

    missing = []
    dead_anchors = []
    placeholder_links = []

    for ref in refs:
        ref = html.unescape(ref)
        if ref.startswith(('#', 'mailto:', 'tel:', 'javascript:', 'http://', 'https://')):
            if ref == '#':
                placeholder_links.append(ref)
            elif ref.startswith('#') and ref[1:] and ref[1:] not in ids:
                dead_anchors.append(ref)
            continue

        local = (ROOT / ref.split('?')[0]).resolve()
        if not local.exists():
            missing.append(ref)

    return missing, dead_anchors, placeholder_links


def check_js_manifest_refs(path: Path):
    text = read(path)
    refs = set(re.findall(r"'([^'\n]*assets/[^'\n]*)'", text))
    missing = []
    for ref in refs:
        if not (ROOT / ref).exists():
            missing.append(ref)
    return sorted(missing)


def check_manifest(path: Path):
    if not path.exists():
        return ['assets/gallery-manifest.json (missing)']
    text = read(path)
    refs = set(re.findall(r'"(assets/[^"]+)"', text))
    missing = []
    for ref in refs:
        if not (ROOT / ref).exists():
            missing.append(ref)
    return sorted(missing)


def main():
    issues = 0

    for file in [INDEX, VIEWER]:
        missing, dead_anchors, placeholders = check_html_refs(file)
        print(f'[{file.name}]')
        if missing:
            issues += len(missing)
            print('  Missing local refs:')
            for ref in missing:
                print(f'    - {ref}')
        if dead_anchors:
            issues += len(dead_anchors)
            print('  Dead anchor refs:')
            for ref in dead_anchors:
                print(f'    - {ref}')
        if placeholders:
            issues += len(placeholders)
            print('  Placeholder refs (#):')
            for ref in placeholders:
                print(f'    - {ref}')
        if not (missing or dead_anchors or placeholders):
            print('  OK')

    print(f'[{SCRIPT.relative_to(ROOT)}]')
    js_missing = check_js_manifest_refs(SCRIPT)
    if js_missing:
        issues += len(js_missing)
        print('  Missing local asset refs:')
        for ref in js_missing:
            print(f'    - {ref}')
    else:
        print('  OK')

    print(f'[{MANIFEST.relative_to(ROOT)}]')
    manifest_missing = check_manifest(MANIFEST)
    if manifest_missing:
        issues += len(manifest_missing)
        print('  Missing manifest assets:')
        for ref in manifest_missing:
            print(f'    - {ref}')
    else:
        print('  OK')

    if issues:
        print(f'\nResult: {issues} issue(s) found.')
        raise SystemExit(1)

    print('\nResult: all checks passed.')


if __name__ == '__main__':
    main()
