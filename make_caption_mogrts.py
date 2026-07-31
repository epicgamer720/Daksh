#!/usr/bin/env python3
"""Stamp out pre-texted caption .mogrts from one template mogrt.

Premiere Pro 2026 removed getMGTComponent(), so no script can set mogrt text on
the timeline any more — and editing the live text layer's textEditValue renders
blank. The reliable way is to bake each caption INTO its own copy of the mogrt:
the text lives in a FlatBuffers blob (inside an Adobe 12-byte wrapper: u32 size,
u32 0, magic 44 33 22 11) in the mogrt's inner .prproj; the string sits at the
buffer's tail, so swapping it needs only the length prefix, padding, and the
wrapper size fixed up.

Usage:
  python3 make_caption_mogrts.py <template.mogrt> <captions> <out_dir>

<captions> is a cut-list spreadsheet (captions become "LABEL VS TEAM", in caps,
numbered by tape order) or a .txt file with one caption per line. Then use the
Clip Chopper Labels panel in Premiere: it places 'NN - CAPTION.mogrt' onto every
V1 clip named 'NN - ...'.
"""
import base64
import gzip
import io
import json
import re
import struct
import sys
import uuid
import zipfile
from pathlib import Path

PLACEHOLDER_RE = re.compile(
    r'(<StartKeyframeValue Encoding="base64"[^>]*>)([A-Za-z0-9+/=\s]+?)(</StartKeyframeValue>)')


def load_template(mogrt_path):
    """Return everything needed to stamp copies: parts, xml, blob location."""
    outer = zipfile.ZipFile(mogrt_path)
    parts = {n: outer.read(n) for n in outer.namelist()}
    inner = zipfile.ZipFile(io.BytesIO(parts['project.prgraphic']))
    prproj = [n for n in inner.namelist() if n.endswith('.prproj')][0]
    xml = gzip.decompress(inner.read(prproj)).decode('utf-8', errors='surrogateescape')
    definition = json.loads(parts['definition.json'].decode('utf-8'))
    # the editable text control's current value is the placeholder to hunt for
    placeholder = None
    for c in definition.get('clientControls', []):
        for s in c.get('value', {}).get('strDB', []):
            if s.get('str'):
                placeholder = s['str']
    if not placeholder:
        raise SystemExit('No text control with a value found in definition.json — '
                         'give the template text a distinctive placeholder first')
    span = raw = None
    for m in PLACEHOLDER_RE.finditer(xml):
        b = base64.b64decode(''.join(m.group(2).split()))
        if placeholder.encode() in b:
            span, raw = m.span(2), b
            break
    if not raw:
        raise SystemExit(f'Placeholder text {placeholder!r} not found in the mogrt payload')
    t = raw.find(placeholder.encode())
    if struct.unpack_from('<I', raw, t - 4)[0] != len(placeholder.encode()):
        raise SystemExit('Text length prefix not where expected — format changed?')
    if raw[8:12] != bytes.fromhex('44332211') or \
            struct.unpack_from('<I', raw, 0)[0] != len(raw) - 12:
        raise SystemExit('Adobe wrapper header not recognized — format changed?')
    if raw[t + len(placeholder.encode()):].strip(b'\x00') != b'':
        raise SystemExit('Bytes after the text are not padding — cannot patch safely')
    return parts, inner, prproj, xml, span, raw, t, definition


def stamp(parts, inner, prproj, xml, span, raw, t, definition, out_path, num, text):
    tb = text.encode('utf-8')
    body = bytearray(raw[:t - 4]) + struct.pack('<I', len(tb)) + tb + b'\x00'
    while len(body) % 4:
        body += b'\x00'
    struct.pack_into('<I', body, 0, len(body) - 12)
    new_xml = xml[:span[0]] + base64.b64encode(bytes(body)).decode() + xml[span[1]:]
    gz = gzip.compress(new_xml.encode('utf-8', errors='surrogateescape'))
    ibuf = io.BytesIO()
    with zipfile.ZipFile(ibuf, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in inner.namelist():
            z.writestr(n, gz if n == prproj else inner.read(n))
    d = json.loads(json.dumps(definition))
    d['capsuleID'] = str(uuid.uuid4())
    d['capsuleName'] = f'{num:02d} {text}'
    d['capsuleNameLocalized']['strDB'][0]['str'] = f'{num:02d} {text}'
    for c in d.get('clientControls', []):
        for s in c.get('value', {}).get('strDB', []):
            if s.get('str'):
                s['str'] = text
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            if n == 'definition.json':
                z.writestr(n, json.dumps(d, ensure_ascii=False))
            elif n == 'project.prgraphic':
                z.writestr(n, ibuf.getvalue())
            else:
                z.writestr(n, data)


def captions_from(source):
    p = Path(source)
    if p.suffix.lower() in ('.xlsx', '.csv'):
        from chopper import parse_sheet, label_text
        return [label_text(r, True).upper() for r in parse_sheet(p)]
    return [ln.strip().upper() for ln in p.read_text(encoding='utf-8').splitlines()
            if ln.strip()]


if __name__ == '__main__':
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    template, source, out_dir = sys.argv[1:]
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    tpl = load_template(template)
    caps = captions_from(source)
    for n, text in enumerate(caps, 1):
        stamp(*tpl, out / f'{n:02d} - {text.replace("/", "-")}.mogrt', n, text)
    print(f'{len(caps)} caption mogrts written to {out}')
