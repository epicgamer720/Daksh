#!/usr/bin/env python3
"""Self-check for chopper.py — run: py test_chopper.py"""

import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')
from chopper import (Row, build_xmeml, find_columns, list_videos, match_score,
                     match_videos, norm_tokens, parse_range, parse_sheet, parse_tc,
                     render_label_png, sanitize_filename)

HERE = Path(__file__).parent

# --- timecode grammar -------------------------------------------------------
assert parse_tc('14:18') == 858
assert parse_tc('1:00:10') == 3610
assert parse_tc('3:06') == 186
assert parse_tc('nope') is None
assert parse_range('52:10-52:17') == (3130, 3137)
assert parse_range('1:00:10 – 1:00:20') == (3610, 3620)   # en-dash + spaces
assert parse_range('Clip in Folder') is None
assert parse_range('') is None

# --- matching ---------------------------------------------------------------
assert norm_tokens('Georgetown Prep #2 (Daksh)') == ['georgetown', 'prep', '2']
assert norm_tokens('Taft (Daksh )') == ['taft']
s2 = match_score('Georgetown Prep #2 (Daksh)', 'Georgetown Prep 2')
s1 = match_score('Georgetown Prep #2 (Daksh)', 'Georgetown Prep 1')
assert s2 > s1 > 0, (s1, s2)
assert match_score('Bullis #1 (Daksh)', 'vs Bullis Game 1 Spring') > 0.55

# --- header detection on a differently-formatted sheet ----------------------
alt = [
    ['Some title junk', '', ''],
    ['Opponent', 'Timestamp', 'Description', 'Comments'],
    ['Riverside', '10:00-10:05', 'Goal', 'zoom in'],
    ['Lakeview', '1:02:03-1:02:10', 'Assist', ''],
]
hrow, cols = find_columns(alt)
assert hrow == 1
assert cols['game'] == 0 and cols['clip'] == 1 and cols['label'] == 2
assert cols['notes'] == [3]

# headerless sheet falls back to value-pattern detection
headerless = [
    ['Riverside', '10:00-10:05', 'Goal'],
    ['Lakeview', '11:00-11:08', 'Assist'],
    ['Riverside', '12:00-12:04', 'Goal'],
]
hrow, cols = find_columns(headerless)
assert hrow == -1 and cols['clip'] == 1 and cols.get('game') == 0

# a "Clip #" header must not steal the clip column from the real timecode column
steal = [
    ['Clip #', 'Game', 'Timecodes', 'Play'],
    ['1', 'Riverside', '10:00-10:05', 'Goal'],
    ['2', 'Lakeview', '11:00-11:04', 'Save'],
]
hrow, cols = find_columns(steal)
assert hrow == 0 and cols['clip'] == 2, cols

# unrecognized clip header ("Cut") is rescued by value pattern
cut = [
    ['Game', 'Cut', 'Notes'],
    ['Riverside', '10:00-10:05', ''],
    ['Lakeview', '11:00-11:04', ''],
]
hrow, cols = find_columns(cut)
assert hrow == 0 and cols['clip'] == 1, cols

# typo'd timecodes are flagged, never silently whole-file; real whole-file text still works
with tempfile.TemporaryDirectory() as td:
    p = Path(td) / 'cuts.csv'
    p.write_text('Game,Clip,Label\n'
                 'Riverside,10:00-10:05,Goal\n'
                 'Riverside,52.10-52.17,Assist\n'
                 'Lakeview,Clip in Folder,Save\n', encoding='utf-8')
    crows = parse_sheet(p)
assert crows[0].start == 600.0 and not crows[0].flags
assert not crows[1].whole_file and crows[1].flags and 'could not read' in crows[1].flags[0]
assert crows[2].whole_file and not crows[2].flags

# hidden files (macOS "._*" AppleDouble on USB drives, dot-dirs) are never videos
with tempfile.TemporaryDirectory() as td:
    (Path(td) / 'vs calvary.mp4').touch()
    (Path(td) / '._vs calvary.mp4').touch()
    (Path(td) / '.Trashes').mkdir()
    (Path(td) / '.Trashes' / 'old.mp4').touch()
    vids = list_videos(td)
    assert [v.name for v in vids] == ['vs calvary.mp4'], vids

# digit guard: "Game 2" must not match "Game 1.mp4" when Game 2's video is missing
with tempfile.TemporaryDirectory() as td:
    (Path(td) / 'Game 1.mp4').touch()
    row2 = Row(sheet_row=2, game='Game 2', start=1.0, end=2.0)
    match_videos([row2], td)
    assert row2.src is None and row2.flags, (row2.src, row2.flags)
    (Path(td) / 'Game 2.mp4').touch()
    row2b = Row(sheet_row=2, game='Game 2', start=1.0, end=2.0)
    match_videos([row2b], td)
    assert row2b.src and row2b.src.name == 'Game 2.mp4'
    # manual picks survive rematch
    manual = Row(sheet_row=3, game='Game 2', start=1.0, end=2.0,
                 manual=True, src=Path(td) / 'Game 1.mp4')
    match_videos([manual], td)
    assert manual.src.name == 'Game 1.mp4'

# --- the real spreadsheet ---------------------------------------------------
rows = parse_sheet(HERE / 'Spring_2026_Clips_sorted_1.xlsx')
assert len(rows) == 43, f'expected 43 clip rows, got {len(rows)}'
assert [r.order for r in rows] == list(range(1, 44))

r1 = rows[0]
assert 'georgetown prep' in r1.game.lower()
assert (r1.start, r1.end) == (3610.0, 3620.0)          # 1:00:10-1:00:20
assert r1.label == 'CTO (Army Commit)'
assert 'commentary' in r1.notes.lower()

whole = [r for r in rows if r.whole_file]
assert len(whole) == 7, f'expected 7 "Clip in Folder" rows, got {len(whole)}'

flagged = [r for r in rows if r.flags and not r.whole_file]
bad = next(r for r in rows if r.range_text == '3:06:00-3:14')
assert bad.flags and 'before start' in bad.flags[0], bad.flags

# notes from both Notes columns are merged
gilman = next(r for r in rows if r.order == 3)
assert 'Slide' in gilman.notes

# --- xmeml ------------------------------------------------------------------
fake = HERE / 'fake_video.mp4'
test_rows = [
    Row(sheet_row=2, game='Test Game', start=10.0, end=15.0, order=1,
        label='Goal', notes='speed up', src=fake),
    Row(sheet_row=3, game='Test Game', whole_file=True, order=2,
        label='Whole clip', src=fake),
]
probes = {str(fake): {'fps': 29.97, 'width': 1920, 'height': 1080,
                      'duration': 100.0, 'has_audio': True}}
xml = build_xmeml(test_rows, probes, 'Test Sequence')
root = ET.fromstring(xml)
assert root.tag == 'xmeml'
vitems = root.findall('.//video/track/clipitem')
aitems = root.findall('.//audio/track/clipitem')
assert len(vitems) == 2 and len(aitems) == 2
# clip 1: in/out at file fps (29.97), 10s-15s
assert vitems[0].find('in').text == str(round(10.0 * 29.97))
assert vitems[0].find('out').text == str(round(15.0 * 29.97))
assert vitems[0].find('start').text == '0'
# clip 2 = whole file, starts where clip 1 ended on the timeline
assert vitems[1].find('in').text == '0'
assert vitems[1].find('start').text == vitems[0].find('end').text
# ntsc rate
assert root.find('sequence/rate/timebase').text == '30'
assert root.find('sequence/rate/ntsc').text == 'TRUE'
# file defined once, referenced after
files = root.findall('.//file')
full_defs = [f for f in files if f.find('pathurl') is not None]
assert len(full_defs) == 1
assert 'file://localhost/' in full_defs[0].find('pathurl').text
# marker carries the note
marker = root.find('sequence/marker')
assert marker is not None and marker.find('comment').text == 'speed up'

assert sanitize_filename('CTO: "Army" <Commit>?') == 'CTO Army Commit'

# --- usage ping --------------------------------------------------------------
import http.server
import threading
import time
import chopper
from chopper import machine_fingerprint, sheet_fingerprint

g1 = [['Game', 'Clip'], ['A', '1:00-1:05']]
g2 = [['Game', 'Clip'], ['A', '1:00-1:06']]
assert sheet_fingerprint(g1) == sheet_fingerprint(g1)
assert sheet_fingerprint(g1) != sheet_fingerprint(g2)
assert len(machine_fingerprint()) == 12

captured = []
class _H(http.server.BaseHTTPRequestHandler):
    def do_POST(self):
        captured.append(self.rfile.read(int(self.headers['Content-Length'])).decode())
        self.send_response(200)
        self.end_headers()
    def log_message(self, *a):
        pass
srv = http.server.HTTPServer(('127.0.0.1', 0), _H)
threading.Thread(target=srv.serve_forever, daemon=True).start()

chopper.send_usage_ping('deadbeef')          # TRACK_URL empty -> must be a no-op
time.sleep(0.3)
assert not captured

chopper.TRACK_URL = f'http://127.0.0.1:{srv.server_port}/'
chopper.TRACK_FIELDS = {'sheet': 'entry.1', 'machine': 'entry.2'}
chopper.send_usage_ping('deadbeef')
for _ in range(30):
    if captured:
        break
    time.sleep(0.1)
assert captured and 'entry.1=deadbeef' in captured[0] and 'entry.2=' in captured[0], captured
chopper.TRACK_URL = ''
srv.shutdown()

# --- label overlays ----------------------------------------------------------
from PIL import Image
with tempfile.TemporaryDirectory() as td:
    png = Path(td) / 'lbl.png'
    render_label_png('CTO (Army Commit)', None, 48, 1920, 1080, png)
    im = Image.open(png)
    assert im.size == (1920, 1080) and im.mode == 'RGBA'
    assert im.getpixel((1900, 20))[3] == 0                    # top-right corner transparent
    assert im.crop((0, 700, 960, 1080)).getbbox() is not None  # content bottom-left

    # 4K sequence -> text scales up proportionally
    png4k = Path(td) / 'lbl4k.png'
    render_label_png('Goal', None, 48, 3840, 2160, png4k)
    assert Image.open(png4k).size == (3840, 2160)

# second video track with one still per labeled row, aligned to its clip
xml_l = build_xmeml(test_rows, probes, 'Seq', {0: Path('lbl.png')})
root_l = ET.fromstring(xml_l)
vtracks = root_l.findall('.//video/track')
assert len(vtracks) == 2
overlays = vtracks[1].findall('clipitem')
assert len(overlays) == 1                                     # only row 0 got a label png
base = vtracks[0].findall('clipitem')[0]
assert overlays[0].find('start').text == base.find('start').text
assert overlays[0].find('end').text == base.find('end').text
assert overlays[0].find('in').text == '0'
purl = overlays[0].find('file/pathurl').text
assert purl.startswith('file://localhost/') and purl.endswith('lbl.png')
# audio track untouched, and no second video track when no labels
assert len(root_l.findall('.//audio/track/clipitem')) == 2
assert len(ET.fromstring(build_xmeml(test_rows, probes, 'Seq')).findall('.//video/track')) == 1

print(f'ALL CHECKS PASSED — {len(rows)} rows parsed from the real sheet, '
      f'{len(whole)} whole-file clips, {len(flagged)} flagged for review')
