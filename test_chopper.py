#!/usr/bin/env python3
"""Self-check for chopper.py — run: py test_chopper.py"""

import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')
from chopper import (Row, build_xmeml, clean_team, find_columns, label_text,
                     list_videos, match_score, match_videos, norm_tokens,
                     opponent_tokens, parse_range, parse_sheet, parse_tc,
                     place_kind, render_label_png, sanitize_filename,
                     timeline_layout)

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

# Game film is "<us> vs <them>" and the sheet names <them>, so our own team is noise
# that every file in the library repeats — and noise that scores.
assert opponent_tokens(norm_tokens('29s 3D NE Red vs 91 georgia')) == ['91', 'georgia']
assert opponent_tokens(norm_tokens('vs stealth')) == ['stealth']
assert opponent_tokens(norm_tokens('sweetlax upstate')) == ['sweetlax', 'upstate']  # no "vs"
assert opponent_tokens(norm_tokens('3d 29 vs')) == ['3d', '29', 'vs']  # nothing after it
# "3D Georgia" took BOTH its words off a tape of a different club — "3d" from our own
# name, "georgia" from "91 georgia" — and beat its own game by 0.006.
right = match_score('3D Georgia', '3d NE Red 2029 vs 3d Georgia')
wrong = match_score('3D Georgia', '29s 3D NE Red vs 91 georgia')
assert right - wrong > 0.1, (right, wrong)

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

# nested tournament folders: the folder names take part in matching, and several
# root folders can be searched at once (deduped when they overlap)
with tempfile.TemporaryDirectory() as td:
    root_a = Path(td) / 'Next level summer'
    (root_a / 'Alliance' / '2027').mkdir(parents=True)
    (root_a / 'Alliance' / '2027' / 'vs calvary.mp4').touch()
    (root_a / 'Rumble' / '2027').mkdir(parents=True)
    (root_a / 'Rumble' / '2027' / 'vs calvary.mp4').touch()
    root_b = Path(td) / 'Spring tapes'
    root_b.mkdir()
    (root_b / 'vs Landon.mp4').touch()

    assert len(list_videos(root_a)) == 2                    # recurses into tournaments
    assert len(list_videos([root_a, root_b])) == 3          # several roots
    assert len(list_videos([root_a, root_a / 'Alliance'])) == 2  # overlap deduped

    # identical filenames in two tournaments -> the folder name decides
    r_all = Row(sheet_row=2, game='Alliance vs Calvary', start=1.0, end=2.0)
    r_rum = Row(sheet_row=3, game='Rumble vs Calvary', start=1.0, end=2.0)
    match_videos([r_all, r_rum], [root_a, root_b])
    assert r_all.src and r_all.src.parent.parent.name == 'Alliance', r_all.src
    assert r_rum.src and r_rum.src.parent.parent.name == 'Rumble', r_rum.src
    # a game found only in the second root still matches
    r_lan = Row(sheet_row=4, game='Landon', start=1.0, end=2.0)
    match_videos([r_lan], [root_a, root_b])
    assert r_lan.src and r_lan.src.name == 'vs Landon.mp4', r_lan.src

    # naming a tournament pins the match to that folder even when only one root is given
    r_only = Row(sheet_row=8, game='Rumble vs Calvary', start=1.0, end=2.0)
    match_videos([r_only], root_a)
    assert r_only.src and r_only.src.parent.parent.name == 'Rumble', r_only.src
    # a game naming no known folder is unaffected by the folder rule
    r_plain = Row(sheet_row=9, game='Calvary', start=1.0, end=2.0)
    match_videos([r_plain], [root_a, root_b])
    assert r_plain.src and r_plain.src.name == 'vs calvary.mp4', r_plain.src
    # digit identity still holds across folders: year folder is not a game number
    r_two = Row(sheet_row=5, game='Alliance #2', start=1.0, end=2.0)
    match_videos([r_two], [root_a, root_b])
    assert r_two.src is None, r_two.src

    # a tape NOT filed under its tournament folder must still win over a same-named
    # file that is (folder evidence must never hard-filter the right answer away)
    (root_a / 'Rumble vs Landon.mp4').touch()
    r_loose = Row(sheet_row=10, game='Rumble vs Landon', start=1.0, end=2.0)
    match_videos([r_loose], root_a)
    assert r_loose.src and r_loose.src.name == 'Rumble vs Landon.mp4', r_loose.src
    assert not r_loose.flags, r_loose.flags
    (root_a / 'Rumble vs Landon.mp4').unlink()

    # adding a subfolder BEFORE its parent must not strip that folder's name
    for order in ([root_a / 'Rumble', root_a], [root_a, root_a / 'Rumble']):
        r_ord = Row(sheet_row=11, game='Rumble vs Calvary', start=1.0, end=2.0)
        match_videos([r_ord], order)
        assert r_ord.src and r_ord.src.parent.parent.name == 'Rumble', (order, r_ord.src)

    # a numbered folder that isn't the game ("Day 2") can't satisfy a game number
    day = Path(td) / 'Summer'
    (day / 'Day 2').mkdir(parents=True)
    (day / 'Day 2' / 'Georgetown Prep 1.mp4').touch()
    r_gp2 = Row(sheet_row=12, game='Georgetown Prep #2', start=1.0, end=2.0)
    match_videos([r_gp2], day)
    assert r_gp2.src is None, r_gp2.src
    # ...but a folder that names the game carries its number just fine
    (day / 'Georgetown Prep 2' / '2026').mkdir(parents=True)
    (day / 'Georgetown Prep 2' / '2026' / 'full game.mp4').touch()
    r_gp2b = Row(sheet_row=13, game='Georgetown Prep #2', start=1.0, end=2.0)
    match_videos([r_gp2b], day)
    assert r_gp2b.src and r_gp2b.src.parent.parent.name == 'Georgetown Prep 2', r_gp2b.src
    # and an unrelated numbered folder doesn't hide a correctly-named flat file
    (day / 'Georgetown Prep 3.mp4').touch()
    r_gp3 = Row(sheet_row=14, game='Georgetown Prep #3', start=1.0, end=2.0)
    match_videos([r_gp3], day)
    assert r_gp3.src and r_gp3.src.name == 'Georgetown Prep 3.mp4', r_gp3.src

    # a pre-scanned file list is reused as-is, and `exclude` still applies to it
    scanned = list_videos([root_a, root_b])
    r_pre = Row(sheet_row=6, game='Alliance vs Calvary', start=1.0, end=2.0)
    match_videos([r_pre], [root_a, root_b], videos=scanned)
    assert r_pre.src and r_pre.src.parent.parent.name == 'Alliance'
    r_ex = Row(sheet_row=7, game='Alliance vs Calvary', start=1.0, end=2.0)
    match_videos([r_ex], [root_a, root_b], videos=scanned, exclude=root_a / 'Alliance' / '2027')
    assert r_ex.src != root_a / 'Alliance' / '2027' / 'vs calvary.mp4', r_ex.src

# the summer '29 library: our own team's name is in every file, so the sheet's words can
# all land on a tape of a different club (real misses, "Final summer highlight" sheet)
with tempfile.TemporaryDirectory() as td:
    summer = Path(td) / "3D '29 summer"
    for folder, names in {
        'LI showdown': ['29s 3D NE Red vs 91 georgia', '29s 3D NE vs Hilltop',
                        '29s 3D NE Red vs leading edge', '29s 3D NE Maddog'],
        'Sweetlax': ['3d NE Red 2029 vs 3d Georgia', '3d NE Red 2029 vs Sweetlax Upstate',
                     '3d NE Red 2029 vs NJ Riot', '3d NE Red 2029 vs 2Way'],
        'Great 8': ['3d 29 vs 91 Colorado Great 8', '3d 29 vs Shore2Shore Great 8',
                    '3d 29 vs Tri-State Great 8', '3d 29 vs Mad Dog West Great 8'],
        'NAL': ['sweetlax upstate', 'vs resolute ohio'],
        'Alliance': ['vs 2way', 'vs maddog east', 'vs Sweetlax'],
    }.items():
        (summer / folder).mkdir(parents=True)
        for n in names:
            (summer / folder / f'{n}.mp4').touch()

    # "3D Georgia" must not take "91 georgia": both words fit that file — "3d" off OUR
    # name, "georgia" off a different club — and it used to win by 0.006.
    want = {'3D Georgia': ('Sweetlax', '3d NE Red 2029 vs 3d Georgia.mp4'),
            '91 Colorado': ('Great 8', '3d 29 vs 91 Colorado Great 8.mp4'),
            'Shore2Shore': ('Great 8', '3d 29 vs Shore2Shore Great 8.mp4'),
            'Tri-State': ('Great 8', '3d 29 vs Tri-State Great 8.mp4'),
            'NJ riot': ('Sweetlax', '3d NE Red 2029 vs NJ Riot.mp4'),
            'Maddog East': ('Alliance', 'vs maddog east.mp4'),
            'Leading Edge': ('LI showdown', '29s 3D NE Red vs leading edge.mp4'),
            'Hilltop': ('LI showdown', '29s 3D NE vs Hilltop.mp4'),
            'Resolute Ohio': ('NAL', 'vs resolute ohio.mp4')}
    for game, (folder, fname) in want.items():
        r = Row(sheet_row=2, game=game, start=1.0, end=2.0)
        match_videos([r], summer)
        assert r.src and (r.src.parent.name, r.src.name) == (folder, fname), (game, r.src)
        assert not r.flags, (game, r.flags)

    # two real games against one opponent can't be told apart by name: pick one, but say
    # so — never silently. ("sweetlax upstate" at NAL vs the Sweetlax tournament tape.)
    for game in ('Sweetlax Upstate', '2way'):
        r = Row(sheet_row=3, game=game, start=1.0, end=2.0)
        match_videos([r], summer)
        assert r.src, game
        assert any('ambiguous' in f for f in r.flags), (game, r.flags)

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
# clip 1: in/out at the sequence rate (here also 29.97), 10s-15s
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

# --- team names on labels ----------------------------------------------------
# "Goal" + game "Sweetlax Upstate" -> "Goal vs Sweetlax Upstate", straight off the sheet.
assert label_text(Row(sheet_row=2, label='Goal', game='Sweetlax Upstate')) == 'Goal'
assert label_text(Row(sheet_row=2, label='Goal', game='Sweetlax Upstate'), True) == \
    'Goal vs Sweetlax Upstate'
assert clean_team('Georgetown Prep #2 (Daksh )') == 'Georgetown Prep #2'
assert clean_team('vs Legacy') == 'Legacy'          # sheet already says "vs" — don't double it
assert label_text(Row(sheet_row=2, label='Save', game='vs Legacy'), True) == 'Save vs Legacy'
assert label_text(Row(sheet_row=2, label='Goal vs Sweetlax', game='Sweetlax'), True) == \
    'Goal vs Sweetlax'                              # label already names the team
assert label_text(Row(sheet_row=2, label='Goal vs Sweetlax', game='Sweetlax Upstate'), True) == \
    'Goal vs Sweetlax'                              # ...even when it abbreviates the team
assert label_text(Row(sheet_row=2, label='', game='2way'), True) == 'vs 2way'
assert label_text(Row(sheet_row=2, label='Goal', game=''), True) == 'Goal'

# --- toggled-off clips leave gaps ---------------------------------------------
# A clip toggled off — or one with no video but a known length — must hold its slot:
# dropping it would shift every later clip earlier and wreck the planned edit.
r_on = Row(sheet_row=2, game='A', label='Goal', start=10.0, end=15.0, src=fake)
r_off = Row(sheet_row=3, game='B', start=100.0, end=112.0, src=fake, enabled=False)
r_novid = Row(sheet_row=4, game='C', start=5.0, end=9.0)      # no video matched
r_wf = Row(sheet_row=5, game='F', whole_file=True, src=fake, enabled=False)
r_hopeless = Row(sheet_row=6, game='D')                        # no video AND no times
r_last = Row(sheet_row=7, game='E', start=20.0, end=26.0, src=fake)
assert place_kind(r_on) == 'clip'
assert place_kind(r_off) == 'gap'
assert place_kind(r_novid) == 'gap'
assert place_kind(r_wf) == 'gap'                   # off whole-file: gap of the file's length
assert place_kind(r_hopeless) is None

grows = [r_on, r_off, r_novid, r_wf, r_hopeless, r_last]
segs = timeline_layout(grows, probes)
assert [s['kind'] for s in segs] == ['clip', 'gap', 'gap', 'gap', 'clip']
assert segs[1]['t0'] == 5.0 and segs[1]['dur'] == 12.0
assert segs[3]['t0'] == 21.0 and segs[3]['dur'] == 100.0   # whole-file gap = probed length
assert segs[4]['t0'] == 121.0                      # 5 + 12 + 4 + 100: every gap held its slot

xml_g = build_xmeml(grows, probes, 'Gaps',
                    {4: Path('lbl.png')})          # index into PLACEABLE rows -> r_last
gr = ET.fromstring(xml_g)
gitems = gr.findall('.//video/track')[0].findall('clipitem')   # track 1: the footage
assert len(gitems) == 2                            # only real clips carry footage
fps = 29.97
gap_frames = round(5 * fps) + round(12 * fps) + round(4 * fps) + round(100 * fps)
# the second clip starts after clip1 + all three gaps — nothing shifted earlier
assert gitems[1].find('start').text == str(gap_frames)
assert gitems[1].find('end').text == str(gap_frames + round(6 * fps))
# numbering counts gaps (toggling 02 off cannot renumber 05); note a row whose length
# is unknowable (r_hopeless) IS still dropped and does renumber what follows — the
# Generate dialog warns about exactly those rows
assert gitems[0].find('name').text.startswith('01 - ')
assert gitems[1].find('name').text.startswith('05 - ')
# each gap leaves a marker saying what belongs in the hole
gmarks = gr.findall('sequence/marker')
assert any('02 - B' in m.find('name').text and 'toggled off' in m.find('comment').text
           for m in gmarks), [ET.tostring(m) for m in gmarks]
assert any('03 - C' in m.find('name').text and 'no video' in m.find('comment').text
           for m in gmarks)
assert any('04 - F' in m.find('name').text and 'toggled off' in m.find('comment').text
           for m in gmarks)
# the gap marker sits at the gap's start on the sequence
mk = next(m for m in gmarks if '02 - B' in m.find('name').text)
assert mk.find('in').text == str(round(5 * fps))
# audio track skips gaps too, and the label PNG landed on the right clip (05)
assert len(gr.findall('.//audio/track/clipitem')) == 2
overlay = gr.findall('.//video/track')[1].findall('clipitem')
assert len(overlay) == 1 and overlay[0].find('name').text.startswith('05 - ')
# sequence runs the full length including gaps
assert gr.find('sequence/duration').text == str(gap_frames + round(6 * fps))
# all clips toggled off -> a clear error, not a broken timeline
try:
    build_xmeml([r_off], probes, 'x')
    assert False, 'expected ValueError'
except ValueError:
    pass
# a whole-file row that was never probed has no knowable length -> loud error,
# not a TypeError or a silently wrong timeline
try:
    build_xmeml([r_on, Row(sheet_row=8, game='G', whole_file=True,
                           src=Path('never_probed.mp4'), enabled=False)],
                probes, 'x')
    assert False, 'expected ValueError'
except ValueError as e:
    assert 'probed' in str(e)

# --- mixed frame rates on one timeline ---------------------------------------
# Premiere reads every frame number on a clipitem at the SEQUENCE rate, whatever <rate>
# the clipitem carries. Counting in/out in source frames sent 29.97 film to half its
# timecode and 119.88 to double, and pushed a late clip in a fast file past the end of
# its media, where it imported as nothing at all. Only footage matching the sequence
# came in right, which is what made it look like a handful of bad rows.
slow, fast, base = HERE / 'slow.mp4', HERE / 'fast.mp4', HERE / 'base.mp4'
mixed_probes = {
    str(base): {'fps': 60 / 1.001, 'width': 1920, 'height': 1080,
                'duration': 3600.0, 'has_audio': True},
    str(slow): {'fps': 30 / 1.001, 'width': 1920, 'height': 1080,
                'duration': 3600.0, 'has_audio': True},
    str(fast): {'fps': 120 / 1.001, 'width': 2784, 'height': 1566,
                'duration': 2237.2, 'has_audio': True},
}
mixed_rows = [  # two 59.94 rows make it the sequence rate, as the real library does
    Row(sheet_row=2, game='Futures', start=1826.0, end=1840.0, order=1, src=base),
    Row(sheet_row=3, game='Hilltop', start=705.0, end=715.0, order=2, src=base),
    Row(sheet_row=4, game='3D Georgia', start=869.0, end=878.0, order=3, src=slow),
    Row(sheet_row=5, game='Sweetlax Upstate', start=1753.0, end=1766.0, order=4, src=fast),
]
mroot = ET.fromstring(build_xmeml(mixed_rows, mixed_probes, 'Mixed'))
mseq = int(mroot.find('sequence/rate/timebase').text) / 1.001
assert round(mseq, 2) == round(60 / 1.001, 2), mseq
for item, row in zip(mroot.findall('.//video/track/clipitem'), mixed_rows):
    landed = int(item.find('in').text) / mseq
    assert abs(landed - row.start) < 0.05, (row.game, landed, row.start)
    # and it must still point inside the media, or Premiere imports no clip at all
    assert landed < mixed_probes[str(row.src)]['duration'], (row.game, landed)
    assert item.find('rate/timebase').text == mroot.find('sequence/rate/timebase').text
# the <file> keeps its own rate — that is how Premiere conforms the media
fast_def = [f for f in mroot.findall('.//file')
            if f.find('name') is not None and f.find('name').text == 'fast.mp4'][0]
assert fast_def.find('rate/timebase').text == '120', ET.tostring(fast_def.find('rate'))

# --- generate(): a wrong-length video becomes a gap, not a silent shift ------
# Matching picked a file that's too short for the row's times. The row's length is
# still known, so it must hold its slot — the preview and the XML must agree.
import chopper as _ch
_orig_probe, _orig_ff = _ch.probe, _ch.find_ffmpeg
_ch.find_ffmpeg = lambda: ('ffmpeg', 'ffprobe')
_ch.probe = lambda fp, path: {'fps': 30.0, 'width': 1920, 'height': 1080,
                              'duration': 2400.0, 'has_audio': True, 'vfr': False}
try:
    with tempfile.TemporaryDirectory() as td:
        vid = Path(td) / 'a.mp4'
        g_rows = [Row(sheet_row=2, game='A', start=0.0, end=5.0, src=vid),
                  Row(sheet_row=3, game='B', start=3000.0, end=3020.0, src=vid),  # past 2400s
                  Row(sheet_row=4, game='C', start=20.0, end=26.0, src=vid)]
        msgs = []
        xmlp, ok, failed = _ch.generate(g_rows, Path(td), 'GapGen', False, msgs.append)
        rg = ET.fromstring(xmlp.read_text())
        items = rg.findall('.//video/track')[0].findall('clipitem')
        assert [i.find('name').text[:2] for i in items] == ['01', '03'], \
            [i.find('name').text for i in items]
        # C still starts after A + B's 20s slot — B became a gap, nothing shifted
        assert items[1].find('start').text == str(round(5 * 30.0) + round(20 * 30.0))
        assert any('gap' in (m.find('comment').text or '')
                   for m in rg.findall('sequence/marker'))
        assert any('Left as a gap' in m for m in msgs), msgs
finally:
    _ch.probe, _ch.find_ffmpeg = _orig_probe, _orig_ff

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

# NEVER let tests hit the real form: blank the baked-in URL for this whole section
orig_track_url = chopper.TRACK_URL
chopper.TRACK_URL = ''
chopper.send_usage_ping('deadbeef')          # empty TRACK_URL -> must be a no-op
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
chopper.TRACK_URL = ''   # deliberately NOT restored: nothing after this may ping
srv.shutdown()
assert orig_track_url is not None  # silence unused warning; original stays blanked

# --- frame rate detection ----------------------------------------------------
# A wrong fps puts every timeline clip at the wrong TIME, drifting further the deeper
# into the game. Concatenated/VFR footage is where the nominal rate lies.
import json as _json
import subprocess as _sub
from chopper import _fraction, probe

assert _fraction('30000/1001') == 30000 / 1001
assert _fraction('30/1') == 30.0
assert _fraction('0/0') == 0.0            # ffprobe's "unknown"
assert _fraction('90000/1') == 0.0        # cover-art nonsense rate
assert _fraction(None) == 0.0


def _fake_probe(streams, duration='3600.0'):
    payload = _json.dumps({'streams': streams, 'format': {'duration': duration}})
    real = _sub.run
    _sub.run = lambda *a, **k: types.SimpleNamespace(stdout=payload)
    try:
        chopper_probe_cache = __import__('chopper')._probe_cache
        chopper_probe_cache.clear()
        return probe('ffprobe', f'fake{len(payload)}{duration}.mp4')
    finally:
        _sub.run = real


import types
# honest constant-rate NTSC: used as-is, not flagged
i = _fake_probe([{'codec_type': 'video', 'width': 1920, 'height': 1080,
                  'r_frame_rate': '30000/1001', 'avg_frame_rate': '30000/1001',
                  'nb_frames': '107892'}])
assert not i['vfr'] and abs(i['fps'] - 30000 / 1001) < 1e-6, i

# nominal rate lies (concatenated game film): trust the average, flag it
j = _fake_probe([{'codec_type': 'video', 'width': 1920, 'height': 1080,
                  'r_frame_rate': '30/1', 'avg_frame_rate': '45/2', 'nb_frames': '81000'}])
assert j['vfr'] and abs(j['fps'] - 22.5) < 1e-6, j

# avg_frame_rate missing/zero -> fall back to frames/duration, then to nominal
k = _fake_probe([{'codec_type': 'video', 'width': 1920, 'height': 1080,
                  'r_frame_rate': '30/1', 'avg_frame_rate': '0/0', 'nb_frames': '90000'}])
assert abs(k['fps'] - 25.0) < 1e-6, k        # 90000 frames / 3600 s
m = _fake_probe([{'codec_type': 'video', 'width': 1920, 'height': 1080,
                  'r_frame_rate': '25/1', 'avg_frame_rate': '0/0'}])
assert abs(m['fps'] - 25.0) < 1e-6 and not m['vfr'], m

# cover art must not decide the sequence rate
n = _fake_probe([{'codec_type': 'video', 'width': 100, 'height': 100,
                  'r_frame_rate': '90000/1', 'avg_frame_rate': '90000/1',
                  'disposition': {'attached_pic': 1}},
                 {'codec_type': 'video', 'width': 1920, 'height': 1080,
                  'r_frame_rate': '30000/1001', 'avg_frame_rate': '30000/1001'},
                 {'codec_type': 'audio'}])
assert n['width'] == 1920 and abs(n['fps'] - 30000 / 1001) < 1e-6 and n['has_audio'], n

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
