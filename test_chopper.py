#!/usr/bin/env python3
"""Self-check for chopper.py — run: py test_chopper.py"""

import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')
from chopper import (Row, age_groups_of, build_xmeml, clean_team, filter_age_group,
                     find_columns, label_text, list_age_groups, list_videos,
                     match_score, match_videos, norm_tokens, opponent_tokens,
                     parse_range, parse_sheet, parse_tc, place_kind,
                     qualifiers_fit, render_label_png, sanitize_filename,
                     timeline_layout, video_age_groups)

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
# real sheets arrive with shift-slip typos: ';' for ':' and a dropped leading zero
assert parse_range('11:07 - 11;14') == (667, 674)
assert parse_range(':52 - :57') == (52, 57)
assert parse_tc(';52') == 52

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

# --- labels living in a "Notes" column, ordinal order values -----------------
# Will's sheet: no Label column at all — the play word sits in a second column
# HEADED "Notes", and Order says "10th"/"11th". Captions came out "VS TEAM"
# with no play until the notes column was promoted to labels.
will = [
    ['', 'Game/Showcase', 'Game', '# of Seconds', 'Notes', 'Order', '', '',
     'Order for Tape', 'Clip', 'Notes', ''],
    ['', '', 'vs 2Way', '', '', '10th', '', '', '', '24:06-24:16', 'Goal',
     'https://youtube.com/x'],
    ['', '3d Upstate', 'vs Primetime', '', '', '1st', '', '', '', '09:35-09:45', 'Goal',
     'https://youtube.com/y'],
    ['', '', 'vs Riot', '', '', '2nd', '', '', '', '05:12-05:22', 'CTO',
     'https://youtube.com/z'],
]
with tempfile.TemporaryDirectory() as td:
    import csv as _csv
    wp = Path(td) / 'will.csv'
    with open(wp, 'w', newline='') as f:
        _csv.writer(f).writerows(will)
    wrows = parse_sheet(wp)
    assert [r.game for r in wrows] == ['vs Primetime', 'vs Riot', 'vs 2Way']  # 1st,2nd,10th
    assert [r.label for r in wrows] == ['Goal', 'CTO', 'Goal']    # promoted from Notes
    assert [r.order for r in wrows] == [1, 2, 10]                 # ordinals parsed
    assert all('youtube' not in r.label for r in wrows)   # URLs never become labels
    assert all('youtube' not in r.notes for r in wrows)   # (their column has no header)

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

# --- compound names and team qualifiers ---------------------------------------
# The sheet says "Maddog", the file says "Mad Dog" — either spelling must cover
# the other. And East/West are IDENTITY, like digits: "Maddog West" must never
# take the East tape just because the club name agrees. (Real failure: the club
# fields a Maddog East AND a Maddog West.)
assert match_score('Maddog West', '3d 29 vs Mad Dog West Great 8') > 0.7
assert match_score('Mad Dog East', 'vs maddogeast') > 0.55        # fused the other way
# initialisms both directions, sheet or file: DCE=DC Express, BTB=Be The Best,
# PT=Prime Time, NL=Next level — and 'preds' is 'Predators', 'ten' is '10'
assert match_score('vs DCE', 'vs DC express') > 0.7
assert match_score('vs Be The Best', 'vs BTB') > 0.7
assert match_score('vs Prime Time', 'vs PT') > 0.7
assert match_score('Next level', 'vs NL') > 0.55
assert match_score('vs Predators', 'vs preds') > 0.7
assert match_score('Team 10', 'vs team ten') > 0.7
assert norm_tokens('vs team ten') == ['vs', 'team', '10']
# 'vs' is a separator, never a name: it must not hand 'vs mesa' a score for
# the game 'vs S2S' (it once read as the initialism of 'Vs S2s' and did)
assert match_score('vs S2S', '2028- vs S2S') > match_score('vs S2S', 'vs mesa') + 0.3
# a terse file fully contained in the game still loses to one naming more of it
assert match_score('vs 3D New England', 'vs 3D') > 0.4
assert match_score('vs 3D Garden State', 'vs GS') > match_score('vs 3D Garden State', 'vs 3D')
assert qualifiers_fit(['maddog', 'east'], ['maddog', 'east'], []) is True
assert qualifiers_fit(['maddog', 'east'], ['maddog', 'west'], []) is False
assert qualifiers_fit(['maddog', 'east'], ['maddog'], []) is True   # unmarked survives
# our own team's color must never disqualify a file — only the opponent half counts
assert qualifiers_fit(['sweetlax', 'white'],
                      opponent_tokens(norm_tokens('3d NE Red 2029 vs Sweetlax White')),
                      []) is True
# a folder that names the game carries its qualifier; one that doesn't claims nothing
assert qualifiers_fit(['maddog', 'east'], ['game1'], [{'maddog', 'west'}]) is False
assert qualifiers_fit(['maddog', 'east'], ['game1'], [{'gold', 'bracket'}]) is True

with tempfile.TemporaryDirectory() as td:
    (Path(td) / 'vs maddog east.mp4').touch()
    (Path(td) / '3d 29 vs Mad Dog West Great 8.mp4').touch()
    r_e = Row(sheet_row=2, game='Maddog East', start=1.0, end=2.0)
    r_w = Row(sheet_row=3, game='Maddog West', start=1.0, end=2.0)
    match_videos([r_e, r_w], td)
    assert r_e.src and r_e.src.name == 'vs maddog east.mp4', r_e.src
    assert not r_e.flags, r_e.flags
    assert r_w.src and 'West' in r_w.src.name, r_w.src
    assert not r_w.flags, r_w.flags
    # a game naming NO qualifier still sees both tapes (and gets flagged ambiguous)
    r_m = Row(sheet_row=4, game='Maddog', start=1.0, end=2.0)
    match_videos([r_m], td)
    assert r_m.src is not None

# --- age groups ---------------------------------------------------------------
# One club, several age groups: each tournament folder holds 2027/…/2031 with
# near-identical filenames. Selecting an age group must keep matching from
# picking another team's tape of the same opponent.
assert age_groups_of('2029') == {'2029'}
assert age_groups_of('29') == {'2029'} and age_groups_of('29s') == {'2029'}
assert age_groups_of("3d ne '29") == {'2029'}          # apostrophe year mid-name
assert age_groups_of('3d NE Red 2029 vs 2Way') == {'2029'}
assert age_groups_of('2027 tournament June 27') == {'2027'}
# dates, scores, and jersey numbers must NOT read as age groups
assert age_groups_of('June 27') == set()
assert age_groups_of('Day 2') == set()
assert age_groups_of('Great 8') == set()
assert age_groups_of('3d 29 vs 91 Colorado Great 8') == set()   # bare 2-digit mid-name
assert age_groups_of('91 georgia') == set()

with tempfile.TemporaryDirectory() as td:
    big = Path(td) / 'Club'
    for tourn in ('Sweetlax', 'NAL'):
        for year in ('2027', '2029', '2031'):
            (big / tourn / year).mkdir(parents=True)
            (big / tourn / year / 'vs sweetlax.mp4').touch()
    (big / 'Sweetlax' / 'loose scrimmage.mp4').touch()   # no age marker anywhere
    vids = list_videos(big)
    assert list_age_groups(vids) == ['2027', '2029', '2031']
    assert video_age_groups(big / 'Sweetlax' / '2029' / 'vs sweetlax.mp4') == {'2029'}
    assert video_age_groups(big / 'Sweetlax' / 'loose scrimmage.mp4') == set()
    kept = filter_age_group(vids, '2029')
    # the other years' tapes are gone; unmarked files survive (they might be ours)
    assert len(kept) == 3, kept                          # 2 x 2029 + the loose file
    assert all('2027' not in str(v) and '2031' not in str(v) for v in kept)
    assert filter_age_group(vids, '') == vids            # no selection = no filter
    # end to end: same opponent in every age folder — the filter disambiguates
    r_sw = Row(sheet_row=2, game='Sweetlax vs sweetlax', start=1.0, end=2.0)
    match_videos([r_sw], big, videos=filter_age_group(vids, '2029'))
    assert r_sw.src and '2029' in str(r_sw.src), r_sw.src

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

# Premiere resolves linked A/V by the clip's POSITION ON ITS TRACK. Gaps hold a
# slot but put nothing on the track, and a mute clip advances only the video
# track — numbering links by slot sent every link past its clip, and Premiere
# silently dropped the whole import (a loading bar, then nothing).
mute = HERE / 'mute_src.mp4'
probes_l = dict(probes)
probes_l[str(mute)] = {'fps': 29.97, 'width': 1920, 'height': 1080,
                       'duration': 100.0, 'has_audio': False}
lrows = [r_on, r_off,                                          # clip, gap
         Row(sheet_row=8, game='M', start=1.0, end=4.0, src=mute),   # video-only clip
         r_last]                                               # clip with audio
lr = ET.fromstring(build_xmeml(lrows, probes_l, 'Links'))
vtrack = lr.findall('.//video/track')[0].findall('clipitem')
atrack = lr.findall('.//audio/track')[0].findall('clipitem')
assert len(vtrack) == 3 and len(atrack) == 2                   # mute clip: no audio item
vids_l = [c.get('id') for c in vtrack]
aids_l = [c.get('id') for c in atrack]
links_seen = 0
for c in vtrack + atrack:
    for ln in c.findall('link'):
        ref, mt = ln.findtext('linkclipref'), ln.findtext('mediatype')
        pool = vids_l if mt == 'video' else aids_l
        assert ref in pool, (c.get('id'), ref)
        assert pool.index(ref) + 1 == int(ln.findtext('clipindex')), \
            (c.get('id'), ref, mt, ln.findtext('clipindex'))
        links_seen += 1
assert links_seen == 8                                         # 2 linked pairs x 2 links x 2 items

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

# --- watchable preview video (needs ffmpeg; skipped if unavailable) -----------
# The preview must BE the timeline: clips in order, black for exactly each gap's
# length, one uniform stream no matter how the sources differ.
_ff = None
try:
    _ff = _ch.find_ffmpeg()
except Exception:
    print('NOTE: ffmpeg unavailable — preview-video render test skipped')
if _ff:
    import json as _pjson
    import subprocess as _psub
    ffmpeg_bin, ffprobe_bin = _ff
    with tempfile.TemporaryDirectory() as td:
        srcs = {}
        for name, audio in (('loud.mp4', True), ('mute.mp4', False)):
            p = Path(td) / name
            cmd = [ffmpeg_bin, '-y', '-hide_banner', '-loglevel', 'error',
                   '-f', 'lavfi', '-i', 'testsrc=size=320x180:rate=25:duration=3']
            if audio:
                cmd += ['-f', 'lavfi', '-t', '3', '-i', 'sine=frequency=440',
                        '-c:a', 'aac', '-shortest']
            cmd += ['-c:v', 'libx264', '-pix_fmt', 'yuv420p', str(p)]
            _psub.run(cmd, capture_output=True, check=True)
            srcs[name] = p
        w_rows = [Row(sheet_row=2, game='A', start=0.5, end=1.5, src=srcs['loud.mp4']),
                  Row(sheet_row=3, game='B', start=0.0, end=2.0, src=srcs['loud.mp4'],
                      enabled=False),                       # -> 2s of black
                  Row(sheet_row=4, game='C', start=1.0, end=2.0, src=srcs['mute.mp4'])]
        w_probes = {str(p): _ch.probe(ffprobe_bin, p) for p in srcs.values()}
        w_segs = timeline_layout(w_rows, w_probes)
        out = _ch.render_preview_video(w_segs, w_probes, ffmpeg_bin)
        fmt = _pjson.loads(_psub.run(
            [ffprobe_bin, '-v', 'error', '-show_entries', 'format=duration',
             '-of', 'json', str(out)], capture_output=True, text=True).stdout)
        got = float(fmt['format']['duration'])
        assert abs(got - 4.0) < 0.35, got     # 1 + 2 (gap) + 1, small container slack
        # mixed audio/no-audio sources still concat into one playable stream
        streams = _pjson.loads(_psub.run(
            [ffprobe_bin, '-v', 'error', '-show_entries', 'stream=codec_type',
             '-of', 'json', str(out)], capture_output=True, text=True).stdout)
        kinds = sorted(s['codec_type'] for s in streams['streams'])
        assert kinds == ['audio', 'video'], kinds

        # whistle detection: a 3kHz blast over crowd noise reads as a whistle,
        # plain crowd noise does not, and a silent track gives NO verdict
        from chopper import whistle_score, WHISTLE_MIN
        wf = Path(td) / 'whistle.mp4'
        _psub.run([ffmpeg_bin, '-y', '-hide_banner', '-loglevel', 'error',
                   '-f', 'lavfi', '-i', 'anoisesrc=color=brown:amplitude=0.25:duration=12',
                   '-f', 'lavfi', '-i',
                   'sine=frequency=3000:duration=1.2,adelay=8000|8000,volume=3',
                   '-filter_complex', '[0][1]amix=inputs=2:duration=first',
                   '-c:a', 'aac', str(wf)], capture_output=True, check=True)
        s_at = whistle_score(ffmpeg_bin, wf, 8.0, back=0.5, ahead=2.5)
        s_off = whistle_score(ffmpeg_bin, wf, 3.0, back=0.5, ahead=2.5)
        assert s_at is not None and s_at >= WHISTLE_MIN, s_at
        assert s_off is not None and s_off < s_at, (s_off, s_at)
        assert whistle_score(ffmpeg_bin, srcs['mute.mp4'], 1.5) is None  # no audio

# --- caption mogrt stamping ---------------------------------------------------
# Premiere 2026 removed the mogrt-text scripting API, so captions are baked into
# per-clip mogrt copies. The text lives in a FlatBuffers blob behind a 12-byte
# Adobe wrapper (u32 payload size, u32 0, magic 44 33 22 11), length-prefixed at
# the buffer tail. The stamp must swap the string, re-pad to 4 bytes, and fix
# the wrapper size — and must refuse cleanly if the format ever changes.
import base64 as _b64
import gzip as _gz
import io as _io
import json as _mjson
import re as _mre
import struct as _st
import zipfile as _zf
from make_caption_mogrts import load_template, stamp

with tempfile.TemporaryDirectory() as td:
    ph = 'GOAL VS TEAM...'
    blob = bytearray(_st.pack('<II', 0, 0) + bytes.fromhex('44332211'))
    blob += b'\x0c\x00\x00\x00TABLEJUNK\x00\x00\x00'          # fake flatbuffer guts
    blob += _st.pack('<I', len(ph)) + ph.encode() + b'\x00'
    while len(blob) % 4:
        blob += b'\x00'
    _st.pack_into('<I', blob, 0, len(blob) - 12)
    xml_t = ('<x><StartKeyframeValue Encoding="base64" BinaryHash="k">'
             + _b64.b64encode(bytes(blob)).decode() + '</StartKeyframeValue></x>')
    inner_t = _io.BytesIO()
    with _zf.ZipFile(inner_t, 'w') as z:
        z.writestr('g.prproj', _gz.compress(xml_t.encode()))
    definition = {'capsuleID': 'x', 'capsuleName': 'T',
                  'capsuleNameLocalized': {'strDB': [{'localeString': 'en_US', 'str': 'T'}]},
                  'clientControls': [{'id': '4', 'value': {'strDB': [
                      {'localeString': 'en_US', 'str': ph}]}}]}
    tpl_path = Path(td) / 't.mogrt'
    with _zf.ZipFile(tpl_path, 'w') as z:
        z.writestr('definition.json', _mjson.dumps(definition))
        z.writestr('project.prgraphic', inner_t.getvalue())
    tpl = load_template(tpl_path)
    outp = Path(td) / '01.mogrt'
    stamp(*tpl, outp, 1, 'GOAL VS DCE AND MORE WORDS')        # longer than placeholder
    with _zf.ZipFile(outp) as z:
        d2 = _mjson.loads(z.read('definition.json'))
        assert d2['clientControls'][0]['value']['strDB'][0]['str'] == 'GOAL VS DCE AND MORE WORDS'
        assert d2['capsuleName'] == '01 GOAL VS DCE AND MORE WORDS'
        assert d2['capsuleID'] != 'x'                          # fresh id per stamped copy
        inz = _zf.ZipFile(_io.BytesIO(z.read('project.prgraphic')))
        x2 = _gz.decompress(inz.read('g.prproj')).decode()
        b2 = _b64.b64decode(_mre.search(r'>([A-Za-z0-9+/=]+)<', x2).group(1))
        t2 = b2.find(b'GOAL VS DCE AND MORE WORDS')
        assert t2 > 0 and ph.encode() not in b2
        assert _st.unpack_from('<I', b2, t2 - 4)[0] == len('GOAL VS DCE AND MORE WORDS')
        assert _st.unpack_from('<I', b2, 0)[0] == len(b2) - 12  # wrapper size fixed
        assert len(b2) % 4 == 0                                 # alignment preserved
        assert b2[4:t2 - 4] == bytes(blob[4:t2 - 4])            # nothing else disturbed
    # a template whose format we don't recognize must refuse, not corrupt
    bad = bytearray(blob)
    bad[8:12] = b'XXXX'                                        # wrong magic
    xml_b = ('<x><StartKeyframeValue Encoding="base64" BinaryHash="k">'
             + _b64.b64encode(bytes(bad)).decode() + '</StartKeyframeValue></x>')
    inner_b = _io.BytesIO()
    with _zf.ZipFile(inner_b, 'w') as z:
        z.writestr('g.prproj', _gz.compress(xml_b.encode()))
    bad_path = Path(td) / 'bad.mogrt'
    with _zf.ZipFile(bad_path, 'w') as z:
        z.writestr('definition.json', _mjson.dumps(definition))
        z.writestr('project.prgraphic', inner_b.getvalue())
    try:
        load_template(bad_path)
        assert False, 'expected refusal on unknown wrapper'
    except SystemExit as e:
        assert 'wrapper' in str(e)

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

    # color and position are honored: red text lands in the top-right quadrant
    png_tr = Path(td) / 'tr.png'
    render_label_png('GOAL', None, 48, 640, 360, png_tr, color='#ff0000', pos='top right')
    im_tr = Image.open(png_tr)
    bl, bt, br, bb = im_tr.getchannel('A').getbbox()
    assert bl >= 320 and bb <= 180, (bl, bt, br, bb)
    reds = [im_tr.getpixel((x, y)) for y in range(bt, bb) for x in range(bl, br)
            if im_tr.getpixel((x, y))[3] > 200]
    assert any(p[0] > 200 and p[1] < 80 for p in reds)     # the fill really is red
    # bottom center is centered horizontally and low
    png_bc = Path(td) / 'bc.png'
    render_label_png('GOAL', None, 48, 640, 360, png_bc, pos='bottom center')
    bl, bt, br, bb = Image.open(png_bc).getchannel('A').getbbox()
    assert bt >= 180 and abs((bl + br) / 2 - 320) < 30, (bl, bt, br, bb)
    # a junk color falls back to white instead of crashing
    render_label_png('GOAL', None, 48, 640, 360, png_bc, color='not-a-color')

    # free percentage positions — how the Premiere template places text: left edge
    # at X%, bottom at Y%
    png_pct = Path(td) / 'pct.png'
    render_label_png('GOAL', None, 48, 640, 360, png_pct, pos='50%,50%')
    bl, bt, br, bb = Image.open(png_pct).getchannel('A').getbbox()
    assert abs(bl - 320) <= 6 and abs(bb - 180) <= 6, (bl, bt, br, bb)
    # a nudge can't push the text off-frame
    render_label_png('GOAL', None, 48, 640, 360, png_pct, pos='99%,99%')
    bl, bt, br, bb = Image.open(png_pct).getchannel('A').getbbox()
    assert br <= 640 and bb <= 360, (bl, bt, br, bb)

# the baked-in template style ("Darsh_template.prproj": Futura-Medium 122px at
# 2160p, x1.41 Motion scale = 86 at 1080p, bottom-left corner) must stay valid
from chopper import DARSH_STYLE, POS_PCT_RE, style_font
assert POS_PCT_RE.fullmatch(DARSH_STYLE['pos'])
assert DARSH_STYLE['size'] == 86 and DARSH_STYLE['color'] == '#ffffff'
f = style_font()
assert f == '' or Path(f).exists()      # machine-dependent, but never a broken path

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
