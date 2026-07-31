#!/usr/bin/env python3
"""Clip Chopper — turn a spreadsheet cut list into a Premiere Pro timeline + clip files.

Drop a spreadsheet (.xlsx/.csv or a shared Google Sheets URL) and point at your
game-video folder. Review the parsed cut list, then Generate:
  * <sheet>_timeline.xml  — import into Premiere (File > Import), clips in order
  * clips/NN - Label - Game.mp4 — optional physically cut clip files

Source videos are only ever read, never copied or re-encoded.
"""

import copy
import csv
import difflib
import hashlib
import json
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

VIDEO_EXTS = {'.mp4', '.mov', '.m4v', '.mts', '.m2ts', '.avi', '.mkv', '.mpg', '.mpeg', '.wmv'}

# ---------------------------------------------------------------- spreadsheet

HEADER_KEYS = {
    'game': ['game', 'opponent', 'match', 'video', 'film', 'source'],
    'clip': ['clip', 'timestamp', 'timecode', 'time', 'range', 'start-stop'],
    'order': ['order', 'seq', 'sequence', '#', 'no.', 'num'],
    'label': ['label', 'description', 'play', 'action', 'type', 'skill'],
    'notes': ['notes', 'comments', 'comment', 'instructions'],
}

DASHES = re.compile(r'[‐-―−]')  # unicode dashes -> '-'
TC_RE = re.compile(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$')
RANGE_RE = re.compile(r'^\s*([\d:]+)\s*-\s*([\d:]+)\s*$')
WHOLE_FILE_RE = re.compile(r'clip|folder|whole|full|entire|file', re.I)


@dataclass
class Row:
    sheet_row: int
    game: str = ''
    range_text: str = ''
    start: float | None = None   # seconds
    end: float | None = None
    whole_file: bool = False     # "Clip in Folder" -> use entire matched file
    order: int | None = None
    label: str = ''
    notes: str = ''
    flags: list = field(default_factory=list)
    src: Path | None = None      # matched video file
    manual: bool = False         # user picked the file by hand — never auto-rematch
    enabled: bool = True         # toggled off -> the clip's slot stays as a gap


def parse_tc(text):
    """'14:18' -> 858.0, '1:00:10' -> 3610.0, ':52' -> 52.0, else None.

    Real sheets arrive with shift-slip typos: ';' for ':' and a dropped leading
    zero (':52' meaning 0:52). Both read fine to a human, so they parse here.
    """
    text = text.strip().replace(';', ':')
    if m := re.fullmatch(r':(\d{1,2})', text):
        return int(m.group(1))
    m = TC_RE.match(text)
    if not m:
        return None
    a, b, c = m.groups()
    if c is None:
        return int(a) * 60 + int(b)
    return int(a) * 3600 + int(b) * 60 + int(c)


def parse_range(text):
    """Return (start, end) seconds or None if text is not a timecode range."""
    m = RANGE_RE.match(DASHES.sub('-', text.replace(';', ':')))
    if not m:
        return None
    start, end = parse_tc(m.group(1)), parse_tc(m.group(2))
    if start is None or end is None:
        return None
    return start, end


def _cell_str(v):
    if v is None:
        return ''
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def load_grid(path):
    """Read .xlsx or .csv into a list of rows of strings."""
    path = Path(path)
    if path.suffix.lower() == '.csv':
        with open(path, newline='', encoding='utf-8-sig') as f:
            return [[c.strip() for c in row] for row in csv.reader(f)]
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    grid = [[_cell_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return grid


def fetch_google_sheet(url):
    """Download a link-shared Google Sheet (honoring the #gid= tab); return temp file path."""
    m = re.search(r'/spreadsheets/d/([\w-]+)', url)
    if not m:
        raise ValueError('Not a Google Sheets URL')
    gid = re.search(r'[#?&]gid=(\d+)', url)
    if gid:  # xlsx export always returns tab 1; csv export returns exactly the linked tab
        export = (f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export'
                  f'?format=csv&gid={gid.group(1)}')
        suffix = '.csv'
    else:
        export = f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx'
        suffix = '.xlsx'
    req = urllib.request.Request(export, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = r.read()
        title = r.headers.get_filename() or f'google sheet{suffix}'
    if (suffix == '.xlsx' and data[:2] != b'PK') or data.lstrip()[:1] == b'<':
        raise ValueError('Could not download sheet — is it shared "Anyone with the link"?')
    tmp = Path(tempfile.gettempdir()) / sanitize_filename(Path(title).stem + suffix)
    tmp.write_bytes(data)
    return tmp


def _match_quality(cell, keywords):
    """3 = exact keyword, 2 = starts with, 1 = contains, 0 = no match."""
    c = cell.lower().strip()
    if not c:
        return 0
    best = 0
    for kw in keywords:
        if c == kw:
            return 3
        if c.startswith(kw):
            best = max(best, 2)
        elif kw in c:
            best = max(best, 1)
    return best


def find_columns(grid):
    """Locate the header row and map field -> column index.

    Returns (header_row_index, {field: col or [cols] for notes}).
    Falls back to value-pattern detection if no header row is found.
    """
    best_row, best_score, best_map = None, 0, {}
    for i, row in enumerate(grid[:15]):
        colmap, score = {}, 0
        for f, kws in HEADER_KEYS.items():
            cands = [(q, j) for j, cell in enumerate(row)
                     if (q := _match_quality(cell, kws))]
            if not cands:
                continue
            top = max(q for q, _ in cands)
            cols = [j for q, j in cands if q == top]
            colmap[f] = cols if f == 'notes' else cols[0]
            score += top
        if len(colmap) >= 2 and score > best_score:
            best_row, best_score, best_map = i, score, colmap

    ncols = max((len(r) for r in grid), default=0)

    def range_hits(col, first_row):
        return sum(1 for r in grid[first_row:] if len(r) > col and parse_range(r[col]))

    if best_row is not None:
        # Trust the header's clip column only if it actually contains timecode ranges —
        # otherwise rescue by value pattern (e.g. a "Clip #" header stealing from "Timecodes",
        # or a range column under an unrecognized header like "Cut").
        cur = best_map.get('clip')
        if cur is None or range_hits(cur, best_row + 1) == 0:
            counts = [range_hits(j, best_row + 1) for j in range(ncols)]
            if counts and max(counts) > 0:
                best_map['clip'] = counts.index(max(counts))
        return best_row, best_map

    # No header row: detect the clip column by timecode-range values.
    counts = [range_hits(j, 0) for j in range(ncols)]
    if not counts or max(counts) == 0:
        raise ValueError('Could not find a timecode column (like "12:34-12:45") in this sheet')
    clip_col = counts.index(max(counts))
    # game column: text column left of clip with the most repeated non-timecode values
    game_col = None
    for j in range(clip_col - 1, -1, -1):
        vals = [r[j] for r in grid if len(r) > j and r[j] and not parse_range(r[j])]
        if len(vals) >= max(counts) // 2:
            game_col = j
            break
    colmap = {'clip': clip_col}
    if game_col is not None:
        colmap['game'] = game_col
    return -1, colmap


def parse_sheet(path):
    """Parse spreadsheet into ordered list of Row. Raises ValueError if hopeless."""
    grid = load_grid(path)
    header_row, cols = find_columns(grid)
    if 'clip' not in cols:
        raise ValueError('No clip/timecode column found')

    def get(row, col):
        return row[col].strip() if col is not None and len(row) > col else ''

    rows = []
    for i, raw in enumerate(grid):
        if i <= header_row:
            continue
        game = get(raw, cols.get('game'))
        clip_text = get(raw, cols['clip'])
        if not game and not clip_text:
            continue
        r = Row(sheet_row=i + 1, game=game, range_text=clip_text,
                label=get(raw, cols.get('label')))
        r.notes = ' | '.join(v for c in cols.get('notes', []) if (v := get(raw, c)))
        if (o := get(raw, cols.get('order'))):
            try:
                r.order = int(float(o))
            except ValueError:
                pass
        if not clip_text:
            r.flags.append('no timecode')
        else:
            rng = parse_range(clip_text)
            if rng:
                r.start, r.end = float(rng[0]), float(rng[1])
                if r.end <= r.start:
                    r.flags.append('end is before start — check times')
                elif r.end - r.start > 120:
                    r.flags.append(f'clip is {r.end - r.start:.0f}s long — check times')
                if r.start > 4 * 3600:
                    r.flags.append('start is past 4 hours — check times')
            elif WHOLE_FILE_RE.search(clip_text):
                r.whole_file = True  # e.g. "Clip in Folder"
            else:
                # A typo here must never silently become a full-game copy.
                r.flags.append(f'could not read timecode "{clip_text}" — double-click In/Out to set times')
        rows.append(r)
    if not rows:
        raise ValueError('No clip rows found in this sheet')
    big = 10 ** 6
    rows.sort(key=lambda r: r.order if r.order is not None else big + r.sheet_row)
    return rows


# ------------------------------------------------------------ video matching

NUM_WORDS = {'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
             'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10',
             'eleven': '11', 'twelve': '12'}


def norm_tokens(name):
    """'Georgetown Prep #2 (Daksh)' -> ['georgetown', 'prep', '2'].

    Number words become digits so "Team 10" finds "vs team ten.mp4" — and its
    "10" counts as carried identity for the digit rule.
    """
    name = re.sub(r'\([^)]*\)', ' ', name.lower())
    return [NUM_WORDS.get(t, t) for t in re.sub(r'[^a-z0-9]+', ' ', name).split()]


def as_folders(folders):
    """Accept one folder or several; always return a list of Paths."""
    if isinstance(folders, (str, Path)):
        return [Path(folders)]
    return [Path(f) for f in folders]


def list_videos(folders, exclude=None):
    """Every video under the chosen folder(s), recursively. Deduped if they overlap."""
    found = set()
    for folder in as_folders(folders):
        for p in folder.rglob('*'):
            if (p.suffix.lower() in VIDEO_EXTS and p.is_file()
                    # skip hidden files/dirs — esp. macOS "._*" AppleDouble junk on USB drives
                    and not any(part.startswith('.') for part in p.relative_to(folder).parts)
                    and not (exclude and exclude in p.parents)):
                found.add(p)
    return sorted(found)


def rel_to_root(path, roots):
    """This file's path relative to the SHALLOWEST chosen folder that contains it.

    Shallowest, not first-added: adding "Big/Rumble" and later "Big" must not strip the
    "Rumble" folder from that subtree, or a game named after it loses its identity.
    """
    for root in sorted(roots, key=lambda r: len(r.parts)):
        try:
            return path.relative_to(root)
        except ValueError:
            continue
    return Path(path.name)


# One club, several age groups: the same tournament folder holds 2027/, 2029/, 2031/
# with near-identical filenames inside. Selecting an age group keeps matching from
# picking another team's tape of the same opponent.
AGE_PATTERNS = (re.compile(r'\b20([23]\d)\b'),     # "2029", "2029 red"
                re.compile(r"'([23]\d)\b"),        # "3d ne '29"
                re.compile(r'^([23]\d)s?$'))       # a folder named just "29" / "29s"


def age_groups_of(text):
    """Age-group years a folder or file NAME declares: "3d ne '29" -> {'2029'}.

    Deliberately strict: a bare 2-digit number inside a longer name ("June 27",
    "Great 8", "3d 29 vs X") is NOT a marker — dates and scores would poison it.
    Only a 4-digit year, an apostrophe year, or a folder named just the year counts.
    """
    found = set()
    for pat in AGE_PATTERNS:
        found |= {f'20{m}' for m in pat.findall(text.strip())}
    return found


def video_age_groups(path):
    """Every age group marked anywhere on this file's FULL path (folders + stem).

    The full path, not the part under the chosen root: the year often lives in the
    root folder's own name ("3d ne '29/..."), which rel_to_root strips away.
    """
    groups = set()
    for part in Path(path).parts[:-1]:
        groups |= age_groups_of(part)
    groups |= age_groups_of(Path(path).stem)
    return groups


def list_age_groups(videos):
    """Sorted age groups marked across these videos — the dropdown's choices."""
    groups = set()
    for v in videos:
        groups |= video_age_groups(v)
    return sorted(groups)


def filter_age_group(videos, group):
    """Videos for one age group: keeps files marked with it AND unmarked files.

    A file marked only as a different year is definitively another team's tape;
    an unmarked file might be this team's game filed loosely, so it must survive —
    dropping it would hide the right answer, which folder evidence never gets to do.
    """
    if not group:
        return videos
    return [v for v in videos
            if not (marks := video_age_groups(v)) or group in marks]


def digits_fit(game_toks, digits, stem_toks, comp_toks):
    """Does this file carry the game's number as identity, not as coincidence?

    A number counts if it is in the filename, or in a folder that also names the game:
    "Georgetown Prep 2/2026/game.mp4" is Prep #2, but "Day 2/Georgetown Prep 1.mp4" is not.
    """
    if not digits:
        return True
    words = {t for t in game_toks if not t.isdigit()}
    carriers = set(stem_toks)
    for toks in comp_toks:
        if toks & words:
            carriers |= toks
    return digits <= carriers


# Words that split one club into several teams, grouped by axis. Like digits, these
# are IDENTITY: "Maddog East" must never take a tape marked West just because the
# club name agrees.
QUALIFIER_AXES = (
    {'east', 'west', 'north', 'south'},
    {'red', 'white', 'blue', 'black', 'green', 'gold', 'grey', 'gray', 'orange', 'silver'},
)


def qualifiers_fit(game_toks, opp_toks, comp_toks):
    """Does this file avoid claiming a CONFLICTING team qualifier?

    Checked against the opponent half of the filename (after "vs") plus folders
    that name the game, never our own team's half — "3d NE Red 2029 vs Sweetlax
    White" says Red about US, and must still fit a game named "Sweetlax White".
    A file naming the game's own qualifier anywhere survives; one naming only a
    rival qualifier is killed.
    """
    words = {t for t in game_toks if not t.isdigit()
             and not any(t in axis for axis in QUALIFIER_AXES)}
    carriers = set(opp_toks)
    for toks in comp_toks:
        if toks & words:      # "Gold bracket/" claims nothing about a game it doesn't name
            carriers |= toks
    for axis in QUALIFIER_AXES:
        named = axis & set(game_toks)
        if not named:
            continue
        claimed = axis & carriers
        if claimed and not (named & claimed):
            return False
    return True


def opponent_tokens(toks):
    """The part of a file name after "vs" — game film is named "<us> vs <them>".

    The sheet names the opponent, so our own team is noise that every file repeats.
    Worse, it is noise that scores: "3D Georgia" took all of its words from
    "29s 3D NE Red vs 91 georgia" — "3d" off OUR name, "georgia" off a different
    club — and beat its own tape by six thousandths.
    """
    for i in range(len(toks) - 1, -1, -1):
        if toks[i] in ('vs', 'v'):
            return toks[i + 1:] or toks
    return toks


def _abbrev(token, run):
    """Is `token` an initialism of this run of tokens? 'gs'~('garden','state'),
    'btb'~('be','the','best'), and 'dce'~('dc','express') — the first word may
    contribute itself, not just its first letter."""
    return token in (''.join(w[0] for w in run),
                     run[0] + ''.join(w[0] for w in run[1:]))


def _covers(t, u):
    """'preds' covers 'predators': equal, or a >=4-char prefix (plural 's' ignored)."""
    if t == u:
        return True
    a, b = t.rstrip('s'), u.rstrip('s')
    return ((a.startswith(b) and len(b) >= 4) or (b.startswith(a) and len(a) >= 4))


def _overlap_score(g, f):
    # Names differ every way people type them — "Maddog"/"Mad Dog"/"maddogeast",
    # "DCE"/"DC express", "BTB"/"Be The Best", "preds"/"Predators". Cover a game
    # token by: the same token, a fused run, an initialism (either direction), or
    # a shared >=4-char stem. A covered run covers ALL of its parts.
    # "vs" is a separator, never a name: left in, it hands every file a free
    # match — and reads as the INITIALISM of any "Vs S…" run, which once scored
    # "vs mesa" at 0.85 for the game "vs S2S".
    g = [t for t in g if t not in ('vs', 'v')] or g
    f = [t for t in f if t not in ('vs', 'v')] or f
    fset = set(f)
    fset |= {a + b for a, b in zip(f, f[1:])}
    fset |= {a + b + c for a, b, c in zip(f, f[1:], f[2:])}
    fruns = list(zip(f, f[1:])) + list(zip(f, f[1:], f[2:]))
    covered = {t for t in g
               if t in fset
               or any(_abbrev(t, run) for run in fruns)
               or any(_covers(t, u) for u in f)}
    gruns = list(zip(g, g[1:])) + list(zip(g, g[1:], g[2:]))
    for run in gruns:
        if ''.join(run) in fset or any(_abbrev(u, run) for u in f):
            covered |= set(run)
    gcover = len(covered) / len(g)
    # A terse file fully contained in the game name ("vs 3D.mp4" for "3D New
    # England") deserves credit too — weighted by gcover so a candidate naming
    # MORE of the game still beats the terse one.
    gset = set(g) | {''.join(r) for r in gruns}
    fcover = sum(1 for u in f
                 if u in gset
                 or any(_abbrev(u, run) for run in gruns)
                 or any(_covers(u, t) for t in g)) / len(f)
    overlap = max(gcover, fcover * (0.5 + 0.5 * gcover))
    ratio = difflib.SequenceMatcher(None, ' '.join(g), ' '.join(f)).ratio()
    return 0.6 * overlap + 0.4 * ratio


def match_score(game, filename):
    g, f = norm_tokens(game), norm_tokens(filename)
    if not g or not f:
        return 0.0
    # Score the whole name and the opponent alone, best wins: a file that names only
    # the opponent keeps its score, and one that buries the opponent behind our own
    # team is judged on the half that actually identifies the game.
    return max(_overlap_score(g, f), _overlap_score(g, opponent_tokens(f)))


def match_videos(rows, folders, threshold=0.55, exclude=None, videos=None):
    """Fill row.src by fuzzy-matching game names to files. One match per game name.

    Pass `videos` to reuse an existing scan — rglob over a big external drive is slow.
    """
    roots = as_folders(folders)
    if videos is None:
        videos = list_videos(roots, exclude=exclude)
    elif exclude:
        videos = [v for v in videos if exclude not in v.parents]
    # Score against the filename and against folders+filename, whichever fits better:
    # a name that matches on its own keeps its score, a name that only makes sense with
    # its tournament folder still gets found.
    rels = {v: rel_to_root(v, roots) for v in videos}
    stem_toks = {v: set(norm_tokens(rels[v].stem)) for v in videos}
    opp_toks = {v: set(opponent_tokens(norm_tokens(rels[v].stem))) for v in videos}
    comp_toks = {v: [set(norm_tokens(p)) for p in rels[v].parts[:-1]] for v in videos}
    labels = {v: ' '.join(rels[v].parts[:-1] + (rels[v].stem,)).strip() for v in videos}
    cache = {}
    for r in rows:
        if r.manual and r.src:
            continue
        if not r.game:
            r.flags.append('no game name')
            continue
        if r.game not in cache:
            # Digits are identity: "Prep #2" must never match "Prep 1.mp4" just because
            # the words agree.
            game_toks = norm_tokens(r.game)
            digits = {t for t in game_toks if t.isdigit()}
            pool = [v for v in videos
                    if digits_fit(game_toks, digits, stem_toks[v], comp_toks[v])
                    and qualifiers_fit(game_toks, opp_toks[v], comp_toks[v])]
            # Sitting in a folder the game names is evidence, never a filter: nudge those
            # candidates up rather than dropping the rest, or a tape filed outside its
            # tournament folder would vanish and a wrong one would win unflagged.
            def score(v):
                s = max(match_score(r.game, rels[v].stem), match_score(r.game, labels[v]))
                named = {t for t in game_toks
                         if any(t in toks for toks in comp_toks[v])}
                return min(1.0, s + 0.1) if named else s
            scored = sorted(((score(v), v) for v in pool),
                            reverse=True, key=lambda t: t[0])
            best = scored[0] if scored else (0.0, None)
            second = scored[1][0] if len(scored) > 1 else 0.0
            if best[0] < threshold:
                cache[r.game] = (None, 'no matching video — double-click File to pick')
            elif best[0] - second < 0.1 and second >= threshold:
                cache[r.game] = (best[1], f'ambiguous match ({best[1].name}?) — double-click File to confirm')
            else:
                cache[r.game] = (best[1], None)
        src, flag = cache[r.game]
        r.src = src
        if flag and flag not in r.flags:
            r.flags.append(flag)
    return rows


# ------------------------------------------------------- timeline placement

def place_kind(r):
    """'clip' = lands with footage, 'gap' = holds its length as empty track, None = no length.

    A row toggled off — or missing its video — must keep its slot: dropping it would
    shift every later clip earlier and wreck an edit planned around the sheet's order.
    Only a row whose length is unknowable (no timecode, no probe-able file) is dropped.
    """
    timed = r.start is not None and r.end is not None and r.end > r.start
    if r.enabled and r.src is not None and (r.whole_file or timed):
        return 'clip'
    if timed or (r.whole_file and r.src is not None):
        return 'gap'
    return None


def timeline_layout(rows, probes=None):
    """Where every row lands on the sequence, in seconds — the ONE place the timeline's
    shape is decided. build_xmeml and the preview window both draw from this, so what
    the preview shows is by construction what Premiere will import.

    Returns [{'row', 'kind', 'start', 'end', 'dur', 't0'}]: source in/out, clip length,
    and sequence position. A whole-file row not yet probed has end/dur of None — the
    preview shows a placeholder until probing lands; generate() never passes one.
    """
    probes = probes or {}
    segs, t = [], 0.0
    for r in rows:
        kind = place_kind(r)
        if not kind:
            continue
        if r.whole_file:
            p = probes.get(str(r.src)) if r.src else None
            start, end = 0.0, (p['duration'] if p else None)
        else:
            start, end = r.start, r.end
        dur = None if end is None else end - start
        segs.append({'row': r, 'kind': kind, 'start': start, 'end': end,
                     'dur': dur, 't0': t})
        t += dur or 0.0
    return segs


# ------------------------------------------------------------------- labels

def clean_team(game):
    """Sheet game name -> the opponent as you'd caption it.

    'Georgetown Prep #2 (Daksh )' -> 'Georgetown Prep #2', 'vs Legacy' -> 'Legacy'.
    """
    t = re.sub(r'\([^)]*\)', ' ', game)
    t = re.sub(r'^\s*(?:vs|v)\.?\s+', '', t, flags=re.I)
    return re.sub(r'\s+', ' ', t).strip()


def label_text(row, with_team=False):
    """Overlay text for a row: the label, plus 'vs <team>' straight off the sheet.

    A label that already names an opponent is left alone — 'Goal vs Sweetlax' must not
    become 'Goal vs Sweetlax vs Sweetlax Upstate', even when the label abbreviates the
    team. A row with no label still gets 'vs <team>' so every clip says who it was against.
    """
    if not with_team:
        return row.label
    team = clean_team(row.game) if row.game else ''
    if not team or (row.label and (team.lower() in row.label.lower()
                                   or re.search(r'\bvs\.?\s', row.label, re.I))):
        return row.label
    return f'{row.label} vs {team}' if row.label else f'vs {team}'


# ------------------------------------------------------------------- ffmpeg

def ensure_pip(module, pip_name):
    """Import module, pip-installing it first if missing (launched outside Chop.bat/.command).

    Raises ImportError if it still can't be imported after the install attempt.
    """
    try:
        return __import__(module)
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', pip_name],
                       capture_output=True)
        import importlib
        importlib.invalidate_caches()
        return __import__(module)


def find_ffmpeg():
    """Return (ffmpeg, ffprobe) paths. System PATH first, else static-ffmpeg."""
    ff, fp = shutil.which('ffmpeg'), shutil.which('ffprobe')
    if ff and fp:
        return ff, fp
    try:
        ensure_pip('static_ffmpeg', 'static-ffmpeg')
        from static_ffmpeg import run
    except ImportError:
        raise RuntimeError(
            'ffmpeg is missing. Launch the app with Chop.command (Mac) or Chop.bat '
            '(Windows) to auto-install everything, or run: '
            'python3 -m pip install static-ffmpeg') from None
    return run.get_or_fetch_platform_executables_else_raise()


_probe_cache = {}


def _fraction(text):
    """'30000/1001' -> 29.97002997, or 0.0 if unusable."""
    num, _, den = (text or '').partition('/')
    try:
        fps = float(num) / float(den or 1)
    except (ValueError, ZeroDivisionError):
        return 0.0
    return fps if 1.0 <= fps <= 240.0 else 0.0


def probe(ffprobe, path):
    """Return {'fps', 'width', 'height', 'duration', 'has_audio', 'vfr'} for a video.

    Timeline in/out points are frame numbers, so a wrong fps puts every clip at the
    wrong TIME in Premiere — and the error grows the deeper into the game you go
    (reporting 30 for 29.97 footage is +3s an hour in). Variable-frame-rate and
    concatenated files are where the nominal rate lies, so prefer the real average.
    """
    path = str(path)
    if path in _probe_cache:
        return _probe_cache[path]
    out = subprocess.run(
        [ffprobe, '-v', 'error', '-show_entries',
         'stream=codec_type,width,height,r_frame_rate,avg_frame_rate,duration,nb_frames'
         ':stream_disposition=attached_pic:format=duration',
         '-of', 'json', path],
        capture_output=True, text=True, check=True).stdout
    data = json.loads(out)
    info = {'fps': 30.0, 'width': 1920, 'height': 1080, 'duration': 0.0,
            'has_audio': False, 'vfr': False}
    if fd := data.get('format', {}).get('duration'):
        info['duration'] = float(fd)
    got_video = False
    for s in data.get('streams', []):
        if s.get('codec_type') == 'video' and not got_video:
            if s.get('disposition', {}).get('attached_pic'):
                continue  # embedded cover art, not the footage
            if 'width' in s:
                info['width'], info['height'] = s['width'], s['height']
            nominal = _fraction(s.get('r_frame_rate'))
            average = _fraction(s.get('avg_frame_rate'))
            if not info['duration'] and s.get('duration'):
                info['duration'] = float(s['duration'])
            # frames / seconds is ground truth when the container offers it
            try:
                measured = int(s['nb_frames']) / info['duration']
            except (KeyError, TypeError, ValueError, ZeroDivisionError):
                measured = 0.0
            measured = measured if 1.0 <= measured <= 240.0 else 0.0
            real = average or measured or nominal
            if nominal and real and abs(real - nominal) / nominal > 0.005:
                info['vfr'] = True          # nominal rate lies; trust the average
            info['fps'] = real or nominal or 30.0
            got_video = True
        elif s.get('codec_type') == 'audio':
            info['has_audio'] = True
    _probe_cache[path] = info
    return info


def cut_clip(ffmpeg, row, out_path):
    """Frame-accurate cut of one row into out_path. Source is only read."""
    if row.whole_file:
        shutil.copy2(row.src, out_path)
        return
    cmd = [ffmpeg, '-y', '-hide_banner', '-loglevel', 'error',
           '-ss', f'{row.start:.3f}', '-i', str(row.src),
           '-t', f'{row.end - row.start:.3f}',
           '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '18',
           '-c:a', 'aac', '-b:a', '192k', '-movflags', '+faststart', str(out_path)]
    subprocess.run(cmd, capture_output=True, text=True, check=True)


def grab_frame(ffmpeg, src, sec, out_path, width=None):
    """One frame at `sec` as a PNG — preview fodder. Source is only read."""
    vf = ['-vf', f'scale={width}:-2'] if width else []
    subprocess.run([ffmpeg, '-y', '-hide_banner', '-loglevel', 'error',
                    '-ss', f'{max(sec, 0):.3f}', '-i', str(src), '-frames:v', '1',
                    *vf, str(out_path)],
                   capture_output=True, text=True, check=True)


# ------------------------------------------------------------ preview video

PREVIEW_SIZE = (960, 540)   # low-res mezzanine: fast to encode, plenty to judge a cut
PREVIEW_FPS = 30
_watch_dir = None
_watch_parts = {}           # segment key -> encoded part path (cached per session)


def watch_dir():
    global _watch_dir
    if _watch_dir is None or not _watch_dir.exists():
        _watch_dir = Path(tempfile.mkdtemp(prefix='chop_watch_'))
    return _watch_dir


def encode_preview_part(ffmpeg, seg, probes):
    """One timeline segment as a uniform 960x540/30fps mp4; cached for the session.

    Every part lands on the same size/rate/codecs no matter what the source looks
    like — that is what lets the full-timeline preview concat them with stream copy
    instead of re-encoding everything again. Gaps become black (with silence) for
    exactly their duration, so what you watch is what Premiere will import.
    """
    r, dur = seg['row'], seg['dur']
    w, h = PREVIEW_SIZE
    if seg['kind'] == 'gap':
        key = ('gap', round(dur, 3))
    else:
        key = (str(r.src), round(seg['start'], 3), round(seg['end'], 3))
    part = _watch_parts.get(key)
    if part and part.exists():
        return part
    part = watch_dir() / f'part_{hashlib.sha256(repr(key).encode()).hexdigest()[:12]}.mp4'
    common = ['-t', f'{dur:.3f}', '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '28',
              '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '128k',
              '-ac', '2', '-ar', '48000', '-movflags', '+faststart', str(part)]
    if seg['kind'] == 'gap':
        cmd = [ffmpeg, '-y', '-hide_banner', '-loglevel', 'error',
               '-f', 'lavfi', '-i', f'color=black:s={w}x{h}:r={PREVIEW_FPS}:d={dur:.3f}',
               '-f', 'lavfi', '-t', f'{dur:.3f}', '-i', 'anullsrc=r=48000:cl=stereo',
               '-shortest', *common]
    else:
        has_audio = probes[str(r.src)]['has_audio']
        vf = (f'scale={w}:{h}:force_original_aspect_ratio=decrease,'
              f'pad={w}:{h}:(ow-iw)/2:(oh-ih)/2,fps={PREVIEW_FPS}')
        cmd = [ffmpeg, '-y', '-hide_banner', '-loglevel', 'error',
               '-ss', f'{seg["start"]:.3f}', '-i', str(r.src)]
        if not has_audio:
            cmd += ['-f', 'lavfi', '-t', f'{dur:.3f}', '-i', 'anullsrc=r=48000:cl=stereo']
        cmd += ['-vf', vf, '-map', '0:v:0', '-map', f'{"0" if has_audio else "1"}:a:0',
                '-shortest', *common]
    subprocess.run(cmd, capture_output=True, text=True, check=True)
    _watch_parts[key] = part
    return part


def render_preview_video(segs, probes, ffmpeg, log=lambda m: None):
    """Concat every placed segment into one watchable low-res file; returns its path."""
    segs = [s for s in segs if s['dur'] is not None]
    if not segs:
        raise ValueError('Nothing to watch — no placeable rows')
    total = sum(s['dur'] for s in segs)
    log(f'Rendering {fmt_tc(total)} preview at {PREVIEW_SIZE[1]}p...')
    parts = []
    for i, seg in enumerate(segs, 1):
        r = seg['row']
        what = 'gap' if seg['kind'] == 'gap' else (r.label or r.game or r.src.name)
        slow = ' — whole game, this one takes a while' if seg['dur'] > 120 else ''
        log(f'[preview {i}/{len(segs)}] {what} ({fmt_tc(seg["dur"])}){slow}')
        parts.append(encode_preview_part(ffmpeg, seg, probes))
    lst = watch_dir() / 'concat.txt'
    lst.write_text(''.join(f"file '{p}'\n" for p in parts), encoding='utf-8')
    out = watch_dir() / 'timeline_preview.mp4'
    subprocess.run([ffmpeg, '-y', '-hide_banner', '-loglevel', 'error',
                    '-f', 'concat', '-safe', '0', '-i', str(lst),
                    '-c', 'copy', '-movflags', '+faststart', str(out)],
                   capture_output=True, text=True, check=True)
    return out


# ------------------------------------------------------------ whistle check

def _pcm(ffmpeg, src, start, dur, filters=None):
    """Mono 8kHz 16-bit samples for a window of the source's audio."""
    import array
    cmd = [ffmpeg, '-v', 'error', '-ss', f'{max(start, 0):.2f}', '-i', str(src),
           '-t', f'{dur:.2f}', '-vn', '-ac', '1', '-ar', '8000']
    if filters:
        cmd += ['-af', filters]
    out = subprocess.run(cmd + ['-f', 's16le', '-'],
                         capture_output=True, check=True).stdout
    a = array.array('h')
    a.frombytes(out[:len(out) // 2 * 2])
    return a


def whistle_score(ffmpeg, src, sec, back=1.0, ahead=4.0):
    """How whistle-like is the loudest moment near `sec`? 0..~1, or None.

    A referee's whistle is a LOUD, sustained, near-pure tone around 2-4.5kHz —
    in a 0.1s window it dominates the signal, where crowd/announcer energy sits
    mostly below 2kHz. Score = the best high-band/full-band energy ratio of any
    non-quiet window in [sec-back, sec+ahead]. A goal clip is cut to end at the
    goal, so the whistle lands just after `sec` = the clip's end.

    Returns None when the window has no usable audio (no audio stream at all,
    or a track of pure digital silence) — no verdict, not a failing one.
    """
    dur = back + ahead
    try:
        full = _pcm(ffmpeg, src, sec - back, dur)
    except subprocess.CalledProcessError:
        return None                                  # no audio stream to decode
    n, win = len(full), 800                          # 0.1s at 8kHz
    if n < win:
        return None
    energies = [sum(x * x for x in full[i:i + win]) / win
                for i in range(0, n - win, win // 2)]
    if max(energies) < 1e4:
        return None                                  # silent track: nothing to hear
    band = _pcm(ffmpeg, src, sec - back, dur, 'highpass=f=2200,lowpass=f=4500')
    n = min(n, len(band))
    loud = sorted(energies)[len(energies) // 2]
    best = 0.0
    for i in range(0, n - win, win // 2):
        fe = sum(x * x for x in full[i:i + win]) / win
        if fe < max(loud * 0.5, 1e4):                # skip lulls: whistles are loud
            continue
        be = sum(x * x for x in band[i:i + win]) / win
        best = max(best, be / fe)
    return best


# Tuned on this library's real tape: sampled true goal cuts score 0.22-0.60; a
# random moment scores 0 about half the time (lacrosse is full of whistles, so
# presence is weak evidence — ABSENCE after a "goal" is what's worth flagging).
WHISTLE_MIN = 0.2


def open_media(path):
    """Open a file in the OS default video player."""
    if sys.platform == 'darwin':
        subprocess.Popen(['open', str(path)])
    elif os.name == 'nt':
        os.startfile(str(path))
    else:
        subprocess.Popen(['xdg-open', str(path)])


# ------------------------------------------------------------ label graphics

LABEL_POSITIONS = ('top left', 'top center', 'top right',
                   'middle center',
                   'bottom left', 'bottom center', 'bottom right')

# A position can also be a free point: '17.4%,83.4%' puts the text's left edge
# 17.4% in from the left with its bottom 83.4% down — how Premiere places text.
POS_PCT_RE = re.compile(r'\s*([\d.]+)%\s*,\s*([\d.]+)%\s*$')

# The user's own Premiere template (Darsh_template.prproj), decoded from its
# "GOAL VS TEAM..." text layer: Futura-Medium at 122px on a 3840x2160 sequence,
# text at (17.4%, 82%) INSIDE a graphic that Motion then scales 141% about
# center at (0.494, 0.5). Net effect on screen: ~172px (= 86 at 1080p) with the
# text tucked into the bottom-left corner at (3.4%, ~96%).
DARSH_STYLE = {
    'size': 86, 'color': '#ffffff', 'pos': '3.4%,96%',
    'fonts': ('/System/Library/Fonts/Supplemental/Futura.ttc',
              '/Library/Fonts/Futura.ttc',
              'C:/Windows/Fonts/Futura.ttc', 'C:/Windows/Fonts/futura.ttf',
              'C:/Windows/Fonts/futuram.ttf'),
}


def style_font(style=DARSH_STYLE):
    """First of the style's font files present on this machine, or ''."""
    for p in style['fonts']:
        if Path(p).exists():
            return p
    return ''


def render_label_png(text, font_path, size, width, height, out_path,
                     color='#ffffff', pos='bottom left'):
    """Full-frame transparent PNG with the label — a Premiere overlay still.

    `size` means pixels at 1080p; scaled proportionally for other sequence heights.
    `pos` is '<top|middle|bottom> <left|center|right>'; `color` any hex/named color.
    """
    from PIL import Image, ImageColor, ImageDraw, ImageFont
    px = max(8, round(size * height / 1080))
    try:
        font = ImageFont.truetype(str(font_path), px) if font_path else ImageFont.load_default(px)
    except Exception:
        font = ImageFont.load_default(px)
    try:
        rgb = ImageColor.getrgb(color)[:3]
    except ValueError:
        rgb = (255, 255, 255)
    img = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    margin = round(height * 0.05)
    stroke = max(2, px // 24)  # slim outline so the text survives busy footage
    l, t, r, b = d.textbbox((0, 0), text, font=font, stroke_width=stroke)
    tw, th = r - l, b - t
    if m := POS_PCT_RE.fullmatch(pos or ''):
        x = round(width * float(m.group(1)) / 100)
        y = round(height * float(m.group(2)) / 100) - th
    else:
        vert, _, horiz = (pos or '').partition(' ')
        x = {'left': margin, 'center': (width - tw) // 2,
             'right': width - margin - tw}.get(horiz, margin)
        y = {'top': margin, 'middle': (height - th) // 2,
             'bottom': height - margin - th}.get(vert, height - margin - th)
    x = max(0, min(x, width - tw))     # never let a nudge push the text off-frame
    y = max(0, min(y, height - th))
    # the outline flips to keep contrast: dark behind bright text, light behind dark
    dark_text = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2] < 100
    d.text((x - l, y - t), text, font=font, fill=rgb + (255,), stroke_width=stroke,
           stroke_fill=(255, 255, 255, 220) if dark_text else (0, 0, 0, 200))
    img.save(out_path)


# -------------------------------------------------------------- timeline XML

def _rate_xml(fps):
    tb = round(fps)
    ntsc = 'TRUE' if abs(fps - tb) > 0.01 else 'FALSE'
    return f'<rate><timebase>{tb}</timebase><ntsc>{ntsc}</ntsc></rate>'


def build_xmeml(rows, probes, sequence_name, label_pngs=None):
    """FCP7 XML (xmeml v4) — the interchange format Premiere imports natively.

    rows: matched Rows (src set, times resolved); probes: {path_str: probe info}.
    Whole-file rows use the entire source. Notes become sequence markers.
    Rows toggled off (or with no video but a known length) stay as GAPS: empty track
    for exactly their duration, so nothing after them shifts, plus a marker naming
    what belongs in the hole. Numbering counts gaps too — clip 07 stays 07 whether
    or not 05 is toggled off.
    label_pngs: optional {index_into_placeable_rows: png_path} — full-frame stills
    laid on a second video track above each clip (imports far more reliably than
    XML text generators).
    """
    segs = timeline_layout(rows, probes)
    clip_probes = [probes[str(s['row'].src)] for s in segs if s['kind'] == 'clip']
    if not clip_probes:
        raise ValueError('Every clip is toggled off or missing its video — '
                         'the timeline would be all gap')
    fps_list = [p['fps'] for p in clip_probes]
    seq_fps = max(set(fps_list), key=fps_list.count)
    first = clip_probes[0]

    v_items, a_items, o_items, markers = [], [], [], []
    file_defs = {}   # path -> file id (emit full <file> once, then reference)
    png_defs = {}
    t = 0            # sequence playhead in frames
    vpos = apos = 0  # 1-based clip ordinal PER TRACK — links must use track position
    for n, seg in enumerate(segs, 1):
        r = seg['row']
        start, end = seg['start'], seg['end']
        if seg['dur'] is None:      # generate() probes every placed file; fail loud if not
            raise ValueError(f'whole-file row "{r.label or r.game}" was never probed — '
                             'cannot know how long its slot is')
        dur = round(seg['dur'] * seq_fps)
        name = escape(f'{n:02d} - {r.label or r.game}')
        if seg['kind'] == 'gap':
            # A toggled-off or videoless row still owns its slot: leave the track
            # empty for its whole duration and drop a marker saying what goes there.
            why = 'toggled off' if not r.enabled else 'no video'
            note = f' — {r.notes}' if r.notes else ''
            markers.append(f'<marker><name>{name}</name>'
                           f'<comment>{escape(f"gap ({why}){note}")}</comment>'
                           f'<in>{t}</in><out>-1</out></marker>')
            t += dur
            continue
        p = probes[str(r.src)]
        # Every frame number ON THE CLIPITEM is read at the SEQUENCE rate, never at the
        # source's own rate, whatever <rate> the clipitem carries. Counting in/out in
        # source frames put every clip whose footage differs from the sequence at the
        # wrong time — 29.97 film landed at half its timecode, 119.88 at double, and a
        # late clip in a fast file pointed past the end of the media and imported as
        # nothing at all. Only footage that happened to match the sequence came in right.
        f_in, f_out = round(start * seq_fps), round(end * seq_fps)
        clip_frames = round(p['duration'] * seq_fps)
        file_frames = round(p['duration'] * p['fps'])   # <file> stays in its own rate
        path = str(r.src)

        if path not in file_defs:
            file_defs[path] = f'file-{len(file_defs) + 1}'
            audio_xml = ('<audio><samplecharacteristics><depth>16</depth>'
                         '<samplerate>48000</samplerate></samplecharacteristics>'
                         '<channelcount>2</channelcount></audio>') if p['has_audio'] else ''
            file_xml = (
                f'<file id="{file_defs[path]}"><name>{escape(r.src.name)}</name>'
                f'<pathurl>{escape(r.src.resolve().as_uri().replace("file:///", "file://localhost/"))}</pathurl>'
                f'{_rate_xml(p["fps"])}<duration>{file_frames}</duration>'
                f'<media><video><samplecharacteristics>{_rate_xml(p["fps"])}'
                f'<width>{p["width"]}</width><height>{p["height"]}</height>'
                f'</samplecharacteristics></video>{audio_xml}</media></file>')
        else:
            file_xml = f'<file id="{file_defs[path]}"/>'

        # <clipindex> is the clip's 1-based POSITION ON ITS TRACK, not the slot
        # number: gaps hold slots but put nothing on the track, and a video-only
        # clip advances the video track while the audio track stands still.
        # Counting slots sent every link past its clip and Premiere silently
        # dropped the whole import — a loading bar, then nothing.
        vpos += 1
        if p['has_audio']:
            apos += 1

        def links(vid, aid):
            return (f'<link><linkclipref>{vid}</linkclipref><mediatype>video</mediatype>'
                    f'<trackindex>1</trackindex><clipindex>{vpos}</clipindex></link>'
                    f'<link><linkclipref>{aid}</linkclipref><mediatype>audio</mediatype>'
                    f'<trackindex>1</trackindex><clipindex>{apos}</clipindex></link>')

        vid, aid = f'clipitem-v{n}', f'clipitem-a{n}'
        common = (f'<enabled>TRUE</enabled><duration>{clip_frames}</duration>'
                  f'{_rate_xml(seq_fps)}<start>{t}</start><end>{t + dur}</end>'
                  f'<in>{f_in}</in><out>{f_out}</out>')
        link_xml = links(vid, aid) if p['has_audio'] else ''
        v_items.append(f'<clipitem id="{vid}"><name>{name}</name>{common}{file_xml}'
                       f'<sourcetrack><mediatype>video</mediatype><trackindex>1</trackindex>'
                       f'</sourcetrack>{link_xml}</clipitem>')
        if p['has_audio']:
            a_items.append(f'<clipitem id="{aid}"><name>{name}</name>{common}'
                           f'<file id="{file_defs[path]}"/>'
                           f'<sourcetrack><mediatype>audio</mediatype><trackindex>1</trackindex>'
                           f'</sourcetrack>{links(vid, aid)}</clipitem>')
        png = (label_pngs or {}).get(n - 1)
        if png:
            ppath = str(png)
            if ppath not in png_defs:
                png_defs[ppath] = f'file-png{len(png_defs) + 1}'
                png_xml = (
                    f'<file id="{png_defs[ppath]}"><name>{escape(Path(ppath).name)}</name>'
                    f'<pathurl>{escape(Path(ppath).resolve().as_uri().replace("file:///", "file://localhost/"))}</pathurl>'
                    f'{_rate_xml(seq_fps)}<duration>108000</duration>'
                    f'<media><video><samplecharacteristics>{_rate_xml(seq_fps)}'
                    f'<width>{first["width"]}</width><height>{first["height"]}</height>'
                    f'</samplecharacteristics></video></media></file>')
            else:
                png_xml = f'<file id="{png_defs[ppath]}"/>'
            o_items.append(f'<clipitem id="clipitem-l{n}"><name>{name}</name>'
                           f'<enabled>TRUE</enabled><duration>108000</duration>{_rate_xml(seq_fps)}'
                           f'<start>{t}</start><end>{t + dur}</end><in>0</in><out>{dur}</out>'
                           f'{png_xml}</clipitem>')
        if r.notes:
            markers.append(f'<marker><name>{name}</name><comment>{escape(r.notes)}</comment>'
                           f'<in>{t}</in><out>-1</out></marker>')
        t += dur

    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE xmeml>\n'
        f'<xmeml version="4"><sequence id="sequence-1"><name>{escape(sequence_name)}</name>'
        f'<duration>{t}</duration>{_rate_xml(seq_fps)}<media>'
        f'<video><format><samplecharacteristics>{_rate_xml(seq_fps)}'
        f'<width>{first["width"]}</width><height>{first["height"]}</height>'
        f'<pixelaspectratio>square</pixelaspectratio></samplecharacteristics></format>'
        f'<track>{"".join(v_items)}</track>'
        + (f'<track>{"".join(o_items)}</track>' if o_items else '') + '</video>'
        f'<audio><track>{"".join(a_items)}</track></audio>'
        f'</media>{"".join(markers)}</sequence></xmeml>')


# ---------------------------------------------------------------- generation

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', name).strip(' .') or 'clip'


def generate(rows, out_dir, sequence_name, export_clips, log, labels_cfg=None):
    """Write timeline XML (+ optional clip files). Returns (xml_path, ok, failed).

    Rows toggled off — or with no video but a known length — stay as gaps in the
    timeline instead of shifting everything after them.
    labels_cfg: optional {'font': Path|None, 'size': int, 'team': bool, 'color': str,
    'pos': str} — render each row's label as a still-image overlay on a second
    timeline track; 'team' appends 'vs <game>' from the sheet.
    """
    ffmpeg, ffprobe = find_ffmpeg()
    rows = [r for r in rows if place_kind(r)]
    if not rows:
        raise ValueError('No usable rows — every row is missing a video file or timecode')

    log(f'Probing {len({str(r.src) for r in rows if r.src})} video file(s)...')
    probes, unreadable = {}, set()
    for path in {str(r.src) for r in rows if r.src}:
        try:
            probes[path] = probe(ffprobe, path)
        except Exception:
            unreadable.add(path)
            log(f'WARNING: could not read {Path(path).name}')
    for path, p in probes.items():
        if p['vfr']:
            log(f'NOTE: {Path(path).name} has a variable frame rate ({p["fps"]:.3f} fps '
                f'average) — timeline positions use the measured rate. If clips still land '
                f'late in Premiere, re-export this file at a constant frame rate.')
    kept = []
    for r in rows:
        if r.src and str(r.src) in unreadable:
            # an unreadable video must not shift every later clip — a timed row
            # still knows its own length, so it holds its slot as a gap
            r.src = None
            if place_kind(r):
                log(f'   ...its "{r.label or r.game}" clip stays as a gap in the timeline')
                kept.append(r)
            else:
                log(f'   ...its "{r.label or r.game}" clip is skipped (no timecode)')
            continue
        kept.append(r)
    rows = kept
    if not any(place_kind(r) == 'clip' for r in rows):
        raise ValueError('None of the matched video files could be read' if unreadable else
                         'Every clip is toggled off or missing its video — '
                         'the timeline would be all gap')
    kept = []
    for r in rows:  # enforce real durations now that we know them
        if place_kind(r) != 'clip':
            kept.append(r)      # gaps have no media to check
            continue
        d = probes[str(r.src)]['duration']
        if r.whole_file or not d:
            kept.append(r)
            continue
        if r.start >= d:
            if r.manual:
                # Hand-picked a short pre-cut clip for a row whose times belong to the
                # full game — the only sensible reading is "use this clip whole".
                log(f'NOTE: "{r.label or r.game}" — using all of {r.src.name} '
                    f'({fmt_tc(d)}); its times are past the end of this file')
                r.whole_file = True
                kept.append(r)
                continue
            # Wrong video, but the row knows its own length — hold its slot as a gap,
            # exactly like an unreadable file, so nothing after it shifts or renumbers.
            log(f'"{r.label or r.game}" starts at {fmt_tc(r.start)} but {r.src.name} is '
                f'only {fmt_tc(d)} long — wrong video or wrong time? Left as a gap')
            r.src = None
            kept.append(r)
            continue
        if r.end > d:
            log(f'NOTE: "{r.label or r.game}" trimmed to the end of {r.src.name}')
            r.end = d
        kept.append(r)
    rows = kept
    clips = [r for r in rows if place_kind(r) == 'clip']
    if not clips:
        raise ValueError('No usable rows left after checking video durations')
    if len(clips) < len(rows):
        log(f'{len(rows) - len(clips)} clip(s) left as gaps in the timeline '
            '(toggled off or no video)')

    label_pngs = {}
    if labels_cfg:
        try:
            ensure_pip('PIL', 'pillow')
        except ImportError:
            log('WARNING: Pillow missing and could not auto-install — labels skipped. '
                'Run: python3 -m pip install pillow')
            labels_cfg = None
    if labels_cfg:
        team_on = bool(labels_cfg.get('team'))
        first = probes[str(clips[0].src)]
        label_dir = out_dir / 'labels'
        label_dir.mkdir(exist_ok=True)
        by_text = {}
        for i, r in enumerate(rows):
            text = label_text(r, team_on) if place_kind(r) == 'clip' else ''
            if not text:
                continue
            if text not in by_text:
                png = label_dir / f'{sanitize_filename(text)}.png'
                try:
                    render_label_png(text, labels_cfg.get('font'), labels_cfg.get('size', 48),
                                     first['width'], first['height'], png,
                                     color=labels_cfg.get('color', '#ffffff'),
                                     pos=labels_cfg.get('pos', 'bottom left'))
                    by_text[text] = png
                except Exception as e:
                    by_text[text] = None
                    log(f'WARNING: could not render label "{text}": {e}')
            if by_text[text]:
                label_pngs[i] = by_text[text]
        if label_pngs:
            log(f'{sum(1 for v in by_text.values() if v)} label graphic(s) written to labels/')

    xml_path = out_dir / f'{sanitize_filename(sequence_name)}_timeline.xml'
    xml_path.write_text(build_xmeml(rows, probes, sequence_name, label_pngs), encoding='utf-8')
    log(f'Timeline written: {xml_path.name}  (Premiere: File > Import)')

    ok, failed = 0, []
    if export_clips:
        clip_dir = out_dir / 'clips'
        clip_dir.mkdir(exist_ok=True)
        done = 0
        # n numbers ALL rows, gaps included, so file names match the timeline's numbering
        for n, r in enumerate(rows, 1):
            if place_kind(r) != 'clip':
                continue
            done += 1
            stem = sanitize_filename(f'{n:02d} - {r.label or "clip"} - {norm_tokens(r.game) and " ".join(norm_tokens(r.game)) or r.src.stem}')
            out_path = clip_dir / f'{stem}{r.src.suffix.lower() if r.whole_file else ".mp4"}'
            try:
                cut_clip(ffmpeg, r, out_path)
                if not out_path.exists() or out_path.stat().st_size < 1000:
                    raise RuntimeError('ffmpeg produced an empty file')
                ok += 1
                log(f'[{done}/{len(clips)}] {out_path.name}')
            except Exception as e:
                failed.append(r)
                err = str(getattr(e, 'stderr', '') or e).strip()
                log(f'[{done}/{len(clips)}] FAILED {out_path.name}: {err[:200]}')
    return xml_path, ok, failed


def estimate_clip_bytes(rows):
    """Rough output size at ~10 Mbps for cut clips, actual size for whole-file copies."""
    total = 0
    for r in rows:
        if place_kind(r) != 'clip':
            continue            # gaps and toggled-off rows are never exported
        if r.whole_file:
            total += r.src.stat().st_size
        else:
            total += int((r.end - r.start) * 1.25e6)
    return total


# ---------------------------------------------------------------- usage ping

# Usage stats, documented in README.md: when TRACK_URL is set, loading a spreadsheet
# submits its FILE NAME + a short content hash and this computer's name to a Google Form
# owned by the app's developer, once per sheet per machine. Sheet contents never leave
# the machine. Opt out any time: set the environment variable CLIP_CHOPPER_NO_TRACK=1.
TRACK_URL = ('https://docs.google.com/forms/d/e/'
             '1FAIpQLSe56Lj6Ipmnb7-PpnMEdczWWgz701ruhh4yqzTa85iPp7D0SQ/formResponse')
TRACK_FIELDS = {'sheet': 'entry.1140511568', 'machine': 'entry.584151074'}


def sheet_fingerprint(grid):
    payload = '\n'.join('\t'.join(row) for row in grid)
    return hashlib.sha256(payload.encode('utf-8')).hexdigest()[:16]


def machine_fingerprint():
    raw = platform.node() + str(Path.home()) + sys.platform
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()[:12]


def send_usage_ping(sheet_id, log=None):
    """Fire-and-forget anonymous ping (hashed ids only). Never blocks or raises."""
    if not TRACK_URL or os.environ.get('CLIP_CHOPPER_NO_TRACK'):
        return

    def post():
        try:
            data = urllib.parse.urlencode({
                TRACK_FIELDS['sheet']: sheet_id,
                TRACK_FIELDS['machine']: f'{platform.node()} ({machine_fingerprint()})'}).encode()
            req = urllib.request.Request(TRACK_URL, data=data,
                                         headers={'User-Agent': 'Mozilla/5.0'})
            urllib.request.urlopen(req, timeout=5).read()
        except Exception:
            pass

    threading.Thread(target=post, daemon=True).start()
    if log:
        log('(usage ping sent — sheet name only, no contents; set CLIP_CHOPPER_NO_TRACK=1 to disable)')


# ------------------------------------------------------------------- GUI

SETTINGS_PATH = Path.home() / '.clip_chopper.json'


def load_settings():
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding='utf-8'))
    except Exception:
        return {}


def save_settings(d):
    try:
        SETTINGS_PATH.write_text(json.dumps(d), encoding='utf-8')
    except Exception:
        pass


def fmt_tc(sec):
    if sec is None:
        return ''
    sec = int(round(sec))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f'{h}:{m:02d}:{s:02d}' if h else f'{m}:{s:02d}'


def run_gui():
    import tkinter as tk
    from tkinter import colorchooser, filedialog, messagebox, ttk
    DND_FILES, dnd_error = None, None
    try:
        ensure_pip('tkinterdnd2', 'tkinterdnd2')   # may be missing if run outside the launcher
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        has_dnd = True
    except Exception as e:
        root = tk.Tk()
        has_dnd = False
        dnd_error = f'{type(e).__name__}: {e}'

    root.title('Clip Chopper')
    root.geometry('1100x640')
    state = {'rows': [], 'sheet': None, 'videos': [], 'video_files': []}
    log_q = queue.Queue()

    # --- top: spreadsheet + videos folder pickers -------------------------
    top = ttk.Frame(root, padding=8)
    top.pack(fill='x')
    drop_text = ('Drop spreadsheet or video folders here (or use the buttons)' if has_dnd
                 else 'Choose spreadsheet (Browse…)')
    sheet_lbl = ttk.Label(top, text=drop_text, relief='groove', anchor='center', padding=10)
    sheet_lbl.grid(row=0, column=0, columnspan=2, sticky='ew', padx=(0, 6))
    ttk.Button(top, text='Browse…', command=lambda: pick_sheet()).grid(row=0, column=2)
    ttk.Label(top, text='or Google Sheets URL:').grid(row=1, column=0, sticky='w', pady=(6, 0))
    url_var = tk.StringVar()
    url_entry = ttk.Entry(top, textvariable=url_var, width=60)
    url_entry.grid(row=1, column=1, sticky='ew', pady=(6, 0), padx=(0, 6))
    ttk.Button(top, text='Load URL', command=lambda: load_url()).grid(row=1, column=2, pady=(6, 0))
    vid_lbl = ttk.Label(top, text='Video folders: (none yet — drop folders here or Add folder…)',
                        anchor='w')
    vid_lbl.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(6, 0))
    vid_btns = ttk.Frame(top)
    vid_btns.grid(row=2, column=2, pady=(6, 0))
    ttk.Button(vid_btns, text='Add folder…', command=lambda: pick_videos()).pack(side='left')
    ttk.Button(vid_btns, text='Clear', width=6,
               command=lambda: clear_videos()).pack(side='left', padx=(4, 0))
    # age-group picker — appears only when the scanned videos span several age groups
    age_var = tk.StringVar(value='All age groups')
    age_lbl = ttk.Label(vid_btns, text='Age:')
    age_box = ttk.Combobox(vid_btns, textvariable=age_var, state='readonly', width=9)
    top.columnconfigure(1, weight=1)

    # --- middle: review table --------------------------------------------
    mid = ttk.Frame(root, padding=(8, 0))
    mid.pack(fill='both', expand=True)
    cols = ('on', 'order', 'game', 'file', 'in', 'out', 'label', 'notes', 'status')
    tree = ttk.Treeview(mid, columns=cols, show='headings', selectmode='extended')
    widths = {'on': 34, 'order': 45, 'game': 180, 'file': 170, 'in': 70, 'out': 70,
              'label': 160, 'notes': 170, 'status': 220}
    for c in cols:
        tree.heading(c, text=c.title())
        tree.column(c, width=widths[c], anchor='w')
    tree.tag_configure('bad', background='#ffd6d6')
    tree.tag_configure('off', foreground='#999')
    ys = ttk.Scrollbar(mid, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=ys.set)
    tree.pack(side='left', fill='both', expand=True)
    ys.pack(side='right', fill='y')

    # --- manual assignment bar (select rows, give them a video) -----------
    assign = ttk.Frame(root, padding=(8, 4))
    assign.pack(fill='x')
    ttk.Label(assign, text='Selected rows:').pack(side='left')
    ttk.Button(assign, text='Assign video…',
               command=lambda: assign_selected()).pack(side='left', padx=(6, 0))
    ttk.Button(assign, text='Clear assignment',
               command=lambda: unassign_selected()).pack(side='left', padx=(4, 0))
    ttk.Label(assign, foreground='#666',
              text='(shift/ctrl-click to pick several · right-click for the same menu · '
                   'double-click any cell to edit it)').pack(side='left', padx=(10, 0))

    # --- bottom: options + generate + log ---------------------------------
    cfg = load_settings()
    bot = ttk.Frame(root, padding=8)
    bot.pack(fill='x')

    # label-overlay options
    labels_var = tk.BooleanVar(value=cfg.get('labels_on', False))
    font_var = tk.StringVar(value=cfg.get('font', ''))
    size_var = tk.StringVar(value=str(cfg.get('size', 48)))
    team_var = tk.BooleanVar(value=cfg.get('label_team', False))
    color_var = tk.StringVar(value=cfg.get('label_color', '#ffffff'))
    pos_var = tk.StringVar(value=cfg.get('label_pos', 'bottom left'))
    ttk.Checkbutton(bot, text='Add clip labels as text in the timeline',
                    variable=labels_var).grid(row=0, column=0, sticky='w')
    font_frame = ttk.Frame(bot)
    font_frame.grid(row=0, column=1, columnspan=2, sticky='w', padx=(10, 0))
    font_lbl = ttk.Label(font_frame, width=20, anchor='w',
                         text=Path(font_var.get()).name if font_var.get() else '(default font)')
    def pick_font():
        p = filedialog.askopenfilename(title='Label font',
                                       filetypes=[('Fonts', '*.ttf *.otf'), ('All files', '*.*')])
        if p:
            font_var.set(p)
            font_lbl.configure(text=Path(p).name)
    ttk.Button(font_frame, text='Font…', command=pick_font).pack(side='left')
    font_lbl.pack(side='left', padx=6)
    ttk.Label(font_frame, text='Size:').pack(side='left')
    ttk.Entry(font_frame, textvariable=size_var, width=4).pack(side='left', padx=(4, 0))
    ttk.Checkbutton(font_frame, text='+ team ("Goal vs Sweetlax")',
                    variable=team_var).pack(side='left', padx=(10, 0))
    ttk.Button(font_frame, text='Preview…',
               command=lambda: preview_label()).pack(side='left', padx=(8, 0))
    ttk.Button(font_frame, text='Mogrts…',
               command=lambda: make_mogrts()).pack(side='left', padx=(4, 0))

    export_var = tk.BooleanVar(value=cfg.get('export_clips', True))
    ttk.Style().configure('Big.TCheckbutton', font=('TkDefaultFont', 10, 'bold'))
    export_chk = ttk.Checkbutton(bot, text='Make individual clip files',
                                 variable=export_var, style='Big.TCheckbutton')
    export_chk.grid(row=1, column=0, sticky='w', pady=(6, 0))
    btns = ttk.Frame(bot)
    btns.grid(row=1, column=1, padx=10, pady=(6, 0))
    prev_btn = ttk.Button(btns, text='Preview timeline', state='disabled',
                          command=lambda: preview_timeline())
    prev_btn.pack(side='left')
    gen_btn = ttk.Button(btns, text='Generate', state='disabled', command=lambda: start_generate())
    gen_btn.pack(side='left', padx=(6, 0))
    prog = ttk.Progressbar(bot, mode='determinate')
    prog.grid(row=1, column=2, sticky='ew', padx=(0, 4), pady=(6, 0))
    bot.columnconfigure(2, weight=1)
    log_box = tk.Text(bot, height=6, state='disabled', wrap='none')
    log_box.grid(row=2, column=0, columnspan=3, sticky='ew', pady=(6, 0))

    def log(msg):
        log_q.put(str(msg))

    def drain_log():
        try:
            while True:
                msg = log_q.get_nowait()
                log_box.configure(state='normal')
                log_box.insert('end', msg + '\n')
                log_box.see('end')
                log_box.configure(state='disabled')
        except queue.Empty:
            pass
        root.after(150, drain_log)

    def refresh_table():
        keep = tree.selection()
        tree.delete(*tree.get_children())
        for i, r in enumerate(state['rows']):
            vals = ('☑' if r.enabled else '☐',
                    r.order if r.order is not None else '',
                    r.game,
                    r.src.name if r.src else '',
                    'whole file' if r.whole_file else fmt_tc(r.start),
                    'whole file' if r.whole_file else fmt_tc(r.end),
                    r.label, r.notes,
                    'off — stays as a gap in the timeline' if not r.enabled
                    else '; '.join(r.flags) if r.flags else 'ok')
            tree.insert('', 'end', iid=str(i), values=vals,
                        tags=('off',) if not r.enabled else ('bad',) if r.flags else ())
        still = [i for i in keep if tree.exists(i)]
        if still:
            tree.selection_set(still)
        n_bad = sum(1 for r in state['rows'] if r.flags)
        if state['rows']:
            est = estimate_clip_bytes(state['rows']) / 1e6
            export_chk.configure(text=f'Make individual clip files (~{est:.0f} MB)')
            log(f'{len(state["rows"])} clips parsed'
                + (f', {n_bad} need review (red rows — double-click to fix)' if n_bad else ', all ok'))
        gen_btn.configure(state='normal' if state['rows'] and state['videos'] else 'disabled')
        prev_btn.configure(state='normal' if state['rows'] else 'disabled')

    def selected_age():
        v = age_var.get()
        return v if re.fullmatch(r'20\d\d', v) else ''

    def rematch():
        if state['rows'] and state['videos']:
            for r in state['rows']:
                if r.manual and r.src:
                    continue
                r.src = None
                r.flags = [f for f in r.flags if 'match' not in f and 'video' not in f
                           and 'game name' not in f]
            exclude = state['sheet'].parent / 'clips' if state['sheet'] else None
            match_videos(state['rows'], state['videos'], exclude=exclude,
                         videos=filter_age_group(state['video_files'], selected_age()))
        refresh_table()

    def update_age_choices():
        """Offer the dropdown only when it can actually disambiguate something."""
        groups = list_age_groups(state['video_files'])
        if len(groups) >= 2:
            age_box.configure(values=['All age groups'] + groups)
            saved = cfg.get('age_group', '')
            age_var.set(saved if saved in groups else 'All age groups')
            age_lbl.pack(side='left', padx=(10, 2))
            age_box.pack(side='left')
        else:
            age_var.set('All age groups')
            age_lbl.pack_forget()
            age_box.pack_forget()

    def on_age_change(_e=None):
        cfg['age_group'] = selected_age()
        save_settings(cfg)
        group = selected_age()
        if group:
            n = len(filter_age_group(state['video_files'], group))
            log(f'Age group {group}: matching against {n} of '
                f'{len(state["video_files"])} videos')
        else:
            log(f'Age group filter off — matching against all {len(state["video_files"])} videos')
        rematch()
    age_box.bind('<<ComboboxSelected>>', on_age_change)

    def load_sheet(path):
        try:
            state['rows'] = parse_sheet(path)
            state['sheet'] = Path(path)
            sheet_lbl.configure(text=Path(path).name)
            rematch()
        except Exception as e:
            messagebox.showerror('Could not read spreadsheet', str(e))
            return
        fp = sheet_fingerprint(load_grid(path))
        if fp not in cfg.get('pinged', []):
            send_usage_ping(f'{Path(path).stem} ({fp[:8]})', log)
            cfg['pinged'] = (cfg.get('pinged', []) + [fp])[-100:]
            save_settings(cfg)

    def pick_sheet():
        p = filedialog.askopenfilename(filetypes=[('Spreadsheets', '*.xlsx *.csv'), ('All files', '*.*')])
        if p:
            load_sheet(p)

    def load_url():
        try:
            load_sheet(fetch_google_sheet(url_var.get()))
        except Exception as e:
            messagebox.showerror('Google Sheets', str(e))

    def add_videos(paths):
        """Add one or more video folders (subfolders are always searched too)."""
        added = [Path(p) for p in paths if Path(p).is_dir()]
        new = [p for p in added if p not in state['videos']]
        state['videos'] = state['videos'] + new
        if not state['videos']:
            state['video_files'] = []
            vid_lbl.configure(text='Video folders: (none yet — drop folders here or Add folder…)')
            update_age_choices()
            refresh_table()
            return
        names = ', '.join(p.name or str(p) for p in state['videos'])
        vid_lbl.configure(text=f'Video folders: {names}  (scanning…)')
        root.update_idletasks()
        state['video_files'] = list_videos(state['videos'])   # scanned once, reused by rematch
        vid_lbl.configure(text=f'Video folders: {names}  ({len(state["video_files"])} videos found)')
        update_age_choices()
        rematch()

    def pick_videos():
        p = filedialog.askdirectory(title='Folder with the game videos (subfolders included)')
        if p:
            add_videos([p])

    def clear_videos():
        state['videos'] = []
        add_videos([])

    def assign_file(rows_to_set, title):
        """Ask for a video and pin it to these rows — manual picks survive re-matching."""
        if not rows_to_set:
            log('Select a row in the table first, then Assign video…')
            return
        p = filedialog.askopenfilename(
            title=title,
            initialdir=state['videos'][0] if state['videos'] else '.',
            filetypes=[('Videos', ' '.join(f'*{e}' for e in VIDEO_EXTS)), ('All files', '*.*')])
        if not p:
            return
        for r in rows_to_set:
            r.src = Path(p)
            r.manual = True
            r.flags = [f for f in r.flags if 'match' not in f and 'video' not in f
                       and 'game name' not in f]
        log(f'Assigned {Path(p).name} to {len(rows_to_set)} row(s)')
        refresh_table()

    def selected_rows():
        return [state['rows'][int(i)] for i in tree.selection() if i.isdigit()]

    def assign_selected():
        rows_sel = selected_rows()
        assign_file(rows_sel, f'Video for {len(rows_sel)} selected row(s)')

    def unassign_selected():
        rows_sel = selected_rows()
        if not rows_sel:
            log('Select a row in the table first, then Clear assignment')
            return
        for r in rows_sel:
            r.src, r.manual = None, False
        rematch()   # let auto-matching have another go at them
        log(f'Cleared {len(rows_sel)} assignment(s)')

    def toggle_rows(rows_sel):
        if not rows_sel:
            log('Select a row in the table first')
            return
        turn_on = not rows_sel[0].enabled     # first row decides, so a mixed batch settles
        for r in rows_sel:
            r.enabled = turn_on
        refresh_table()

    def on_toggle_click(event):
        """Click the ☑ column to toggle a clip off — it stays as a gap in the timeline."""
        if tree.identify('region', event.x, event.y) != 'cell':
            return
        if tree.identify_column(event.x) != '#1':
            return
        item = tree.identify_row(event.y)
        if item and item.isdigit():
            toggle_rows([state['rows'][int(item)]])

    def on_right_click(event):
        item = tree.identify_row(event.y)
        if item and item not in tree.selection():
            tree.selection_set(item)
        menu = tk.Menu(tree, tearoff=0)
        menu.add_command(label='Watch this clip (low-res)',
                         command=lambda: watch_rows(selected_rows()))
        menu.add_command(label='Toggle clip on/off (leaves a gap)',
                         command=lambda: toggle_rows(selected_rows()))
        menu.add_command(label='Assign video to selected row(s)…', command=assign_selected)
        menu.add_command(label='Clear assignment', command=unassign_selected)
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

    def on_double_click(event):
        item, col = tree.identify_row(event.y), tree.identify_column(event.x)
        if not item:
            return
        r = state['rows'][int(item)]
        col_name = cols[int(col[1:]) - 1]
        if col_name == 'file':
            # one row's file usually means the whole game's file, unless several rows
            # are deliberately selected or the game name is blank
            sel = selected_rows()
            if len(sel) > 1 and r in sel:
                targets = sel
            elif r.game:
                targets = [o for o in state['rows'] if o.game == r.game]
            else:
                targets = [r]
            assign_file(targets, f'Video for "{r.game}"' if r.game else 'Video for this clip')
        elif col_name in ('in', 'out', 'label', 'notes'):
            edit_cell(item, col, r, col_name)

    def edit_cell(item, col, r, col_name):
        x, y, w, h = tree.bbox(item, col)
        entry = tk.Entry(tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, tree.set(item, col_name))
        entry.focus_set()

        def commit(_=None):
            val = entry.get().strip()
            entry.destroy()
            if col_name in ('in', 'out'):
                sec = parse_tc(val) if ':' in val else (float(val) if re.fullmatch(r'\d+(\.\d+)?', val) else None)
                if sec is None:
                    return
                setattr(r, 'start' if col_name == 'in' else 'end', float(sec))
                r.whole_file = False
                r.flags = [f for f in r.flags if 'time' not in f.lower()]
                if r.start is None or r.end is None:
                    r.flags.append('set both In and Out times')
                elif r.end <= r.start:
                    r.flags.append('end is before start — check times')
            else:
                setattr(r, 'label' if col_name == 'label' else 'notes', val)
            refresh_table()

        entry.bind('<Return>', commit)
        entry.bind('<FocusOut>', commit)
        entry.bind('<Escape>', lambda e: entry.destroy())

    def start_generate():
        placeable = [r for r in state['rows'] if place_kind(r)]
        clips = [r for r in placeable if place_kind(r) == 'clip']
        auto_gap = [r for r in placeable if r.enabled and place_kind(r) == 'gap']
        dropped = len(state['rows']) - len(placeable)
        if not clips:
            log('Nothing to generate — every clip is toggled off or missing its video')
            return
        warn = []
        if auto_gap:
            warn.append(f'{len(auto_gap)} row(s) have no video — '
                        'they will be left as gaps in the timeline.')
        if dropped:
            warn.append(f"{dropped} row(s) can't be placed at all — no valid In/Out times "
                        'and no whole-file video — left out entirely.')
        if warn and not messagebox.askyesno(
                'Some rows have problems',
                '\n'.join(warn) + f'\nContinue with {len(clips)} clip(s)?'):
            return
        out_dir = state['sheet'].parent
        tmp = Path(tempfile.gettempdir())
        if out_dir == tmp or tmp in out_dir.parents:  # e.g. Google Sheets download
            p = filedialog.askdirectory(title='Where should the timeline and clips be saved?')
            if not p:
                return
            out_dir = Path(p)
        gen_btn.configure(state='disabled')
        prog.configure(maximum=max(len(clips), 1), value=0)
        work_rows = copy.deepcopy(placeable)  # worker gets its own rows; table edits can't race it
        name = state['sheet'].stem.replace('_', ' ')
        export = export_var.get()

        try:
            lsize = max(8, int(float(size_var.get())))
        except ValueError:
            lsize = 48
        labels_cfg = None
        if labels_var.get():
            lfont = Path(font_var.get()) if font_var.get() else None
            if lfont and not lfont.exists():
                log(f'Font file not found ({lfont}) — using default font')
                lfont = None
            labels_cfg = {'font': lfont, 'size': lsize, 'team': team_var.get(),
                          'color': color_var.get(), 'pos': pos_var.get()}
        cfg.update({'labels_on': labels_var.get(), 'font': font_var.get(),
                    'size': lsize, 'export_clips': export, 'label_team': team_var.get(),
                    'label_color': color_var.get(), 'label_pos': pos_var.get()})
        save_settings(cfg)

        def work():
            try:
                done = [0]

                def counting_log(msg):
                    log(msg)
                    if msg.startswith('['):
                        done[0] += 1
                        root.after(0, lambda v=done[0]: prog.configure(value=v))
                xml_path, ok, failed = generate(work_rows, out_dir, name, export,
                                                counting_log, labels_cfg)
                log(f'Done. Timeline: {xml_path}')
                if export:
                    log(f'Clips: {ok} written' + (f', {len(failed)} failed' if failed else ''))
            except Exception as e:
                log(f'ERROR: {e}')
            finally:
                root.after(0, lambda: gen_btn.configure(state='normal'))

        threading.Thread(target=work, daemon=True).start()

    def cached_probes(rows):
        return {str(r.src): _probe_cache[str(r.src)] for r in rows
                if r.src and str(r.src) in _probe_cache}

    watch_busy = {'on': False}

    def watch_rows(target_rows=None):
        """Render the actual footage at 540p and open it in the video player.

        target_rows=None watches the WHOLE timeline — clips in order with black
        where the gaps are; a list watches just those rows' clips.
        """
        if watch_busy['on']:
            log('A preview render is already running — wait for it to finish')
            return
        watch_busy['on'] = True

        def wk():
            try:
                ffmpeg, ffprobe = find_ffmpeg()
                rows = state['rows'] if target_rows is None else target_rows
                for r in rows:
                    if r.src and str(r.src) not in _probe_cache:
                        try:
                            probe(ffprobe, r.src)
                        except Exception:
                            pass
                if target_rows is not None:
                    # watching specific rows: show their footage even if toggled off —
                    # checking what's in a clip is exactly why you'd watch it
                    rows = [copy.copy(r) for r in rows]
                    for r in rows:
                        r.enabled = True
                segs = timeline_layout(rows, cached_probes(rows))
                if target_rows is not None:     # single clips: footage only, no gaps
                    segs = [s for s in segs if s['kind'] == 'clip']
                segs = [s for s in segs if s['dur'] is not None]
                if not segs:
                    log('Nothing to watch — that row has no video or timecode')
                    return
                if len(segs) == 1 and segs[0]['kind'] == 'clip':
                    s = segs[0]
                    log(f'Cutting "{s["row"].label or s["row"].game}" at 540p...')
                    out = encode_preview_part(ffmpeg, s, _probe_cache)
                else:
                    out = render_preview_video(segs, _probe_cache, ffmpeg, log)
                open_media(out)
                log('Preview opened in your video player')
            except Exception as e:
                log(f'Preview render failed: {e}')
            finally:
                watch_busy['on'] = False

        threading.Thread(target=wk, daemon=True).start()

    def preview_timeline():
        """Draw the sequence as it will import — clips, gaps, thumbnails — no Premiere needed.

        Uses the same timeline_layout() the XML builder uses, so what this shows is by
        construction what Generate writes.
        """
        win = tk.Toplevel(root)
        win.title('Timeline preview')
        win.geometry('1020x270')
        alive = {'ok': True, 'tdir': None}
        thumbs = {}                       # id(row) -> PhotoImage; ref MUST live on
        win._thumb_refs = thumbs          # or Tk garbage-collects the images off the canvas
        bar = ttk.Frame(win)
        bar.pack(side='bottom', fill='x')
        ttk.Button(bar, text='▶ Watch video',
                   command=lambda: watch_rows(None)).pack(side='right', padx=8, pady=4)
        info = ttk.Label(bar, anchor='w', padding=(8, 4))
        info.pack(side='left', fill='x', expand=True)
        xs = ttk.Scrollbar(win, orient='horizontal')
        canvas = tk.Canvas(win, height=200, background='#20242a', highlightthickness=0,
                           xscrollcommand=xs.set)
        xs.configure(command=canvas.xview)
        xs.pack(side='bottom', fill='x')
        canvas.pack(fill='both', expand=True)

        def draw():
            if not alive['ok']:
                return
            canvas.delete('all')
            rows = state['rows']          # read live: table edits and reloads show up
            rowpos = {id(r): i for i, r in enumerate(rows)}
            segs = timeline_layout(rows, cached_probes(rows))
            dropped = sum(1 for r in rows if not place_kind(r))
            if not segs:
                canvas.create_text(20, 40, anchor='w', fill='#ccc',
                                   text='Nothing to place yet — match or assign videos first')
                return
            durs = [s['dur'] if s['dur'] is not None else 10.0 for s in segs]
            total = sum(durs)
            px = min(40.0, max(3.0, (canvas.winfo_width() - 40) / max(total, 1)))
            y0, y1 = 42, 158
            # ruler: ticks at a step that keeps labels ~70px apart
            step = next((s for s in (1, 2, 5, 10, 15, 30, 60, 120, 300, 600)
                         if s * px >= 70), 600)
            for sec in range(0, int(total) + 1, step):
                x = 16 + sec * px
                canvas.create_line(x, 26, x, y0 - 4, fill='#555')
                canvas.create_text(x + 2, 18, anchor='w', text=fmt_tc(sec),
                                   fill='#888', font=('TkDefaultFont', 8))
            x = 16.0
            for i, s in enumerate(segs):
                r, w, tag = s['row'], durs[i] * px, f'seg{i}'
                if s['kind'] == 'gap':
                    canvas.create_rectangle(x, y0, x + w, y1, fill='#2b2b2b',
                                            outline='#666', dash=(3, 2), tags=tag)
                    if w > 30:
                        canvas.create_text(x + w / 2, (y0 + y1) / 2, text='GAP',
                                           fill='#999', tags=tag)
                else:
                    canvas.create_rectangle(x, y0, x + w, y1, fill='#2f5d8f',
                                            outline='#7aa7d6', tags=tag)
                    img = thumbs.get(id(r))
                    if img and w > img.width() + 6:
                        canvas.create_image(x + 3, y0 + 3, image=img, anchor='nw', tags=tag)
                    maxch = int(w / 6.5)
                    if maxch >= 4:
                        canvas.create_text(x + 4, y1 - 28, anchor='nw', fill='white',
                                           text=f'{i + 1:02d} {r.label or r.game}'[:maxch],
                                           font=('TkDefaultFont', 9), tags=tag)
                        detail = ('whole file' if s['dur'] is None else
                                  f'{fmt_tc(s["start"])}–{fmt_tc(s["end"])}')
                        canvas.create_text(x + 4, y1 - 14, anchor='nw', fill='#bcd3e8',
                                           text=f'{detail} · {r.src.name}'[:maxch],
                                           font=('TkDefaultFont', 8), tags=tag)
                pos = rowpos.get(id(r))
                if pos is not None:
                    def pick(_e, p=pos):
                        if tree.exists(str(p)):
                            tree.selection_set(str(p))
                            tree.see(str(p))
                    canvas.tag_bind(tag, '<Button-1>', pick)
                if s['kind'] == 'clip':
                    canvas.tag_bind(tag, '<Double-1>',
                                    lambda _e, row=r: watch_rows([row]))
                x += w
            canvas.configure(scrollregion=(0, 0, x + 16, 200))
            nclips = sum(1 for s in segs if s['kind'] == 'clip')
            info.configure(text=f'{fmt_tc(total)} total · {nclips} clip(s) · '
                                f'{len(segs) - nclips} gap(s)'
                           + (f' · {dropped} row(s) not placeable (no video or timecode)'
                              if dropped else '')
                           + ' — click a block to select its row, double-click to watch it')

        def add_thumb(row_id, png):
            if not alive['ok']:
                return
            try:
                # keyed by row identity, not position: table edits while thumbnails
                # stream in must not attach a frame to the wrong clip
                thumbs[row_id] = tk.PhotoImage(file=str(png))   # main thread only
            except tk.TclError:
                return
            draw()

        def worker():
            try:
                ffmpeg, ffprobe = find_ffmpeg()
            except Exception as e:
                log(f'Preview: {e}')
                return
            rows = state['rows']          # snapshot; thumbs keyed by identity stay right
            # durations for whole-file rows first, so block widths become real
            for r in rows:
                if not alive['ok']:
                    return
                if r.src and r.whole_file and str(r.src) not in _probe_cache:
                    try:
                        probe(ffprobe, r.src)
                    except Exception:
                        continue
                    root.after(0, draw)
            tdir = Path(tempfile.mkdtemp(prefix='chop_preview_'))
            alive['tdir'] = tdir
            if not alive['ok']:           # closed while we were probing
                shutil.rmtree(tdir, ignore_errors=True)
                return
            for i, s in enumerate(timeline_layout(rows, cached_probes(rows))):
                if not alive['ok']:
                    return
                if s['kind'] != 'clip':
                    continue
                sec = s['start'] + (s['dur'] or 0) / 2    # mid-clip: the play, not the whistle
                png = tdir / f'{i}.png'
                try:
                    grab_frame(ffmpeg, s['row'].src, sec, png, width=96)
                except Exception:
                    continue
                root.after(0, add_thumb, id(s['row']), png)

        job = {'id': None}

        def on_configure(_e):
            if not alive['ok']:
                return
            if job['id']:
                win.after_cancel(job['id'])
            job['id'] = win.after(120, draw)

        def on_close():
            alive['ok'] = False
            win.destroy()
            if alive['tdir']:
                shutil.rmtree(alive['tdir'], ignore_errors=True)
        win.protocol('WM_DELETE_WINDOW', on_close)
        win.bind('<Configure>', on_configure)
        win.bind('<Escape>', lambda e: on_close())
        draw()
        threading.Thread(target=worker, daemon=True).start()

    def make_mogrts():
        """Stamp one pre-texted caption .mogrt per clip from the loaded sheet.

        Premiere 2026 removed the API for setting mogrt text from a script, so
        the text is baked into each file instead; the Clip Chopper Labels panel
        in Premiere then places 'NN - CAPTION.mogrt' onto every clip 'NN - …'.
        """
        rows = [r for r in state['rows'] if place_kind(r)]
        if not rows or not state['sheet']:
            log('Load a spreadsheet first — the captions come from its rows')
            return
        try:
            from make_caption_mogrts import load_template, stamp
        except ImportError:
            log('make_caption_mogrts.py is missing next to chopper.py')
            return
        p = filedialog.askopenfilename(
            title='Your template .mogrt (Graphics and Titles > Export As Motion '
                  'Graphics Template, once)',
            initialdir=str(Path(cfg.get('mogrt_template', '')).parent) if cfg.get('mogrt_template') else '.',
            filetypes=[('Motion Graphics Template', '*.mogrt'), ('All files', '*.*')])
        if not p:
            return
        cfg['mogrt_template'] = p
        save_settings(cfg)
        try:
            tpl = load_template(p)
        except (SystemExit, Exception) as e:
            messagebox.showerror('Template mogrt', str(e))
            return
        out = state['sheet'].parent / f'{sanitize_filename(state["sheet"].stem)} mogrts'
        out.mkdir(exist_ok=True)
        for stale in out.glob('*.mogrt'):
            stale.unlink()          # old copies would shadow renamed ones in the panel
        made = 0
        # numbering mirrors the timeline: gaps keep their slot so 'NN' always
        # matches the clip names Premiere shows after importing the XML.
        # The team is ALWAYS appended — a caption without the opponent is the
        # one thing a "GOAL VS TEAM..." template can never want.
        for n, seg in enumerate(timeline_layout(rows, cached_probes(rows)), 1):
            if seg['kind'] != 'clip':
                continue
            text = label_text(seg['row'], True).upper()
            if not text:
                continue
            try:
                stamp(*tpl, out / f'{n:02d} - {sanitize_filename(text)}.mogrt', n, text)
                made += 1
            except Exception as e:
                log(f'mogrt {n:02d} failed: {e}')
        log(f'{made} caption mogrt(s) -> {out.name}/ — in Premiere: '
            'Window > Extensions (Legacy) > Clip Chopper Labels, pick that folder')
        open_media(out)

    def preview_label():
        """The label editor: text, size, color, position — live on a real frame."""
        team_on = team_var.get()
        sel = [r for r in selected_rows() if r.src]
        cand = (sel or [r for r in state['rows'] if r.src and label_text(r, team_on)]
                or [r for r in state['rows'] if r.src])
        if not cand:
            log('Load your spreadsheet and videos first — no clip to grab a frame from')
            return
        r = cand[0]
        win = tk.Toplevel(root)
        win.title('Label preview')
        alive = {'ok': True}
        holder = {'photo': None, 'frame': None, 'n': 0}   # photo ref must outlive configure()
        img_lbl = ttk.Label(win, text='Grabbing a frame…', anchor='center')
        img_lbl.pack(fill='both', expand=True, padx=8, pady=(8, 0))
        bar = ttk.Frame(win, padding=8)
        bar.pack(side='bottom', fill='x')

        def render_pass(n):
            """Off the UI thread: composite label onto the cached frame, then hand back."""
            if holder['n'] != n or not holder['frame'] or not alive['ok']:
                return
            frame_png, w, h, tdir = holder['frame']
            try:
                from PIL import Image
                text = label_text(r, team_var.get())
                try:
                    size = max(8, int(float(size_var.get())))
                except ValueError:
                    size = 48
                base = Image.open(frame_png).convert('RGBA')
                if text:
                    lbl_png = tdir / 'label.png'
                    font = Path(font_var.get()) if font_var.get() else None
                    render_label_png(text, font, size, w, h, lbl_png,
                                     color=color_var.get(), pos=pos_var.get())
                    base = Image.alpha_composite(base, Image.open(lbl_png).convert('RGBA'))
                scale = min(880 / base.width, 500 / base.height, 1.0)
                view = tdir / f'view{n}.png'
                base.resize((round(base.width * scale), round(base.height * scale))).save(view)

                def show():
                    if holder['n'] != n or not alive['ok']:
                        return
                    try:
                        holder['photo'] = tk.PhotoImage(file=str(view))
                        img_lbl.configure(image=holder['photo'], text='')
                        win.title(f'Label preview — "{text}" at size {size}'
                                  if text else 'Label preview — (this row has no label)')
                    except tk.TclError:
                        pass
                root.after(0, show)
            except Exception as e:
                # bind the message NOW: `e` is unbound once the except block exits,
                # and the lambda runs later on the Tk main loop
                msg = f'Could not render preview: {e}'
                root.after(0, lambda m=msg: alive['ok'] and img_lbl.configure(text=m, image=''))

        def schedule(*_):
            try:
                cfg.update({'size': max(8, int(float(size_var.get())))})
            except ValueError:
                pass
            cfg.update({'label_team': team_var.get(), 'font': font_var.get(),
                        'label_color': color_var.get(), 'label_pos': pos_var.get(),
                        'label_style': style_var.get()})
            save_settings(cfg)      # every tweak sticks — Generate uses exactly this
            holder['n'] += 1
            n = holder['n']
            win.after(250, lambda: holder['n'] == n and threading.Thread(
                target=render_pass, args=(n,), daemon=True).start())

        def touched(*_):
            style_var.set('Custom')     # a hand tweak means it's your style now
            schedule()

        # row 0: the words — edits this row's Label directly, same as the table cell
        ttk.Label(bar, text='Label:').grid(row=0, column=0, sticky='w')
        text_var = tk.StringVar(value=r.label)

        def on_text(*_):
            r.label = text_var.get().strip()
            schedule()
        text_entry = ttk.Entry(bar, textvariable=text_var)
        text_entry.grid(row=0, column=1, columnspan=8, sticky='ew', padx=(4, 0))
        text_entry.bind('<KeyRelease>', on_text)

        # row 1: how it looks
        style_var = tk.StringVar(value=cfg.get('label_style', 'Custom'))
        swatch = tk.Label(bar, width=2, background=color_var.get(), relief='sunken')

        def apply_style(_e=None):
            if style_var.get() != 'Darsh template':
                schedule()
                return
            size_var.set(str(DARSH_STYLE['size']))
            color_var.set(DARSH_STYLE['color'])
            swatch.configure(background=DARSH_STYLE['color'])
            pos_var.set(DARSH_STYLE['pos'])
            f = style_font()
            if f:
                font_var.set(f)
                font_lbl.configure(text=Path(f).name)
            else:
                log('Futura not found on this machine — template style keeps your current font')
            schedule()
        ttk.Label(bar, text='Style:').grid(row=1, column=0, sticky='w', pady=(6, 0))
        style_box = ttk.Combobox(bar, textvariable=style_var, state='readonly', width=13,
                                 values=['Darsh template', 'Custom'])
        style_box.grid(row=1, column=1, sticky='w', padx=(4, 0), pady=(6, 0))
        style_box.bind('<<ComboboxSelected>>', apply_style)
        ttk.Label(bar, text='Size:').grid(row=1, column=2, sticky='w',
                                          padx=(10, 0), pady=(6, 0))
        spin = ttk.Spinbox(bar, from_=8, to=300, increment=4, textvariable=size_var,
                           width=5, command=touched)
        spin.grid(row=1, column=3, sticky='w', padx=(4, 0), pady=(6, 0))
        spin.bind('<KeyRelease>', touched)

        def pick_color():
            c = colorchooser.askcolor(color=color_var.get(), parent=win,
                                      title='Label color')
            if c and c[1]:
                color_var.set(c[1])
                swatch.configure(background=c[1])
                touched()
        ttk.Button(bar, text='Color…', command=pick_color).grid(row=1, column=4,
                                                                padx=(10, 2), pady=(6, 0))
        swatch.grid(row=1, column=5, pady=(6, 0))
        pos_box = ttk.Combobox(bar, textvariable=pos_var, state='readonly', width=13,
                               values=list(LABEL_POSITIONS))
        pos_box.grid(row=1, column=6, padx=(10, 0), pady=(6, 0))
        pos_box.bind('<<ComboboxSelected>>', touched)
        ttk.Checkbutton(bar, text='+ team', variable=team_var,
                        command=schedule).grid(row=1, column=7, padx=(10, 0), pady=(6, 0))
        ttk.Button(bar, text='Font…',
                   command=lambda: (pick_font(), touched())).grid(row=1, column=8,
                                                                 sticky='w', padx=(10, 0),
                                                                 pady=(6, 0))
        bar.columnconfigure(8, weight=1)
        ttk.Label(bar, foreground='#666',
                  text='click the frame to put the text there · arrow keys nudge '
                       '(shift = bigger steps)').grid(row=2, column=0, columnspan=9,
                                                      sticky='w', pady=(4, 0))

        # fine positioning: click the frame to place, arrows to nudge
        def on_place(e):
            ph = holder['photo']
            if not ph:
                return
            iw, ih = ph.width(), ph.height()
            fx = min(max((e.x - (img_lbl.winfo_width() - iw) / 2) / iw, 0), 1)
            fy = min(max((e.y - (img_lbl.winfo_height() - ih) / 2) / ih, 0), 1)
            pos_var.set(f'{fx * 100:.1f}%,{fy * 100:.1f}%')
            touched()
        img_lbl.bind('<Button-1>', on_place)

        NAMED_PCT = {'left': 5.0, 'center': 45.0, 'right': 80.0,
                     'top': 10.0, 'middle': 50.0, 'bottom': 90.0}

        def nudge(dx, dy):
            if m := POS_PCT_RE.fullmatch(pos_var.get()):
                fx, fy = float(m.group(1)), float(m.group(2))
            else:                       # first nudge converts a corner to a free point
                vert, _, horiz = pos_var.get().partition(' ')
                fx, fy = NAMED_PCT.get(horiz, 5.0), NAMED_PCT.get(vert, 90.0)
            pos_var.set(f'{min(max(fx + dx, 0), 100):.1f}%,'
                        f'{min(max(fy + dy, 0), 100):.1f}%')
            touched()

        def on_arrow(e):
            w = win.focus_get()         # let the text fields keep their arrow keys
            if w is not None and w.winfo_class() in ('TEntry', 'TSpinbox', 'TCombobox',
                                                     'Entry', 'Spinbox'):
                return
            step = 2.0 if e.state & 1 else 0.5
            dx, dy = {'Left': (-step, 0), 'Right': (step, 0),
                      'Up': (0, -step), 'Down': (0, step)}[e.keysym]
            nudge(dx, dy)
            return 'break'
        for k in ('<Left>', '<Right>', '<Up>', '<Down>'):
            win.bind(k, on_arrow)

        def first_grab():
            try:
                ffmpeg, ffprobe = find_ffmpeg()
                p = probe(ffprobe, r.src)
                ensure_pip('PIL', 'pillow')
                if r.start is not None and r.end is not None:
                    sec = (r.start + r.end) / 2           # mid-clip: the play, not the whistle
                else:
                    sec = min(60.0, (p['duration'] or 0) / 2)
                tdir = Path(tempfile.mkdtemp(prefix='chop_label_'))
                holder['tdir'] = tdir
                frame_png = tdir / 'frame.png'
                grab_frame(ffmpeg, r.src, sec, frame_png)
                holder['frame'] = (frame_png, p['width'], p['height'], tdir)
                holder['n'] += 1
                render_pass(holder['n'])
            except Exception as e:
                msg = f'Could not grab a frame from {r.src.name}: {e}'   # e dies with the block
                root.after(0, lambda m=msg: alive['ok'] and img_lbl.configure(text=m))

        def on_close():
            alive['ok'] = False
            win.destroy()
            if holder.get('tdir'):
                shutil.rmtree(holder['tdir'], ignore_errors=True)
            refresh_table()     # label edits made here show up in the table
        win.protocol('WM_DELETE_WINDOW', on_close)
        win.bind('<Escape>', lambda e: on_close())
        threading.Thread(target=first_grab, daemon=True).start()

    tree.bind('<Double-1>', on_double_click)
    tree.bind('<Button-1>', on_toggle_click, add='+')
    tree.bind('<space>', lambda e: toggle_rows(selected_rows()))
    tree.bind('<Button-3>', on_right_click)     # Windows/Linux right-click
    tree.bind('<Button-2>', on_right_click)     # macOS right-click / two-finger tap
    if has_dnd:
        def on_drop(event):
            try:
                paths = [Path(p) for p in root.tk.splitlist(event.data)]
                folders = [p for p in paths if p.is_dir()]
                sheets = [p for p in paths if p.is_file()]
                if folders:
                    add_videos(folders)
                if sheets:
                    load_sheet(sheets[0])
                if not folders and not sheets:
                    log(f'Could not read what was dropped: {event.data!r}')
            except Exception as e:      # a raised handler silently kills further drops
                log(f'Drop failed: {e}')

        def register(w, depth=0):
            """Every widget must opt in — a drop onto an unregistered one does nothing."""
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', on_drop)
            except Exception:
                pass                     # some ttk widgets refuse; the rest still work
            for child in w.winfo_children():
                register(child, depth + 1)
        register(root)

    drain_log()
    log('1) Load your spreadsheet   2) Add your video folder(s) — subfolders are searched too'
        '   3) Review the table   4) Generate')
    log('Click the ☑ column (or select rows and press space) to toggle a clip off — '
        'its slot stays as a gap in the timeline.')
    if has_dnd:
        log('Drag and drop is ON — drop your sheet or video folders anywhere in this window.')
    else:
        log('Drag and drop is OFF on this machine — use the Browse… / Add folder… buttons '
            'instead (everything works the same).')
        log(f'   reason: {dnd_error}')
        log('   to fix: quit, run "python3 -m pip install tkinterdnd2" in Terminal, '
            'and start the app with Chop.command (Mac) or Chop.bat (Windows).')
    root.mainloop()


if __name__ == '__main__':
    run_gui()
