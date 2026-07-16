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
import json
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
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


def parse_tc(text):
    """'14:18' -> 858.0, '1:00:10' -> 3610.0, else None."""
    m = TC_RE.match(text.strip())
    if not m:
        return None
    a, b, c = m.groups()
    if c is None:
        return int(a) * 60 + int(b)
    return int(a) * 3600 + int(b) * 60 + int(c)


def parse_range(text):
    """Return (start, end) seconds or None if text is not a timecode range."""
    m = RANGE_RE.match(DASHES.sub('-', text))
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

def norm_tokens(name):
    """'Georgetown Prep #2 (Daksh)' -> ['georgetown', 'prep', '2']."""
    name = re.sub(r'\([^)]*\)', ' ', name.lower())
    return re.sub(r'[^a-z0-9]+', ' ', name).split()


def list_videos(folder, exclude=None):
    folder = Path(folder)
    return sorted(p for p in folder.rglob('*')
                  if p.suffix.lower() in VIDEO_EXTS and p.is_file()
                  # skip hidden files/dirs — esp. macOS "._*" AppleDouble junk on USB drives
                  and not any(part.startswith('.') for part in p.relative_to(folder).parts)
                  and not (exclude and exclude in p.parents))


def match_score(game, filename):
    g, f = norm_tokens(game), norm_tokens(filename)
    if not g or not f:
        return 0.0
    overlap = len(set(g) & set(f)) / len(g)
    ratio = difflib.SequenceMatcher(None, ' '.join(g), ' '.join(f)).ratio()
    return 0.6 * overlap + 0.4 * ratio


def match_videos(rows, folder, threshold=0.55, exclude=None):
    """Fill row.src by fuzzy-matching game names to files. One match per game name."""
    videos = list_videos(folder, exclude=exclude)
    cache = {}
    for r in rows:
        if r.manual and r.src:
            continue
        if not r.game:
            r.flags.append('no game name')
            continue
        if r.game not in cache:
            # Digits are identity: "Prep #2" must never match "Prep 1.mp4" just because
            # the words agree — require every digit token of the game in the filename.
            digits = {t for t in norm_tokens(r.game) if t.isdigit()}
            pool = [v for v in videos if digits <= set(norm_tokens(v.stem))]
            scored = sorted(((match_score(r.game, v.stem), v) for v in pool),
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


def probe(ffprobe, path):
    """Return {'fps', 'width', 'height', 'duration', 'has_audio'} for a video."""
    path = str(path)
    if path in _probe_cache:
        return _probe_cache[path]
    out = subprocess.run(
        [ffprobe, '-v', 'error', '-show_entries',
         'stream=codec_type,width,height,r_frame_rate,duration'
         ':stream_disposition=attached_pic:format=duration',
         '-of', 'json', path],
        capture_output=True, text=True, check=True).stdout
    data = json.loads(out)
    info = {'fps': 30.0, 'width': 1920, 'height': 1080, 'duration': 0.0, 'has_audio': False}
    if fd := data.get('format', {}).get('duration'):
        info['duration'] = float(fd)
    got_video = False
    for s in data.get('streams', []):
        if s.get('codec_type') == 'video' and not got_video:
            if s.get('disposition', {}).get('attached_pic'):
                continue  # embedded cover art, not the footage
            if 'width' in s:
                info['width'], info['height'] = s['width'], s['height']
            num, _, den = s.get('r_frame_rate', '').partition('/')
            try:
                fps = float(num) / float(den or 1)
            except (ValueError, ZeroDivisionError):
                fps = 0.0
            if 1.0 <= fps <= 240.0:
                info['fps'] = fps
            if not info['duration'] and s.get('duration'):
                info['duration'] = float(s['duration'])
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


# ------------------------------------------------------------ label graphics

def render_label_png(text, font_path, size, width, height, out_path):
    """Full-frame transparent PNG with the label bottom-left — a Premiere overlay still.

    `size` means pixels at 1080p; scaled proportionally for other sequence heights.
    """
    from PIL import Image, ImageDraw, ImageFont
    px = max(8, round(size * height / 1080))
    try:
        font = ImageFont.truetype(str(font_path), px) if font_path else ImageFont.load_default(px)
    except Exception:
        font = ImageFont.load_default(px)
    img = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    margin = round(height * 0.05)
    pad = max(6, px // 3)
    l, t, r, b = d.textbbox((0, 0), text, font=font)
    tw, th = r - l, b - t
    box = (margin, height - margin - th - 2 * pad, margin + tw + 2 * pad, height - margin)
    d.rounded_rectangle(box, radius=max(4, pad // 2), fill=(0, 0, 0, 140))
    d.text((margin + pad - l, height - margin - pad - th - t), text,
           font=font, fill=(255, 255, 255, 255))
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
    label_pngs: optional {row_index: png_path} — full-frame stills laid on a second
    video track above each clip (imports far more reliably than XML text generators).
    """
    fps_list = [probes[str(r.src)]['fps'] for r in rows]
    seq_fps = max(set(fps_list), key=fps_list.count)
    first = probes[str(rows[0].src)]

    v_items, a_items, o_items, markers = [], [], [], []
    file_defs = {}   # path -> file id (emit full <file> once, then reference)
    png_defs = {}
    t = 0            # sequence playhead in frames
    for n, r in enumerate(rows, 1):
        p = probes[str(r.src)]
        start, end = (0.0, p['duration']) if r.whole_file else (r.start, r.end)
        f_in, f_out = round(start * p['fps']), round(end * p['fps'])
        dur = round((end - start) * seq_fps)
        file_frames = round(p['duration'] * p['fps'])
        name = escape(f'{n:02d} - {r.label or r.game}')
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

        def links(vid, aid):
            return (f'<link><linkclipref>{vid}</linkclipref><mediatype>video</mediatype>'
                    f'<trackindex>1</trackindex><clipindex>{n}</clipindex></link>'
                    f'<link><linkclipref>{aid}</linkclipref><mediatype>audio</mediatype>'
                    f'<trackindex>1</trackindex><clipindex>{n}</clipindex></link>')

        vid, aid = f'clipitem-v{n}', f'clipitem-a{n}'
        common = (f'<enabled>TRUE</enabled><duration>{file_frames}</duration>'
                  f'{_rate_xml(p["fps"])}<start>{t}</start><end>{t + dur}</end>'
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

    labels_cfg: optional {'font': Path|None, 'size': int} — render each row's label
    as a still-image overlay on a second timeline track.
    """
    ffmpeg, ffprobe = find_ffmpeg()
    rows = [r for r in rows if r.src is not None and (r.whole_file or (
        r.start is not None and r.end is not None and r.end > r.start))]
    if not rows:
        raise ValueError('No usable rows — every row is missing a video file or timecode')

    log(f'Probing {len({str(r.src) for r in rows})} video file(s)...')
    probes, unreadable = {}, set()
    for path in {str(r.src) for r in rows}:
        try:
            probes[path] = probe(ffprobe, path)
        except Exception:
            unreadable.add(path)
            log(f'WARNING: could not read {Path(path).name} — skipping its clips')
    rows = [r for r in rows if str(r.src) not in unreadable]
    if not rows:
        raise ValueError('None of the matched video files could be read')
    kept = []
    for r in rows:  # enforce real durations now that we know them
        d = probes[str(r.src)]['duration']
        if r.whole_file or not d:
            kept.append(r)
            continue
        if r.start >= d:
            log(f'SKIPPED "{r.label or r.game}": starts at {fmt_tc(r.start)} but '
                f'{r.src.name} is only {fmt_tc(d)} long — wrong video or wrong time?')
            continue
        if r.end > d:
            log(f'NOTE: "{r.label or r.game}" trimmed to the end of {r.src.name}')
            r.end = d
        kept.append(r)
    rows = kept
    if not rows:
        raise ValueError('No usable rows left after checking video durations')

    label_pngs = {}
    if labels_cfg:
        try:
            ensure_pip('PIL', 'pillow')
        except ImportError:
            log('WARNING: Pillow missing and could not auto-install — labels skipped. '
                'Run: python3 -m pip install pillow')
            labels_cfg = None
    if labels_cfg:
        first = probes[str(rows[0].src)]
        label_dir = out_dir / 'labels'
        label_dir.mkdir(exist_ok=True)
        by_text = {}
        for i, r in enumerate(rows):
            if not r.label:
                continue
            if r.label not in by_text:
                png = label_dir / f'{sanitize_filename(r.label)}.png'
                try:
                    render_label_png(r.label, labels_cfg.get('font'), labels_cfg.get('size', 48),
                                     first['width'], first['height'], png)
                    by_text[r.label] = png
                except Exception as e:
                    by_text[r.label] = None
                    log(f'WARNING: could not render label "{r.label}": {e}')
            if by_text[r.label]:
                label_pngs[i] = by_text[r.label]
        if label_pngs:
            log(f'{sum(1 for v in by_text.values() if v)} label graphic(s) written to labels/')

    xml_path = out_dir / f'{sanitize_filename(sequence_name)}_timeline.xml'
    xml_path.write_text(build_xmeml(rows, probes, sequence_name, label_pngs), encoding='utf-8')
    log(f'Timeline written: {xml_path.name}  (Premiere: File > Import)')

    ok, failed = 0, []
    if export_clips:
        clip_dir = out_dir / 'clips'
        clip_dir.mkdir(exist_ok=True)
        for n, r in enumerate(rows, 1):
            stem = sanitize_filename(f'{n:02d} - {r.label or "clip"} - {norm_tokens(r.game) and " ".join(norm_tokens(r.game)) or r.src.stem}')
            out_path = clip_dir / f'{stem}{r.src.suffix.lower() if r.whole_file else ".mp4"}'
            try:
                cut_clip(ffmpeg, r, out_path)
                if not out_path.exists() or out_path.stat().st_size < 1000:
                    raise RuntimeError('ffmpeg produced an empty file')
                ok += 1
                log(f'[{n}/{len(rows)}] {out_path.name}')
            except Exception as e:
                failed.append(r)
                err = str(getattr(e, 'stderr', '') or e).strip()
                log(f'[{n}/{len(rows)}] FAILED {out_path.name}: {err[:200]}')
    return xml_path, ok, failed


def estimate_clip_bytes(rows):
    """Rough output size at ~10 Mbps for cut clips, actual size for whole-file copies."""
    total = 0
    for r in rows:
        if r.whole_file and r.src:
            total += r.src.stat().st_size
        elif r.start is not None and r.end is not None and r.end > r.start:
            total += int((r.end - r.start) * 1.25e6)
    return total


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
    from tkinter import ttk, filedialog, messagebox
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        has_dnd = True
    except Exception:
        root = tk.Tk()
        has_dnd = False

    root.title('Clip Chopper')
    root.geometry('1100x640')
    state = {'rows': [], 'sheet': None, 'videos': None}
    log_q = queue.Queue()

    # --- top: spreadsheet + videos folder pickers -------------------------
    top = ttk.Frame(root, padding=8)
    top.pack(fill='x')
    drop_text = 'Drop spreadsheet here (or Browse…)' if has_dnd else 'Choose spreadsheet (Browse…)'
    sheet_lbl = ttk.Label(top, text=drop_text, relief='groove', anchor='center', padding=10)
    sheet_lbl.grid(row=0, column=0, columnspan=2, sticky='ew', padx=(0, 6))
    ttk.Button(top, text='Browse…', command=lambda: pick_sheet()).grid(row=0, column=2)
    ttk.Label(top, text='or Google Sheets URL:').grid(row=1, column=0, sticky='w', pady=(6, 0))
    url_var = tk.StringVar()
    url_entry = ttk.Entry(top, textvariable=url_var, width=60)
    url_entry.grid(row=1, column=1, sticky='ew', pady=(6, 0), padx=(0, 6))
    ttk.Button(top, text='Load URL', command=lambda: load_url()).grid(row=1, column=2, pady=(6, 0))
    vid_lbl = ttk.Label(top, text='Videos folder: (not set)', anchor='w')
    vid_lbl.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(6, 0))
    ttk.Button(top, text='Choose folder…', command=lambda: pick_videos()).grid(row=2, column=2, pady=(6, 0))
    top.columnconfigure(1, weight=1)

    # --- middle: review table --------------------------------------------
    mid = ttk.Frame(root, padding=(8, 0))
    mid.pack(fill='both', expand=True)
    cols = ('order', 'game', 'file', 'in', 'out', 'label', 'notes', 'status')
    tree = ttk.Treeview(mid, columns=cols, show='headings', selectmode='browse')
    widths = {'order': 45, 'game': 180, 'file': 180, 'in': 70, 'out': 70,
              'label': 170, 'notes': 180, 'status': 220}
    for c in cols:
        tree.heading(c, text=c.title())
        tree.column(c, width=widths[c], anchor='w')
    tree.tag_configure('bad', background='#ffd6d6')
    ys = ttk.Scrollbar(mid, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=ys.set)
    tree.pack(side='left', fill='both', expand=True)
    ys.pack(side='right', fill='y')

    # --- bottom: options + generate + log ---------------------------------
    cfg = load_settings()
    bot = ttk.Frame(root, padding=8)
    bot.pack(fill='x')

    # label-overlay options
    labels_var = tk.BooleanVar(value=cfg.get('labels_on', False))
    font_var = tk.StringVar(value=cfg.get('font', ''))
    size_var = tk.StringVar(value=str(cfg.get('size', 48)))
    ttk.Checkbutton(bot, text='Add clip labels as text in the timeline',
                    variable=labels_var).grid(row=0, column=0, sticky='w')
    font_frame = ttk.Frame(bot)
    font_frame.grid(row=0, column=1, columnspan=2, sticky='w', padx=(10, 0))
    font_lbl = ttk.Label(font_frame, width=26, anchor='w',
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

    export_var = tk.BooleanVar(value=cfg.get('export_clips', True))
    ttk.Style().configure('Big.TCheckbutton', font=('TkDefaultFont', 10, 'bold'))
    export_chk = ttk.Checkbutton(bot, text='Make individual clip files',
                                 variable=export_var, style='Big.TCheckbutton')
    export_chk.grid(row=1, column=0, sticky='w', pady=(6, 0))
    gen_btn = ttk.Button(bot, text='Generate', state='disabled', command=lambda: start_generate())
    gen_btn.grid(row=1, column=1, padx=10, pady=(6, 0))
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
        tree.delete(*tree.get_children())
        for i, r in enumerate(state['rows']):
            vals = (r.order if r.order is not None else '',
                    r.game,
                    r.src.name if r.src else '',
                    'whole file' if r.whole_file else fmt_tc(r.start),
                    'whole file' if r.whole_file else fmt_tc(r.end),
                    r.label, r.notes,
                    '; '.join(r.flags) if r.flags else 'ok')
            tree.insert('', 'end', iid=str(i), values=vals,
                        tags=('bad',) if r.flags else ())
        n_bad = sum(1 for r in state['rows'] if r.flags)
        if state['rows']:
            est = estimate_clip_bytes(state['rows']) / 1e6
            export_chk.configure(text=f'Make individual clip files (~{est:.0f} MB)')
            log(f'{len(state["rows"])} clips parsed'
                + (f', {n_bad} need review (red rows — double-click to fix)' if n_bad else ', all ok'))
        gen_btn.configure(state='normal' if state['rows'] and state['videos'] else 'disabled')

    def rematch():
        if state['rows'] and state['videos']:
            for r in state['rows']:
                if r.manual and r.src:
                    continue
                r.src = None
                r.flags = [f for f in r.flags if 'match' not in f and 'video' not in f
                           and 'game name' not in f]
            exclude = state['sheet'].parent / 'clips' if state['sheet'] else None
            match_videos(state['rows'], state['videos'], exclude=exclude)
        refresh_table()

    def load_sheet(path):
        try:
            state['rows'] = parse_sheet(path)
            state['sheet'] = Path(path)
            sheet_lbl.configure(text=Path(path).name)
            rematch()
        except Exception as e:
            messagebox.showerror('Could not read spreadsheet', str(e))

    def pick_sheet():
        p = filedialog.askopenfilename(filetypes=[('Spreadsheets', '*.xlsx *.csv'), ('All files', '*.*')])
        if p:
            load_sheet(p)

    def load_url():
        try:
            load_sheet(fetch_google_sheet(url_var.get()))
        except Exception as e:
            messagebox.showerror('Google Sheets', str(e))

    def pick_videos():
        p = filedialog.askdirectory(title='Folder containing the game videos')
        if p:
            state['videos'] = Path(p)
            vid_lbl.configure(text=f'Videos folder: {p}  ({len(list_videos(p))} videos found)')
            rematch()

    def on_double_click(event):
        item, col = tree.identify_row(event.y), tree.identify_column(event.x)
        if not item:
            return
        r = state['rows'][int(item)]
        col_name = cols[int(col[1:]) - 1]
        if col_name == 'file':
            p = filedialog.askopenfilename(
                title=f'Video for "{r.game}"',
                initialdir=state['videos'] or '.',
                filetypes=[('Videos', ' '.join(f'*{e}' for e in VIDEO_EXTS)), ('All files', '*.*')])
            if p:
                # same game -> same file everywhere; blank game names stay individual
                targets = [o for o in state['rows'] if o.game == r.game] if r.game else [r]
                for other in targets:
                    other.src = Path(p)
                    other.manual = True
                    other.flags = [f for f in other.flags if 'match' not in f and 'video' not in f]
                refresh_table()
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
        usable = [r for r in state['rows'] if r.src and (r.whole_file or
                  (r.start is not None and r.end is not None and r.end > r.start))]
        skipped = len(state['rows']) - len(usable)
        if skipped and not messagebox.askyesno(
                'Some rows will be skipped',
                f'{skipped} row(s) still have problems and will be left out.\nContinue with {len(usable)}?'):
            return
        out_dir = state['sheet'].parent
        tmp = Path(tempfile.gettempdir())
        if out_dir == tmp or tmp in out_dir.parents:  # e.g. Google Sheets download
            p = filedialog.askdirectory(title='Where should the timeline and clips be saved?')
            if not p:
                return
            out_dir = Path(p)
        gen_btn.configure(state='disabled')
        prog.configure(maximum=max(len(usable), 1), value=0)
        work_rows = copy.deepcopy(usable)  # worker gets its own rows; table edits can't race it
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
            labels_cfg = {'font': lfont, 'size': lsize}
        save_settings({'labels_on': labels_var.get(), 'font': font_var.get(),
                       'size': lsize, 'export_clips': export})

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

    tree.bind('<Double-1>', on_double_click)
    if has_dnd:
        def on_drop(event):
            paths = root.tk.splitlist(event.data)
            if paths:
                load_sheet(paths[0])
        for w in (root, sheet_lbl):
            w.drop_target_register(DND_FILES)
            w.dnd_bind('<<Drop>>', on_drop)

    drain_log()
    log('1) Load your spreadsheet   2) Choose the videos folder   3) Review the table   4) Generate')
    root.mainloop()


if __name__ == '__main__':
    run_gui()
