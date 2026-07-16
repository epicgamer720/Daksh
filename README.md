# Daksh — Clip Chopper

Turns a highlight cut-list spreadsheet into a Premiere Pro timeline plus (optionally)
individually cut clip files. Your full-game videos are only ever **read** — never copied,
moved, or re-encoded.

## Run it

- **Windows**: double-click `Chop.bat`
- **macOS**: double-click `Chop.command` (or run `bash Chop.command`)

First launch installs everything it needs (Python packages + a bundled ffmpeg).
Requires Python 3 from [python.org](https://python.org).

## Use it

1. Drop your spreadsheet (.xlsx/.csv) onto the window — or paste a Google Sheets link.
2. Pick the folder that holds the game videos (filenames should mention the opponent).
3. Review the table. Red rows need attention — double-click a cell to fix a time or pick a file.
4. Optional: turn on **Add clip labels as text in the timeline** and pick your font file + size.
5. **Generate** → import the `..._timeline.xml` into Premiere (File > Import). Clips appear in
   order; labels sit on a second video track and stay editable. Keep the `labels/` folder —
   the timeline references it.

The spreadsheet just needs a column of timecode ranges like `52:10-52:17` (or `Clip in Folder`
for pre-cut clips), and ideally columns for game, order, label, and notes — the app finds them
by header name or by what the cells look like.

## Native Premiere text labels (optional)

The timeline's label overlays are images. If you want **editable Premiere text** instead:

1. One-time: in Premiere, make one text graphic styled how you want (your font/size),
   select it → Graphics and Titles → **Export As Motion Graphics Template…**, then delete it.
2. Generate/import the timeline **with the PNG-labels checkbox off** (or delete track V2).
3. File → Scripts → **Run Script…** → pick `add_labels.jsx` → choose your saved .mogrt.

Every clip gets a native text layer above it with its label, editable in Essential Graphics.

## Usage stats

To count how many unique spreadsheets are used, the app sends a one-time ping to the
developer when a sheet is loaded: the **spreadsheet's file name** (plus a short content id)
and **this computer's name**. The sheet's contents — rows, timecodes, everything inside —
never leave your machine. The app prints a log line whenever a ping is sent. Opt out any
time by setting the environment variable `CLIP_CHOPPER_NO_TRACK=1`.
