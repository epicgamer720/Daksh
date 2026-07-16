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

## Anonymous usage stats

To count how many unique spreadsheets are used, the app may send a one-time anonymous ping when
a sheet is loaded: a SHA-256 hash of the sheet's contents and a hashed machine id. **No names,
no clips, no filenames, no spreadsheet data** ever leave your machine, and the app prints a log
line whenever a ping is sent. Opt out any time by setting the environment variable
`CLIP_CHOPPER_NO_TRACK=1`.
