/* add_labels.jsx v6 — place PRE-TEXTED mogrts (one per caption, text baked in).
   Premiere 2026 removed getMGTComponent(), so no script can set mogrt text —
   instead each clip 'NN - …' gets the mogrt file 'NN - ….mogrt' from a folder. */
(function () {
    var seq = app.project.activeSequence;
    if (!seq) { alert('Open the sequence first.'); return; }
    var vTracks = seq.videoTracks;
    var target = vTracks.numTracks - 1;
    var clips = vTracks[0].clips;
    if (!clips.numItems) { alert('No clips on V1.'); return; }
    var top = vTracks[target];
    if (top.clips.numItems > 0) {
        if (!confirm('Top track V' + (target + 1) + ' holds ' + top.clips.numItems +
                     ' old graphics. Remove them and place the captions?')) return;
        for (var r = top.clips.numItems - 1; r >= 0; r--) {
            try { top.clips[r].remove(false, false); } catch (eR) {}
        }
    }
    var dir = Folder.selectDialog('Pick the folder of caption .mogrts');
    if (!dir) return;
    var mogrts = dir.getFiles('*.mogrt');
    function forSlot(slot) {
        for (var i = 0; i < mogrts.length; i++) {
            if (decodeURI(mogrts[i].name).indexOf(slot + ' -') === 0) return mogrts[i];
        }
        return null;
    }
    var placed = 0, skipped = [];
    for (var n = 0; n < clips.numItems; n++) {
        var m = String(clips[n].name).match(/^\s*(\d+)/);
        var mg = m ? forSlot(m[1]) : null;
        if (!mg) { skipped.push(clips[n].name); continue; }
        var item = seq.importMGT(mg.fsName, clips[n].start.ticks, target, 0);
        if (!item) { skipped.push(clips[n].name); continue; }
        placed++;
        try { var t = new Time(); t.ticks = clips[n].end.ticks; item.end = t; } catch (e) {}
    }
    alert('Placed ' + placed + ' captions (text baked into each mogrt).' +
          (skipped.length ? '\nNo mogrt for: ' + skipped.join(', ') : ''));
})();
