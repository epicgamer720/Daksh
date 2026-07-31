/* add_labels.jsx v7 — place pre-texted caption mogrts, no questions asked.
   Clears the top video track (previous caption attempts) without a confirm
   dialog — hidden modals were eating the run. */
(function () {
    var seq = app.project.activeSequence;
    if (!seq) { alert('Open the sequence first.'); return 'no sequence'; }
    var vTracks = seq.videoTracks;
    var target = vTracks.numTracks - 1;
    var clips = vTracks[0].clips;
    if (!clips.numItems) { alert('No clips on V1.'); return 'no clips'; }
    var top = vTracks[target];
    var cleared = 0;
    for (var r = top.clips.numItems - 1; r >= 0; r--) {
        try { top.clips[r].remove(false, false); cleared++; } catch (eR) {}
    }
    var dir = Folder.selectDialog('Pick the folder of caption .mogrts');
    if (!dir) { return 'cancelled'; }
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
    alert('Cleared ' + cleared + ' old graphics, placed ' + placed + ' captions.' +
          (skipped.length ? '\nNo mogrt for: ' + skipped.join(', ') : ''));
    return 'placed ' + placed;
})();
