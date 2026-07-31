/* add_labels.jsx v8 — captionsRun(dirPath): clear top track, place pre-texted
   mogrts matched by clip number. The FOLDER is chosen by the CEP panel (its
   native dialog), because ExtendScript dialogs can fail to display from CEP. */
function captionsRun(dirPath) {
    var seq = app.project.activeSequence;
    if (!seq) { alert('Open the sequence first.'); return 'no sequence'; }
    var vTracks = seq.videoTracks;
    var target = vTracks.numTracks - 1;
    var clips = vTracks[0].clips;
    if (!clips.numItems) { alert('No clips on V1.'); return 'no clips'; }
    if (target < 1) {
        alert('This sequence has only one video track — the captions need an empty ' +
              'track ABOVE the clips. Right-click a track header > Add Track, then ' +
              'click again.');
        return 'need a second video track';
    }
    var top = vTracks[target];
    var cleared = 0;
    for (var r = top.clips.numItems - 1; r >= 0; r--) {
        try { top.clips[r].remove(false, false); cleared++; } catch (eR) {}
    }
    var dir = new Folder(dirPath);
    if (!dir.exists) { return 'folder not found: ' + dirPath; }
    var mogrts = dir.getFiles('*.mogrt');
    if (!mogrts.length) { return 'no .mogrt files in ' + dirPath; }
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
    return 'cleared ' + cleared + ', placed ' + placed +
           (skipped.length ? ', no mogrt for: ' + skipped.join(', ') : '');
}
