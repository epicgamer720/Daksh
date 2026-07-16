/*
 add_labels.jsx — adds an editable text layer above every clip in the active sequence.

 One-time setup in Premiere (this is where your font/size/style comes from):
   1. Make one text graphic styled exactly how you want, select it, then
      Graphics and Titles menu > Export As Motion Graphics Template… (save anywhere).
   2. Delete that graphic from the timeline.

 Then after importing a Clip Chopper timeline XML:
   File > Scripts > Run Script… > pick this file > choose your saved .mogrt.

 Each copy's text is set to the clip's label (the clip name minus its "01 - " prefix)
 and it spans exactly its clip. Labels land on the top video track, which must be
 empty — if you generated the XML with PNG labels on, delete those or add a fresh
 track on top (right-click a track header > Add Track) before running.

 If your Premiere version refuses to set text from a script, the graphics still
 appear in place — edit their text in Window > Essential Graphics.
*/

(function () {
    var seq = app.project.activeSequence;
    if (!seq) {
        alert('Open the imported timeline first, then re-run this script.');
        return;
    }

    var vTracks = seq.videoTracks;
    var target = vTracks.numTracks - 1; // top video track
    if (target <= 0) {
        alert('Add an empty video track above the clips first ' +
              '(right-click a track header > Add Track), then re-run.');
        return;
    }
    var clips = vTracks[0].clips;
    if (clips.numItems === 0) {
        alert('No clips found on V1 of the active sequence.');
        return;
    }
    if (vTracks[target].clips.numItems > 0) {
        alert('The top video track (V' + (target + 1) + ') is not empty ' +
              '(maybe the PNG labels from the XML). Delete its clips or add a new ' +
              'empty track on top, then re-run.');
        return;
    }

    var mogrt = File.openDialog('Pick your exported text .mogrt', '*.mogrt');
    if (!mogrt) {
        return;
    }

    function setText(item, text) {
        var comp = item.getMGTComponent();
        if (!comp) {
            return false;
        }
        var props = comp.properties;
        var param = null;
        try { param = props.getParamForDisplayName('Source Text'); } catch (e) {}
        if (!param) {
            for (var j = 0; j < props.numItems; j++) {
                try {
                    if (typeof props[j].getValue() === 'string') {
                        param = props[j];
                        break;
                    }
                } catch (e2) {}
            }
        }
        if (!param) {
            return false;
        }
        try { param.setValue(text, true); return true; } catch (e3) {}
        try {
            param.setValue('{"textEditValue":"' + text.replace(/"/g, '\\"') + '"}', true);
            return true;
        } catch (e4) {}
        return false;
    }

    var added = 0, textSet = 0, failures = [];
    for (var i = 0; i < clips.numItems; i++) {
        var clip = clips[i];
        var label = String(clip.name).replace(/^\s*\d+\s*-\s*/, '');
        var item = seq.importMGT(mogrt.fsName, clip.start.ticks, target, 0);
        if (!item) {
            failures.push(clip.name);
            continue;
        }
        added++;
        try { // stretch the graphic to cover its clip exactly
            var endT = new Time();
            endT.ticks = clip.end.ticks;
            item.end = endT;
        } catch (e5) {}
        try {
            if (setText(item, label)) {
                textSet++;
            }
            item.name = label;
        } catch (e6) {}
    }

    var msg = 'Done: ' + added + ' label graphics placed, text auto-set on ' + textSet + '.';
    if (failures.length) {
        msg += '\nCould not place: ' + failures.join(', ');
    }
    if (textSet < added) {
        msg += '\nFor the rest: select the graphic and edit its text in Window > Essential Graphics.';
    }
    alert(msg);
})();
