/* =====================================================================
   mysongs.js - Reproductor, buscador y formularios CRUD de PoliMusic
   ===================================================================== */
(function () {
    'use strict';

    var body = document.body;
    var token = body.getAttribute('data-token');
    var audio = document.getElementById('psAudio');
    var rows = Array.prototype.slice.call(document.querySelectorAll('#psSongs tbody tr[data-id]'));
    var emptyRow = document.getElementById('psEmpty');
    var current = -1;
    var counted = false;

    var nowTitle = document.getElementById('psNowTitle');
    var nowSub = document.getElementById('psNowSub');
    var toggleBtn = document.getElementById('psToggle');
    var seek = document.getElementById('psSeek');
    var vol = document.getElementById('psVol');
    var cur = document.getElementById('psCur');
    var dur = document.getElementById('psDur');
    var toast = document.getElementById('psToast');
    var toastTimer;

    function fmt(sec) {
        if (!isFinite(sec) || sec < 0) { sec = 0; }
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ':' + (s < 10 ? '0' : '') + s;
    }

    function fill(range) {
        var pct = (range.value - range.min) / (range.max - range.min) * 100;
        range.style.setProperty('--fill', pct + '%');
    }

    function showToast(msg) {
        toast.textContent = msg;
        toast.hidden = false;
        clearTimeout(toastTimer);
        toastTimer = setTimeout(function () { toast.hidden = true; }, 3500);
    }

    function visibleRows() {
        return rows.filter(function (r) { return !r.hidden; });
    }

    /* ---------- Reproductor ---------- */
    function load(index) {
        var row = rows[index];
        if (!row) { return; }
        if (current >= 0 && rows[current]) { rows[current].classList.remove('is-current'); }
        current = index;
        counted = false;
        row.classList.add('is-current');
        audio.src = row.getAttribute('data-src');
        nowTitle.textContent = row.getAttribute('data-name');
        nowSub.textContent = row.getAttribute('data-file');
        seek.value = 0;
        fill(seek);
    }

    function playIndex(index) {
        if (index !== current) { load(index); }
        var p = audio.play();
        if (p && p.catch) { p.catch(function () { }); }
    }

    function toggle() {
        if (current < 0) {
            var first = visibleRows()[0];
            if (first) { playIndex(rows.indexOf(first)); }
            return;
        }
        if (audio.paused) { playIndex(current); } else { audio.pause(); }
    }

    function step(dir) {
        var list = visibleRows();
        if (!list.length) { return; }
        var pos = list.indexOf(rows[current]);
        var next = list[(pos + dir + list.length) % list.length];
        playIndex(rows.indexOf(next));
    }

    // Suma una reproduccion en la base de datos (una vez por cada vez que empieza la cancion)
    function countPlay(row) {
        var data = new URLSearchParams();
        data.append('action', 'play');
        data.append('id', row.getAttribute('data-id'));
        data.append('token', token);
        fetch('mysongs.asp', { method: 'POST', body: data, credentials: 'same-origin' })
            .then(function (res) { return res.json(); })
            .then(function (json) {
                if (!json.ok) { return; }
                var badge = row.querySelector('.ps-plays');
                var old = parseInt(row.getAttribute('data-plays'), 10) || 0;
                row.setAttribute('data-plays', json.count);
                badge.textContent = json.count;
                badge.classList.add('bump');
                setTimeout(function () { badge.classList.remove('bump'); }, 250);
                var total = document.getElementById('psStatPlays');
                total.textContent = (parseInt(total.textContent, 10) || 0) + (json.count - old);
            })
            .catch(function () { });
    }

    rows.forEach(function (row, index) {
        row.querySelector('.ps-play').addEventListener('click', function () {
            if (index === current && !audio.paused) { audio.pause(); } else { playIndex(index); }
        });
    });

    toggleBtn.addEventListener('click', toggle);
    document.getElementById('psPrev').addEventListener('click', function () { step(-1); });
    document.getElementById('psNext').addEventListener('click', function () { step(1); });

    audio.addEventListener('play', function () {
        body.classList.add('is-playing');
        toggleBtn.setAttribute('aria-label', 'Pause');
        if (!counted && rows[current]) {
            counted = true;
            countPlay(rows[current]);
        }
    });
    audio.addEventListener('pause', function () {
        body.classList.remove('is-playing');
        toggleBtn.setAttribute('aria-label', 'Play');
    });
    audio.addEventListener('ended', function () {
        counted = false;
        step(1);
    });
    audio.addEventListener('loadedmetadata', function () {
        dur.textContent = fmt(audio.duration);
    });
    audio.addEventListener('timeupdate', function () {
        cur.textContent = fmt(audio.currentTime);
        if (audio.duration) {
            seek.value = audio.currentTime / audio.duration * 100;
            fill(seek);
        }
    });
    audio.addEventListener('error', function () {
        if (current < 0) { return; }
        body.classList.remove('is-playing');
        showToast('Could not load the audio file "' + rows[current].getAttribute('data-file') + '"');
    });

    seek.addEventListener('input', function () {
        fill(seek);
        if (audio.duration) { audio.currentTime = seek.value / 100 * audio.duration; }
    });
    vol.addEventListener('input', function () {
        audio.volume = parseFloat(vol.value);
        fill(vol);
    });
    audio.volume = parseFloat(vol.value);
    fill(vol);
    fill(seek);

    // Barra espaciadora = play / pausa (si no se esta escribiendo)
    document.addEventListener('keydown', function (e) {
        var tag = (e.target.tagName || '').toLowerCase();
        if (e.code === 'Space' && tag !== 'input' && tag !== 'button' && !document.querySelector('dialog[open]')) {
            e.preventDefault();
            toggle();
        }
    });

    /* ---------- Buscador ---------- */
    var search = document.getElementById('psSearch');
    search.addEventListener('input', function () {
        var q = search.value.trim().toLowerCase();
        var shown = 0;
        rows.forEach(function (row) {
            var text = (row.getAttribute('data-name') + ' ' + row.getAttribute('data-file')).toLowerCase();
            row.hidden = q !== '' && text.indexOf(q) === -1;
            if (!row.hidden) { shown++; }
        });
        if (rows.length) { emptyRow.hidden = shown > 0; }
    });

    /* ---------- Mensajes flash ---------- */
    Array.prototype.forEach.call(document.querySelectorAll('.ps-flash'), function (flash) {
        var close = flash.querySelector('.ps-flash-close');
        if (!close) { return; }
        function hide() {
            flash.style.opacity = '0';
            setTimeout(function () { flash.remove(); }, 300);
        }
        close.addEventListener('click', hide);
        setTimeout(hide, 6000);
    });

    /* ---------- CRUD (solo existe en la pagina si el rol es admin) ---------- */
    var songDialog = document.getElementById('psSongDialog');
    var deleteDialog = document.getElementById('psDeleteDialog');
    if (!songDialog || !deleteDialog) { return; }

    function open(dialog) {
        if (typeof dialog.showModal === 'function') { dialog.showModal(); } else { dialog.setAttribute('open', ''); }
    }
    function close(dialog) {
        if (typeof dialog.close === 'function') { dialog.close(); } else { dialog.removeAttribute('open'); }
    }

    [songDialog, deleteDialog].forEach(function (dialog) {
        dialog.querySelector('[data-close]').addEventListener('click', function () { close(dialog); });
        // Clic fuera del cuadro = cerrar
        dialog.addEventListener('click', function (e) { if (e.target === dialog) { close(dialog); } });
    });

    function openSongForm(row) {
        var isEdit = !!row;
        document.getElementById('psFormTitle').textContent = isEdit ? 'Edit song' : 'Add song';
        document.getElementById('psFormAction').value = isEdit ? 'edit' : 'add';
        document.getElementById('psFormId').value = isEdit ? row.getAttribute('data-id') : '';
        document.getElementById('psFormName').value = isEdit ? row.getAttribute('data-name') : '';
        document.getElementById('psFormFile').value = isEdit ? row.getAttribute('data-file') : '';
        document.getElementById('psFormPlays').value = isEdit ? row.getAttribute('data-plays') : 0;
        document.getElementById('psFormPlaysWrap').hidden = !isEdit;
        open(songDialog);
        document.getElementById('psFormName').focus();
    }

    document.getElementById('psAddBtn').addEventListener('click', function () { openSongForm(null); });

    rows.forEach(function (row) {
        row.querySelector('.ps-edit').addEventListener('click', function () { openSongForm(row); });
        row.querySelector('.ps-delete').addEventListener('click', function () {
            document.getElementById('psDeleteId').value = row.getAttribute('data-id');
            document.getElementById('psDeleteName').textContent = row.getAttribute('data-name');
            open(deleteDialog);
        });
    });
})();
