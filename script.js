const SECTIONS = [
    {
        title: null,
        columns: [
            { name: 'Employee Code', type: 'text',   required: true, full: true },
            { name: 'Date',          type: 'date',   required: true },
            { name: 'Remarks',       type: 'text',   full: true, placeholder: 'Optional note' },
        ],
    },
    {
        title: 'Privilege Leave',
        columns: [
            { name: 'Privilege Leave Encashable',     type: 'number', label: 'Encashable' },
            { name: 'Privilege Leave Non-Encashable', type: 'number', label: 'Non-Encashable' },
        ],
    },
    {
        title: 'Optional Holiday',
        columns: [
            { name: 'Optional Holiday Encashable',     type: 'number', label: 'Encashable' },
            { name: 'Optional Holiday Non-Encashable', type: 'number', label: 'Non-Encashable' },
        ],
    },
    {
        title: 'Emergency Leave',
        columns: [
            { name: 'Emergency Leave Encashable',     type: 'number', label: 'Encashable' },
            { name: 'Emergency Leave Non-Encashable', type: 'number', label: 'Non-Encashable' },
        ],
    },
    {
        title: 'Comp-off',
        columns: [
            { name: 'Comp-off Encashable',     type: 'number', label: 'Encashable' },
            { name: 'Comp-off Non-Encashable', type: 'number', label: 'Non-Encashable' },
        ],
    },
    {
        title: 'Leave Without Pay',
        columns: [
            { name: 'Leave Without Pay Encashable',     type: 'number', label: 'Encashable' },
            { name: 'Leave Without Pay Non-Encashable', type: 'number', label: 'Non-Encashable' },
        ],
    },
];

const COLUMNS = SECTIONS.flatMap(s => s.columns);
const FILENAME = 'LeaveOpeningBalanceMaster.xlsx';

function slug(name) {
    return 'f_' + name.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

function todayIso() {
    const d = new Date();
    return d.getFullYear() + '-' +
        String(d.getMonth() + 1).padStart(2, '0') + '-' +
        String(d.getDate()).padStart(2, '0');
}

function makeInput(col) {
    const input = document.createElement('input');
    input.id = slug(col.name);
    input.name = col.name;
    if (col.type === 'number') {
        input.type = 'number';
        input.step = 'any';
        input.min = '0';
        input.placeholder = '0';
    } else if (col.type === 'date') {
        input.type = 'date';
        input.value = todayIso();
    } else {
        input.type = 'text';
        if (col.placeholder) input.placeholder = col.placeholder;
    }
    if (col.required) input.required = true;
    return input;
}

function makeFormGroup(col) {
    const group = document.createElement('div');
    group.className = 'form-group' + (col.full ? ' full' : '');

    const label = document.createElement('label');
    label.htmlFor = slug(col.name);
    label.textContent = col.label || col.name;
    if (col.required) {
        const star = document.createElement('span');
        star.className = 'req';
        star.textContent = '*';
        star.setAttribute('aria-label', 'required');
        label.appendChild(star);
    }

    group.appendChild(label);
    group.appendChild(makeInput(col));
    return group;
}

function renderSections() {
    const root = document.getElementById('fields');
    root.innerHTML = '';
    SECTIONS.forEach(section => {
        const wrap = document.createElement('div');
        wrap.className = 'section';

        if (section.title) {
            const h = document.createElement('h3');
            h.className = 'section-title';
            h.textContent = section.title;
            wrap.appendChild(h);
        }

        const grid = document.createElement('div');
        grid.className = 'row-2';
        section.columns.forEach(col => grid.appendChild(makeFormGroup(col)));
        wrap.appendChild(grid);

        root.appendChild(wrap);
    });
}

function buildWorkbook(values) {
    const headers = COLUMNS.map(c => c.name);
    const row = COLUMNS.map(col => {
        const raw = values[col.name];
        if (raw === undefined || raw === null || raw === '') return null;
        if (col.type === 'number') {
            const n = Number(raw);
            return Number.isFinite(n) ? n : raw;
        }
        if (col.type === 'date') {
            const d = new Date(raw);
            return isNaN(d.getTime()) ? raw : d;
        }
        return raw;
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, row], { cellDates: true });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return wb;
}

function workbookToBlob(wb) {
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
}

function collectFormValues() {
    const out = {};
    COLUMNS.forEach(col => {
        const el = document.getElementById(slug(col.name));
        out[col.name] = el ? el.value : '';
    });
    return out;
}

function showResult({ kind, title, message, raw }) {
    const panel = document.getElementById('result-panel');
    const banner = document.getElementById('result-banner');
    const responseEl = document.getElementById('response');

    panel.classList.remove('hidden');
    banner.className = 'result-banner ' + kind;
    banner.innerHTML = '';

    const icon = document.createElement('span');
    icon.className = 'icon';
    icon.textContent = kind === 'success' ? '✓' : '!';

    const body = document.createElement('div');
    body.style.flex = '1';
    const t = document.createElement('div');
    t.className = 'title';
    t.textContent = title;
    const m = document.createElement('div');
    m.textContent = message;
    body.appendChild(t);
    body.appendChild(m);

    banner.appendChild(icon);
    banner.appendChild(body);

    responseEl.textContent = raw;
}

function setSubmitting(busy) {
    const btn = document.getElementById('submit-btn');
    btn.disabled = busy;
    btn.innerHTML = busy
        ? '<span class="spinner"></span>Submitting…'
        : 'Submit';
}

function resetForm() {
    document.getElementById('import-form').reset();
    const dateEl = document.getElementById(slug('Date'));
    if (dateEl) dateEl.value = todayIso();
}

async function submitForm(event) {
    event.preventDefault();

    const form = event.target;
    if (!form.checkValidity()) {
        form.reportValidity();
        return;
    }

    setSubmitting(true);

    try {
        const apiUrl = document.getElementById('api-url').value.trim();
        const aqn = document.getElementById('aqn').value;
        const importBg = document.getElementById('import-bg').value;
        const clearCache = document.getElementById('clear-cache').value;

        const wb = buildWorkbook(collectFormValues());
        const blob = workbookToBlob(wb);
        const file = new File([blob], FILENAME, { type: blob.type });

        const fd = new FormData();
        fd.append('AssemblyQualifiedName', aqn);
        fd.append('ImportInBackground', importBg);
        fd.append('ClearCacheAfterFileImport', clearCache);
        fd.append('file', file, FILENAME);

        const res = await fetch(apiUrl, { method: 'POST', body: fd });
        const text = await res.text();

        let parsed = null;
        try { parsed = JSON.parse(text); } catch (_) { /* not json */ }

        const raw = `HTTP ${res.status}\n\n${text}`;

        if (res.ok && parsed && String(parsed.Status).toLowerCase() === 'success') {
            showResult({
                kind: 'success',
                title: parsed.MessageTitle || 'Imported',
                message: parsed.Message || 'Record imported successfully.',
                raw,
            });
            resetForm();
        } else {
            const title = !res.ok ? `Request failed (HTTP ${res.status})` : 'Import failed';
            const message = (parsed && (parsed.Message || parsed.message))
                || (text || 'Unexpected response from server.');
            showResult({ kind: 'error', title, message, raw });
        }
    } catch (err) {
        showResult({
            kind: 'error',
            title: 'Network error',
            message: (err && err.message) ? err.message : String(err),
            raw: String(err && err.stack || err),
        });
    } finally {
        setSubmitting(false);
    }
}

/* ---------- Admin panel (obscurity, not security) ---------------------
 * Anyone can open DevTools and read this password — it's just a friction
 * gate so casual users don't fiddle with API config. Change as needed. */
const ADMIN_PASSWORD = 'factohr@2026';
const SESSION_KEY = 'lobm_admin_unlocked';

const ADMIN_FIELDS = [
    ['api-url',     'api-url-edit'],
    ['aqn',         'aqn-edit'],
    ['import-bg',   'import-bg-edit'],
    ['clear-cache', 'clear-cache-edit'],
];

function isUnlocked() {
    if (sessionStorage.getItem(SESSION_KEY) === '1') return true;
    const params = new URLSearchParams(window.location.search);
    return params.get('admin') === '1';
}

function showAdminUnlocked() {
    document.getElementById('lock-pane').hidden = true;
    document.getElementById('config-pane').hidden = false;
    document.getElementById('lock-badge').textContent = '🔓';
    // Populate edit inputs from the canonical hidden defaults.
    ADMIN_FIELDS.forEach(([hiddenId, editId]) => {
        const h = document.getElementById(hiddenId);
        const e = document.getElementById(editId);
        if (h && e) e.value = h.value;
        // Sync edits back into the hidden source so submitForm reads them.
        if (e) {
            e.addEventListener('input', () => { h.value = e.value; });
        }
    });
}

function showAdminLocked() {
    document.getElementById('lock-pane').hidden = false;
    document.getElementById('config-pane').hidden = true;
    document.getElementById('lock-badge').textContent = '🔒';
    document.getElementById('admin-password').value = '';
    document.getElementById('lock-error').hidden = true;
    document.getElementById('admin-panel').open = false;
}

function attemptUnlock() {
    const input = document.getElementById('admin-password');
    const errEl = document.getElementById('lock-error');
    if (input.value === ADMIN_PASSWORD) {
        sessionStorage.setItem(SESSION_KEY, '1');
        showAdminUnlocked();
    } else {
        errEl.hidden = false;
        input.focus();
        input.select();
    }
}

function setupAdmin() {
    const panel = document.getElementById('admin-panel');
    const unlocked = isUnlocked();

    if (unlocked) {
        showAdminUnlocked();
    } else {
        showAdminLocked();
        // Block expansion until unlock — autofocus password when opened.
        panel.addEventListener('toggle', () => {
            if (!isUnlocked() && panel.open) {
                setTimeout(() => document.getElementById('admin-password').focus(), 50);
            }
        });
    }

    document.getElementById('admin-unlock').addEventListener('click', attemptUnlock);
    document.getElementById('admin-password').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); attemptUnlock(); }
    });
    document.getElementById('admin-lock').addEventListener('click', () => {
        sessionStorage.removeItem(SESSION_KEY);
        // Strip ?admin=1 (and similar) from the URL so a refresh stays locked.
        const url = new URL(window.location.href);
        url.searchParams.delete('admin');
        window.history.replaceState({}, '', url.toString());
        showAdminLocked();
    });
}

document.addEventListener('DOMContentLoaded', () => {
    renderSections();
    setupAdmin();
    document.getElementById('import-form').addEventListener('submit', submitForm);
    document.getElementById('import-form').addEventListener('reset', () => {
        setTimeout(() => {
            const dateEl = document.getElementById(slug('Date'));
            if (dateEl) dateEl.value = todayIso();
            document.getElementById('result-panel').classList.add('hidden');
        }, 0);
    });
});
