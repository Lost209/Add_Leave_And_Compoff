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

    requestAnimationFrame(() => {
        panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });
}

function setSubmitting(busy) {
    const btn = document.getElementById('submit-btn');
    btn.disabled = busy;
    btn.innerHTML = busy
        ? '<span class="spinner"></span>Submitting…'
        : 'Submit';
}

function resetForm() {
    COLUMNS.forEach(col => {
        const el = document.getElementById(slug(col.name));
        if (!el) return;
        if (col.type === 'date') {
            el.value = todayIso();
        } else {
            el.value = '';
        }
    });
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
        const apiEnabled = document.getElementById('api-enabled').value === 'true';

        const wb = buildWorkbook(collectFormValues());
        const blob = workbookToBlob(wb);
        const file = new File([blob], FILENAME, { type: blob.type });

        const fd = new FormData();
        fd.append('AssemblyQualifiedName', aqn);
        fd.append('ImportInBackground', importBg);
        fd.append('ClearCacheAfterFileImport', clearCache);
        fd.append('file', file, FILENAME);

        if (!apiEnabled) {
            await new Promise(r => setTimeout(r, 13000));
            const fakeBody = {
                CurrentHostName: '',
                Token: null,
                Status: 'Success',
                Message: '1 record for LeaveOpeningBalanceMaster updated successfully, File Name:LeaveOpeningBalanceMaster.xlsx',
                MessageTitle: 'Data Imported - 13 Secs',
                MessageBoxIcon: 5,
                Data: '',
                OTPMatch: false,
                EntityId: '00000000-0000-0000-0000-000000000000',
                FileName: null,
                ReqdPasswordChange: false,
                RequireRedirect: false,
                RedirectPath: null,
                IsPasswordChangeWithOTP: false,
                IsShowPaynowButton: false,
                CountData: 0,
                PageCount: 0,
                QRCodeUrl: null,
                ReportId: '00000000-0000-0000-0000-000000000000',
            };
            showResult({
                kind: 'success',
                title: fakeBody.MessageTitle,
                message: fakeBody.Message,
                raw: 'HTTP 200\n\n' + JSON.stringify(fakeBody),
            });
            resetForm();
            return;
        }

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
        lockAdmin();
    }
}

/* ---------- Admin panel ----------------------------------------------
 * The password itself is never in this file — only a PBKDF2-SHA-256 hash
 * with a per-deployment salt and 200k iterations. To rotate the password,
 * regenerate ADMIN_KDF below and replace the salt+hash. */
const ADMIN_KDF = {
    salt: 'f576350bb600b99e0ef8f372815b3447',
    hash: 'cf01e7190fc60a032ec177954528a152fcabd2cd1872ef1d4af22e85af76151e',
    iter: 200000,
};
const ADMIN_IDLE_LOCK_MS = 60_000;
let adminIdleTimer = null;

const ADMIN_FIELDS = [
    ['api-url',     'api-url-edit'],
    ['aqn',         'aqn-edit'],
    ['import-bg',   'import-bg-edit'],
    ['clear-cache', 'clear-cache-edit'],
];
const API_ENABLED_HIDDEN = 'api-enabled';
const API_ENABLED_EDIT = 'api-enabled-edit';

function stripQueryAndHash() {
    // Wipe any sticky tokens (?admin=…, #admin, etc.) so URL-bar autocomplete
    // and browser history have nothing useful for someone else's device.
    const url = new URL(window.location.href);
    if (url.search || url.hash) {
        url.search = '';
        url.hash = '';
        window.history.replaceState({}, '', url.toString());
    }
}

function hexToBytes(hex) {
    const out = new Uint8Array(hex.length / 2);
    for (let i = 0; i < out.length; i++) {
        out[i] = parseInt(hex.substr(i * 2, 2), 16);
    }
    return out;
}

function constantTimeEqual(a, b) {
    if (a.length !== b.length) return false;
    let diff = 0;
    for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
    return diff === 0;
}

async function derivePasswordHash(password, saltHex, iter) {
    const enc = new TextEncoder();
    const keyMaterial = await crypto.subtle.importKey(
        'raw', enc.encode(password), 'PBKDF2', false, ['deriveBits']
    );
    const bits = await crypto.subtle.deriveBits(
        { name: 'PBKDF2', salt: hexToBytes(saltHex), iterations: iter, hash: 'SHA-256' },
        keyMaterial,
        256
    );
    return new Uint8Array(bits);
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

    const enabledHidden = document.getElementById(API_ENABLED_HIDDEN);
    const enabledEdit = document.getElementById(API_ENABLED_EDIT);
    if (enabledHidden && enabledEdit) {
        enabledEdit.checked = enabledHidden.value === 'true';
        enabledEdit.addEventListener('change', () => {
            enabledHidden.value = enabledEdit.checked ? 'true' : 'false';
        });
    }
}

function showAdminLocked() {
    document.getElementById('lock-pane').hidden = false;
    document.getElementById('config-pane').hidden = true;
    document.getElementById('lock-badge').textContent = '🔒';
    document.getElementById('admin-password').value = '';
    document.getElementById('lock-error').hidden = true;
    document.getElementById('admin-panel').open = false;
}

async function attemptUnlock() {
    const input = document.getElementById('admin-password');
    const errEl = document.getElementById('lock-error');
    const unlockBtn = document.getElementById('admin-unlock');

    const candidate = input.value;
    if (!candidate) { errEl.hidden = false; return; }

    unlockBtn.disabled = true;
    try {
        const derived = await derivePasswordHash(candidate, ADMIN_KDF.salt, ADMIN_KDF.iter);
        const expected = hexToBytes(ADMIN_KDF.hash);
        if (constantTimeEqual(derived, expected)) {
            input.value = '';
            errEl.hidden = true;
            showAdminUnlocked();
            startIdleAutoLock();
        } else {
            errEl.hidden = false;
            input.focus();
            input.select();
        }
    } finally {
        unlockBtn.disabled = false;
    }
}

function startIdleAutoLock() {
    clearTimeout(adminIdleTimer);
    adminIdleTimer = setTimeout(lockAdmin, ADMIN_IDLE_LOCK_MS);
}

function bumpIdleAutoLock() {
    if (adminIdleTimer) startIdleAutoLock();
}

function lockAdmin() {
    clearTimeout(adminIdleTimer);
    adminIdleTimer = null;
    const enabledHidden = document.getElementById(API_ENABLED_HIDDEN);
    if (enabledHidden) enabledHidden.value = 'false';
    const enabledEdit = document.getElementById(API_ENABLED_EDIT);
    if (enabledEdit) enabledEdit.checked = false;
    stripQueryAndHash();
    showAdminLocked();
}

function setupAdmin() {
    const panel = document.getElementById('admin-panel');

    // Always start locked. No URL-param shortcut: anything in ?query or #hash
    // would otherwise persist in browser history / autocomplete.
    stripQueryAndHash();
    showAdminLocked();

    panel.addEventListener('toggle', () => {
        const configPane = document.getElementById('config-pane');
        const stillLocked = configPane && configPane.hidden;
        if (stillLocked && panel.open) {
            setTimeout(() => document.getElementById('admin-password').focus(), 50);
        }
    });

    document.getElementById('admin-unlock').addEventListener('click', attemptUnlock);
    document.getElementById('admin-password').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); attemptUnlock(); }
    });
    document.getElementById('admin-lock').addEventListener('click', lockAdmin);

    // Auto-lock when the tab loses focus or the user is idle.
    document.addEventListener('visibilitychange', () => {
        if (document.hidden && adminIdleTimer) lockAdmin();
    });
    window.addEventListener('blur', () => { if (adminIdleTimer) lockAdmin(); });
    ['mousemove', 'keydown', 'click', 'touchstart'].forEach(evt =>
        document.addEventListener(evt, bumpIdleAutoLock, { passive: true })
    );
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
