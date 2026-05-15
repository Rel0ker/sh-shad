/**
 * Главный экран: справочники, строки изменений, сохранение, ссылки экспорта.
 */

let reference = { teachers: [], classes: [], subjects: [] };

function currentDate() {
  return document.getElementById('change-date').value;
}

function toYMD(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function syncPrikazDatesFromChangeDate() {
  const raw = currentDate();
  if (!raw) return;
  const d = new Date(raw + 'T12:00:00');
  const from = new Date(d.getFullYear(), d.getMonth(), 1);
  const fromEl = document.getElementById('prikaz-from');
  const toEl = document.getElementById('prikaz-to');
  const ordEl = document.getElementById('prikaz-order-date');
  if (fromEl && !fromEl.dataset.touched) fromEl.value = toYMD(from);
  if (toEl && !toEl.dataset.touched) toEl.value = raw;
  if (ordEl && !ordEl.dataset.touched) ordEl.value = toYMD(new Date());
}

function updateExportLinks() {
  const d = encodeURIComponent(currentDate());
  const base = '';
  document.getElementById('exp-xlsx-t').href = `${base}/export/xlsx/teachers?date=${d}`;
  document.getElementById('exp-xlsx-s').href = `${base}/export/xlsx/students?date=${d}`;
  document.getElementById('exp-pdf-t').href = `${base}/export/pdf/teachers?date=${d}`;
  document.getElementById('exp-pdf-s').href = `${base}/export/pdf/students?date=${d}`;
  document.getElementById('exp-png-t').href = `${base}/export/png/teachers?date=${d}`;
  document.getElementById('exp-png-s').href = `${base}/export/png/students?date=${d}`;
  const d1 = document.getElementById('prikaz-from')?.value;
  const d2 = document.getElementById('prikaz-to')?.value;
  const od = document.getElementById('prikaz-order-date')?.value;
  const pr = document.getElementById('exp-xlsx-prikaz');
  if (pr && d1 && d2) {
    let u = `${base}/export/xlsx/prikaz?date_from=${encodeURIComponent(d1)}&date_to=${encodeURIComponent(d2)}`;
    if (od) u += `&order_date=${encodeURIComponent(od)}`;
    pr.href = u;
  }
}

function api(path, opts) {
  return fetch(path, Object.assign({ headers: { Accept: 'application/json' } }, opts)).then(async (r) => {
    if (!r.ok) {
      const t = await r.text();
      throw new Error(t || r.statusText);
    }
    if (r.status === 204) return null;
    return r.json();
  });
}

function fillDatalists() {
  const dlC = document.getElementById('dl-classes');
  const dlT = document.getElementById('dl-teachers');
  const dlS = document.getElementById('dl-subjects');
  dlC.innerHTML = '';
  dlT.innerHTML = '';
  dlS.innerHTML = '';
  reference.classes.forEach((c) => {
    const o = document.createElement('option');
    o.value = c.name;
    o.dataset.id = String(c.id);
    o.dataset.shift = String(c.shift);
    dlC.appendChild(o);
  });
  reference.teachers.forEach((t) => {
    const o = document.createElement('option');
    o.value = t.fio;
    dlT.appendChild(o);
  });
  reference.subjects.forEach((s) => {
    const o = document.createElement('option');
    o.value = s.name;
    dlS.appendChild(o);
  });
}

function lessonOptionsForShift(shift) {
  if (shift === 2) {
    return [-1, 0, 1, 2, 3, 4, 5, 6];
  }
  return [0, 1, 2, 3, 4, 5, 6, 7];
}

function shiftForClassName(name) {
  const n = (name || '').trim();
  if (!n) return 1;
  const c = reference.classes.find((x) => x.name.toLowerCase() === n.toLowerCase());
  return c ? c.shift : 1;
}

function rebuildLessonSelect(selectEl, shift, currentVal) {
  const opts = lessonOptionsForShift(shift);
  selectEl.innerHTML = '';
  opts.forEach((n) => {
    const o = document.createElement('option');
    o.value = String(n);
    o.textContent = String(n);
    selectEl.appendChild(o);
  });
  const s = String(currentVal);
  if (opts.includes(Number(s))) selectEl.value = s;
  else selectEl.value = String(opts[0]);
}

function addRow(data) {
  const tbody = document.getElementById('changes-body');
  const tr = document.createElement('tr');
  const klass = (data && data.klass) || '';
  const shift = data && data.class_id
    ? (reference.classes.find((c) => c.id === data.class_id) || {}).shift || shiftForClassName(klass)
    : shiftForClassName(klass);

  tr.innerHTML = `
    <td><input type="text" class="inp-klass" list="dl-classes" value=""></td>
    <td class="cell-num"><select class="inp-lesson"></select></td>
    <td><input type="text" class="inp-absent" list="dl-teachers" value=""></td>
    <td><input type="text" class="inp-repl" list="dl-teachers" value=""></td>
    <td><input type="text" class="inp-subj" list="dl-subjects" value=""></td>
    <td><input type="text" class="inp-room" value=""></td>
    <td><input type="text" class="inp-note" value=""></td>
    <td class="cell-actions"><button type="button" class="btn-icon" title="Удалить">×</button></td>
  `;
  tbody.appendChild(tr);

  const inpKlass = tr.querySelector('.inp-klass');
  const selLesson = tr.querySelector('.inp-lesson');
  inpKlass.value = klass;
  rebuildLessonSelect(selLesson, shift, (data && data.lesson_no) ?? lessonOptionsForShift(shift)[0]);

  tr.querySelector('.inp-absent').value = (data && data.absent_fio) || '';
  tr.querySelector('.inp-repl').value = (data && data.replacement_fio) || '';
  tr.querySelector('.inp-subj').value = (data && data.subject) || '';
  tr.querySelector('.inp-room').value = (data && data.room) || '';
  tr.querySelector('.inp-note').value = (data && data.note) || '';

  inpKlass.addEventListener('change', () => {
    const sh = shiftForClassName(inpKlass.value);
    rebuildLessonSelect(selLesson, sh, selLesson.value);
  });
  inpKlass.addEventListener('blur', () => {
    const sh = shiftForClassName(inpKlass.value);
    rebuildLessonSelect(selLesson, sh, selLesson.value);
  });

  tr.querySelector('.btn-icon').addEventListener('click', () => tr.remove());
}

function collectRows() {
  const rows = [];
  document.querySelectorAll('#changes-body tr').forEach((tr) => {
    const klass = tr.querySelector('.inp-klass').value.trim();
    if (!klass) return;
    const name = klass;
    const cls = reference.classes.find((c) => c.name.toLowerCase() === name.toLowerCase());
    const class_id = cls ? cls.id : null;
    rows.push({
      class_id,
      klass,
      lesson_no: parseInt(tr.querySelector('.inp-lesson').value, 10),
      absent_fio: tr.querySelector('.inp-absent').value.trim(),
      replacement_fio: tr.querySelector('.inp-repl').value.trim(),
      subject: tr.querySelector('.inp-subj').value.trim(),
      room: tr.querySelector('.inp-room').value.trim(),
      note: tr.querySelector('.inp-note').value.trim(),
    });
  });
  return rows;
}

async function loadDay() {
  const d = currentDate();
  let rows = [];
  try {
    const data = await api(`/api/changes?date=${encodeURIComponent(d)}`);
    rows = data.rows || [];
  } catch (e) {
    console.warn(e);
  }
  document.getElementById('changes-body').innerHTML = '';
  if (rows.length === 0) addRow(null);
  else rows.forEach((r) => addRow(r));
  syncPrikazDatesFromChangeDate();
  updateExportLinks();
}

async function init() {
  reference = await api('/api/reference');
  fillDatalists();
  document.getElementById('change-date').addEventListener('change', () => loadDay());
  ['prikaz-from', 'prikaz-to', 'prikaz-order-date'].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('change', () => {
      el.dataset.touched = '1';
      updateExportLinks();
    });
  });
  document.getElementById('btn-add-row').addEventListener('click', () => addRow(null));
  document.getElementById('btn-save').addEventListener('click', async () => {
    try {
      const rows = collectRows();
      await api('/api/changes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date: currentDate(), rows }),
      });
      alert('Сохранено');
      updateExportLinks();
    } catch (e) {
      alert('Ошибка: ' + e.message);
    }
  });
  await loadDay();
}

init().catch((e) => alert('Не удалось загрузить: ' + e.message));
