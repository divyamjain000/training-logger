/**
 * Workout logger backend (v8: sheet is the source of truth: Settings, Exercises, JSON data columns, all-set history). Paste into the Google Sheet: Extensions > Apps Script.
 * Keep your existing SECRET value. After pasting: Deploy > Manage deployments > pencil icon >
 * Version: New version > Deploy. The /exec URL stays the same.
 */
const SECRET = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function fmt_(v) {
  // Use the spreadsheet's own zone: Sheets parsed the phone's text in that zone, so it round-trips exactly.
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}
function pj_(v) { if (v === '' || v === null || v === undefined) return null; try { return JSON.parse(v); } catch (e) { return null; } }
function ensure_(name, headers) {
  let sh = sheet_(name);
  if (!sh) { sh = SpreadsheetApp.getActive().insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}
function e1rm_(w, r) { return (w > 0 && r >= 1 && r <= 12) ? w * (1 + r / 30) : 0; }

// App bootstrap: template, targets, last session per exercise, and all-time bests per exercise.
function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== SECRET) return json_({ error: 'unauthorized' });

  const template = sheet_('Template').getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[2])
    .map(r => ({ day: r[0], order: r[1], ex: r[2], sets: r[3], lo: r[4], hi: r[5], rest: r[6], data: pj_(r[7]) }));

  const targets = sheet_('Targets').getDataRange().getValues().slice(1)
    .filter(r => r[2] && r[1] !== 'TEST')
    .map(r => ({ date: fmt_(r[0]).slice(0, 10), day: r[1], ex: r[2], sets: r[3], w: r[4], reps: r[5], why: r[6] }));

  // Bests start from the Records tab (full Strong history), then fold in everything logged in Sets.
  const bests = {};
  const rec = sheet_('Records');
  if (rec && rec.getLastRow() >= 2) {
    rec.getRange(2, 1, rec.getLastRow() - 1, 4).getValues().forEach(r => {
      if (r[0]) bests[r[0]] = { maxW: Number(r[1]) || 0, e1rm: Number(r[2]) || 0, vol: Number(r[3]) || 0 };
    });
  }

  const s = sheet_('Sets');
  const n = s.getLastRow();
  const last = {};
  const hist = {};
  const exd = {};
  const xsets = {};
  if (n >= 2) {
    const vols = {};
    s.getRange(2, 1, n - 1, 6).getValues().forEach(r => {
      const ex = r[2];
      if (!ex) return;
      const day = fmt_(r[0]).slice(0, 10);
      const w = Number(r[4]) || 0, reps = Number(r[5]) || 0;
      if (!last[ex] || last[ex].day < day) last[ex] = { day: day, sets: [] };
      if (last[ex].day === day) last[ex].sets.push([r[3], w, reps]);
      const b = bests[ex] || (bests[ex] = { maxW: 0, e1rm: 0, vol: 0 });
      if (w > 0 && reps > 0 && w > b.maxW) b.maxW = w;
      const e = e1rm_(w, reps);
      if (e > b.e1rm) b.e1rm = Math.round(e * 10) / 10;
      const k = ex + '|' + day;
      vols[k] = (vols[k] || 0) + w * reps;
      (xsets[ex] = xsets[ex] || []).push([fmt_(r[0]), r[1], r[3], w, reps]);
      const dd = (exd[ex] = exd[ex] || {});
      const o = (dd[day] = dd[day] || [0, 0, 0]);
      o[0] = Math.max(o[0], e1rm_(w, reps));
      if (w > 0 && reps > 0 && w > o[1]) o[1] = w;
      o[2] += w * reps;
      const ts = fmt_(r[0]), hk = ts + '|' + r[1];
      (hist[hk] = hist[hk] || { ts: ts, day: r[1], sets: [] }).sets.push([ex, r[3], w, reps]);
    });
    Object.keys(vols).forEach(k => {
      const ex = k.split('|')[0];
      if (vols[k] > bests[ex].vol) bests[ex].vol = vols[k];
    });
  }
  const history = Object.keys(hist).map(k => hist[k]).sort((a, b) => a.ts < b.ts ? 1 : -1).slice(0, 60);
  // Dashboard data: every workout (Workouts tab + anything only in Sets) and per-exercise series.
  const wk = [], seenDay = {}, wByTs = {}, wByDay = {};
  const wsh = sheet_('Workouts');
  if (wsh && wsh.getLastRow() >= 2) {
    wsh.getRange(2, 1, wsh.getLastRow() - 1, 9).getValues().forEach(r => {
      if (!r[0]) return;
      const ts = fmt_(r[3] || r[1]);
      wk.push({ ts: ts, day: r[2], vol: Number(r[6]) || 0, n: Number(r[7]) || 0, dur: Number(r[5]) || 0 });
      seenDay[ts.slice(0, 10)] = true;
      const info = { dur: Number(r[5]) || 0, prs: r[8] ? String(r[8]).split(' | ').length : 0 };
      wByTs[ts] = info; wByDay[ts.slice(0, 10)] = info;
    });
  }
  Object.keys(hist).forEach(k => {
    const h = hist[k];
    if (seenDay[h.ts.slice(0, 10)]) return;
    wk.push({ ts: h.ts, day: h.day, vol: Math.round(h.sets.reduce((a, s) => a + s[2] * s[3], 0)), n: h.sets.length, dur: 0 });
  });
  wk.sort((a, b) => a.ts < b.ts ? -1 : 1);
  history.forEach(h => { const i = wByTs[h.ts] || wByDay[h.ts.slice(0, 10)]; if (i) { h.dur = i.dur; h.prs = i.prs; } });
  const eh = sheet_('ExHistory');
  if (eh && eh.getLastRow() >= 2) {
    eh.getRange(2, 1, eh.getLastRow() - 1, 5).getValues().forEach(r => {
      if (!r[1]) return;
      const d = fmt_(r[0]).slice(0, 10), dd = (exd[r[1]] = exd[r[1]] || {});
      if (!dd[d]) dd[d] = [Number(r[2]) || 0, Number(r[3]) || 0, Number(r[4]) || 0];
    });
  }
  const exseries = {};
  Object.keys(exd).forEach(ex => {
    exseries[ex] = Object.keys(exd[ex]).sort().map(d => [d, Math.round(exd[ex][d][0] * 10) / 10, exd[ex][d][1], Math.round(exd[ex][d][2])]);
  });
  const bodyRows = [];
  const bs = sheet_('Body');
  if (bs && bs.getLastRow() >= 2) {
    const bn = bs.getLastRow(), st = Math.max(2, bn - 59);
    const keys = ['w', 'bf', 'neck', 'sh', 'chest', 'waist', 'hips', 'arm', 'thigh', 'calf', 'note'];
    bs.getRange(st, 1, bn - st + 1, 14).getValues().forEach(r => {
      if (!r[12]) return;
      const o = { id: r[12], ts: fmt_(r[0]) };
      keys.forEach((k, i) => { if (r[i + 1] !== '' && r[i + 1] !== null) o[k] = r[i + 1]; });
      const xd = pj_(r[13]); if (xd) Object.keys(xd).forEach(k => { o[k] = xd[k]; });
      bodyRows.push(o);
    });
  }
  const settings = {};
  const ss = sheet_('Settings');
  if (ss && ss.getLastRow() >= 2) ss.getRange(2, 1, ss.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0]) settings[r[0]] = { v: pj_(r[1]), t: Number(r[2]) || 0 }; });
  const exercises = [];
  const xsh = sheet_('Exercises');
  if (xsh && xsh.getLastRow() >= 2) xsh.getRange(2, 1, xsh.getLastRow() - 1, 6).getValues().forEach(r => { if (r[0]) exercises.push({ name: r[0], bp: r[1] || '', cat: r[2] || '', notes: r[3] || '', data: pj_(r[4]), t: Number(r[5]) || 0 }); });
  return json_({ v: 8, workouts: wk, exseries: exseries, xsets: xsets, body: bodyRows, template: template, targets: targets, last: last, bests: bests, history: history, settings: settings, exercises: exercises });
}

// Body is text/plain JSON: { key, rows: [set rows], finishes: [workout summaries] }
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad json' }); }
  if (body.key !== SECRET) return json_({ error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let added = 0;
    const rows = body.rows || [];
    if (rows.length) {
      const s = sheet_('Sets');
      const n = s.getLastRow();
      const seen = {};
      if (n >= 2) s.getRange(2, 12, n - 1, 1).getValues().forEach(r => { if (r[0]) seen[r[0]] = true; });
      const fresh = rows.filter(r => r.id && !seen[r.id]);
      if (fresh.length) {
        const out = fresh.map(r => [r.ts, r.day, r.ex, r.set, r.w, r.reps, r.rest || '', r.rir || '', r.form || '', r.help || '', r.note || '', r.id]);
        s.getRange(n + 1, 1, out.length, 1).setNumberFormat('@'); // keep the phone's local time as plain text
        s.getRange(n + 1, 1, out.length, 12).setValues(out);
        added = fresh.length;
      }
    }
    // Edits and deletes of sets logged by the app, matched by the ID in column L.
    const dels = body.deletes || [], ups = body.updates || [];
    if (dels.length || ups.length) {
      const s = sheet_('Sets');
      const n = s.getLastRow();
      if (n >= 2) {
        const rowOf = {};
        s.getRange(2, 12, n - 1, 1).getValues().forEach((r, i) => { if (r[0]) rowOf[r[0]] = i + 2; });
        ups.forEach(u => {
          const r = rowOf[u.id]; if (!r) return;
          s.getRange(r, 5, 1, 2).setValues([[u.w, u.reps]]);
          s.getRange(r, 8, 1, 3).setValues([[u.rir || '', u.form || '', u.help || '']]);
        });
        dels.map(id => rowOf[id]).filter(Boolean).sort((a, b) => b - a).forEach(r => s.deleteRow(r));
      }
    }
    // New exercises added to a workout in the app -> Template tab (dedupe by day + exercise).
    const tdels = body.templateDeletes || [];
    if (tdels.length) {
      const t0 = sheet_('Template');
      const n0 = t0.getLastRow();
      if (n0 >= 2) {
        const vals = t0.getRange(2, 1, n0 - 1, 7).getValues();
        for (let i = vals.length - 1; i >= 0; i--) {
          if (tdels.some(x => String(x.day) === String(vals[i][0]) && (!x.ex || String(x.ex) === String(vals[i][2])))) t0.deleteRow(i + 2);
        }
      }
    }
    const tadds = body.templateAdds || [];
    if (tadds.length) {
      const t = sheet_('Template');
      const tn = t.getLastRow();
      const seenT = {}, maxOrd = {};
      (tn >= 2 ? t.getRange(2, 1, tn - 1, 7).getValues() : []).forEach(r => {
        seenT[r[0] + '|' + r[2]] = true;
        maxOrd[r[0]] = Math.max(maxOrd[r[0]] || 0, Number(r[1]) || 0);
      });
      const outT = [];
      tadds.forEach(a => {
        const k = a.day + '|' + a.ex;
        if (seenT[k]) return;
        seenT[k] = true; maxOrd[a.day] = (maxOrd[a.day] || 0) + 1;
        outT.push([a.day, maxOrd[a.day], a.ex, a.sets, a.lo, a.hi, a.rest, a.data ? JSON.stringify(a.data) : '']);
      });
      if (outT.length) t.getRange(tn + 1, 1, outT.length, 8).setValues(outT);
    }
    // Template row JSON (column H): supersets, notes. Matched by day + exercise.
    const tdata = body.templateData || [];
    if (tdata.length) {
      const t1 = sheet_('Template');
      if (!t1.getRange(1, 8).getValue()) t1.getRange(1, 8).setValue('Data');
      const n1 = t1.getLastRow();
      if (n1 >= 2) {
        const vals = t1.getRange(2, 1, n1 - 1, 3).getValues();
        tdata.forEach(d => {
          for (let i = 0; i < vals.length; i++) {
            if (String(vals[i][0]) === String(d.day) && String(vals[i][2]) === String(d.ex)) t1.getRange(i + 2, 8).setValue(d.data ? JSON.stringify(d.data) : '');
          }
        });
      }
    }
    // Settings key/value JSON, last write wins by t.
    const sets_ = body.settings || [];
    if (sets_.length) {
      const sh = ensure_('Settings', ['Key', 'Value (JSON)', 'UpdatedAt']);
      const sn = sh.getLastRow(), rowOfK = {}, tOf = {};
      if (sn >= 2) sh.getRange(2, 1, sn - 1, 3).getValues().forEach((r, i) => { rowOfK[r[0]] = i + 2; tOf[r[0]] = Number(r[2]) || 0; });
      sets_.forEach(o => {
        if (!o.key) return;
        if (rowOfK[o.key]) { if ((o.t || 0) >= tOf[o.key]) sh.getRange(rowOfK[o.key], 2, 1, 2).setValues([[JSON.stringify(o.value), o.t || Date.now()]]); }
        else { sh.appendRow([o.key, JSON.stringify(o.value), o.t || Date.now()]); rowOfK[o.key] = sh.getLastRow(); tOf[o.key] = o.t || 0; }
      });
    }
    // Exercises: name, body part, category, notes, data JSON, updatedAt. Upsert by name.
    const exs = body.exercises || [];
    if (exs.length) {
      const sh = ensure_('Exercises', ['Name', 'Body part', 'Category', 'Notes', 'Data (JSON)', 'UpdatedAt']);
      const en = sh.getLastRow(), rowOfN = {}, tOfN = {};
      if (en >= 2) sh.getRange(2, 1, en - 1, 6).getValues().forEach((r, i) => { rowOfN[r[0]] = i + 2; tOfN[r[0]] = Number(r[5]) || 0; });
      exs.forEach(o => {
        if (!o.name) return;
        const row = [o.name, o.bp || '', o.cat || '', o.notes || '', o.data ? JSON.stringify(o.data) : '', o.t || Date.now()];
        if (rowOfN[o.name]) { if ((o.t || 0) >= tOfN[o.name]) sh.getRange(rowOfN[o.name], 1, 1, 6).setValues([row]); }
        else { sh.appendRow(row); rowOfN[o.name] = sh.getLastRow(); tOfN[o.name] = o.t || 0; }
      });
    }
    // Body measurements: tab "Body", ID in column M, extra metrics as JSON in column N.
    const bodies = body.body || [], bdels = body.bodyDeletes || [];
    if (bodies.length || bdels.length) {
      let b = sheet_('Body');
      if (!b) {
        b = SpreadsheetApp.getActive().insertSheet('Body');
        b.appendRow(['Timestamp', 'Weight (kg)', 'Body fat (%)', 'Neck (cm)', 'Shoulders (cm)', 'Chest (cm)', 'Waist (cm)', 'Hips (cm)', 'Arm (cm)', 'Thigh (cm)', 'Calf (cm)', 'Note', 'ID', 'Data (JSON)']);
        b.setFrozenRows(1);
      }
      const bn = b.getLastRow();
      if (!b.getRange(1, 14).getValue()) b.getRange(1, 14).setValue('Data (JSON)');
      const rowOf = {};
      if (bn >= 2) b.getRange(2, 13, bn - 1, 1).getValues().forEach((r, i) => { if (r[0]) rowOf[r[0]] = i + 2; });
      const fresh = bodies.filter(e => e.id && !rowOf[e.id]);
      if (fresh.length) {
        const FIX = { kind: 1, id: 1, ts: 1, w: 1, bf: 1, neck: 1, sh: 1, chest: 1, waist: 1, hips: 1, arm: 1, thigh: 1, calf: 1, note: 1 };
        const out = fresh.map(e => {
          const xd = {}; Object.keys(e).forEach(k => { if (!FIX[k] && e[k] !== '' && e[k] !== null) xd[k] = e[k]; });
          return [e.ts, e.w || '', e.bf || '', e.neck || '', e.sh || '', e.chest || '', e.waist || '', e.hips || '', e.arm || '', e.thigh || '', e.calf || '', e.note || '', e.id, Object.keys(xd).length ? JSON.stringify(xd) : ''];
        });
        b.getRange(bn + 1, 1, out.length, 1).setNumberFormat('@');
        b.getRange(bn + 1, 1, out.length, 14).setValues(out);
      }
      bdels.map(id => rowOf[id]).filter(Boolean).sort((a, c) => c - a).forEach(r => b.deleteRow(r));
    }
    const fins = body.finishes || [];
    if (fins.length) {
      let w = sheet_('Workouts');
      if (!w) {
        w = SpreadsheetApp.getActive().insertSheet('Workouts');
        w.appendRow(['ID', 'Date', 'Workout', 'Start', 'End', 'Duration (min)', 'Total volume (kg)', 'Sets', 'PRs']);
        w.setFrozenRows(1);
      }
      const wn = w.getLastRow();
      const seenW = {};
      if (wn >= 2) w.getRange(2, 1, wn - 1, 1).getValues().forEach(r => { if (r[0]) seenW[r[0]] = true; });
      fins.filter(f => f.id && !seenW[f.id]).forEach(f => {
        w.getRange(w.getLastRow() + 1, 2, 1, 4).setNumberFormat('@');
        w.appendRow([f.id, f.date, f.day, f.start, f.end, f.durMin, f.volume, f.sets, f.prs || '']);
      });
    }
    return json_({
      ok: true, added: added,
      ids: rows.map(r => r.id).concat(fins.map(f => f.id), dels, ups.map(u => u.id), bodies.map(e => e.id), bdels, tadds.map(a => a.id), tdels.map(a => a.id), tdata.map(a => a.id), sets_.map(a => a.id), exs.map(a => a.id))
    });
  } finally {
    lock.releaseLock();
  }
}
