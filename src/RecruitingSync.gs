/**
 * CF_RECRUTING — bezpieczny importer ROOFER v0.1.
 * Najpierw previewRecruitingSync(); bez automatycznego zapisu.
 * Konfiguracja identyfikatorów w Script Properties, NIE w publicznym kodzie.
 */
const RECRUITING_SOURCES = Object.freeze([
  { property: 'CF_RECRUTING_SOURCE_NEODAK_ID', tab: 'Roofers', label: 'NEODAK_RecrutingSources', profile: 'ROOFER' }
]);
const RECRUITING_TARGET = 'KANDYDACI';
const RECRUITING_KEY_COLUMN = 45; // AS, techniczne SOURCE_KEY
const RECRUITING_BATCH_LIMIT = 50;
function recruitingConfig_() {
  const props = PropertiesService.getScriptProperties();
  const targetId = props.getProperty('CF_RECRUTING_TARGET_ID');
  if (!targetId) throw new Error('Brak CF_RECRUTING_TARGET_ID w Script Properties');
  return { targetId: targetId, sources: RECRUITING_SOURCES.map(s => {
    const id = props.getProperty(s.property);
    if (!id) throw new Error('Brak ' + s.property);
    return Object.assign({ id: id }, s);
  }) };
}
function recruitingText_(x) { return String(x == null ? '' : x).trim(); }
function recruitingNorm_(x) { return recruitingText_(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' '); }
function recruitingPhone_(x) { return recruitingText_(x).replace(/[^0-9]/g, ''); }
function recruitingYN_(x) {
  const s = recruitingNorm_(x);
  if (['yes','tak','true','1'].includes(s)) return 'TAK';
  if (['no','nie','false','0'].includes(s)) return 'NIE';
  return '—';
}
function recruitingDate_(x) {
  if (x instanceof Date) return x;
  const s = recruitingText_(x);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]),Number(m[2])-1,Number(m[3])) : '';
}
function recruitingHeaders_(r) {
  const m = {};
  r.forEach((x,i) => { const k = recruitingText_(x).toUpperCase(); if (k) m[k]=i; });
  return m;
}
function recruitingRead_() {
  const cfg = recruitingConfig_();
  const target = SpreadsheetApp.openById(cfg.targetId).getSheetByName(RECRUITING_TARGET);
  if (!target) throw new Error('Brak zakładki KANDYDACI');
  const nr = Math.max(target.getLastRow()-1,0);
  const existing = nr ? target.getRange(2,1,nr,45).getValues() : [];
  const byKey = new Set(), byEmail = new Set(), byPhoneName = new Set();
  existing.forEach(r => {
    if (recruitingText_(r[44])) byKey.add(recruitingText_(r[44]));
    if (recruitingText_(r[7])) byEmail.add(recruitingNorm_(r[7]));
    if (recruitingText_(r[6]) && recruitingText_(r[2])) byPhoneName.add(recruitingNorm_(r[2])+'|'+recruitingPhone_(r[6]));
  });
  const candidates = [], duplicates = [], invalid = [];
  cfg.sources.forEach(src => {
    const sh = SpreadsheetApp.openById(src.id).getSheetByName(src.tab);
    if (!sh) throw new Error('Brak zakładki źródła ' + src.tab);
    const table = sh.getDataRange().getValues();
    if (table.length < 2) return;
    const h = recruitingHeaders_(table[0]);
    ['SUBMISSION ID','SURNAME','NAME','EMAIL','PHONE_NUMBER','SUBMISSION DATE'].forEach(k => {
      if (!(k in h)) throw new Error('Brak kolumny: '+k+' w '+src.label);
    });
    const get = (r,k) => h[k] == null ? '' : r[h[k]];
    table.slice(1).forEach((r,i) => {
      const name = [get(r,'NAME'),get(r,'SURNAME')].map(recruitingText_).filter(Boolean).join(' ');
      const email = recruitingNorm_(get(r,'EMAIL'));
      const phone = recruitingPhone_(get(r,'PHONE_NUMBER'));
      const sid = recruitingText_(get(r,'SUBMISSION ID'));
      const key = src.label + '|' + (sid || 'ROW:'+(i+2));
      if (!name || (!email && !phone)) { invalid.push(i+2); return; }
      const np = recruitingNorm_(name)+'|'+phone;
      if (byKey.has(key) || (email && byEmail.has(email)) || (phone && byPhoneName.has(np))) {
        duplicates.push({source:src.label, row:i+2}); return;
      }
      byKey.add(key); if (email) byEmail.add(email); if (phone) byPhoneName.add(np);
      const out = Array(45).fill('');
      out[0]=src.profile; out[1]=recruitingDate_(get(r,'SUBMISSION DATE'));
      out[2]=name; out[4]=recruitingDate_(get(r,'WHEN_DEPARTURE'));
      out[5]=get(r,'CANDIDATE_AGE'); out[6]=recruitingText_(get(r,'PHONE_NUMBER'));
      out[7]=recruitingText_(get(r,'EMAIL')); out[8]=recruitingYN_(get(r,'HAS_WHATSAPP'));
      out[9]=get(r,'CITY'); out[10]=recruitingYN_(get(r,'NOW_IN_NETHERLANDS'));
      out[11]=recruitingYN_(get(r,'HAS_COMPANY'));out[12]=recruitingYN_(get(r,'HAS_A1'));
      out[13]=recruitingYN_(get(r,'HAS_BSN'));out[14]=recruitingYN_(get(r,'HAS_DRIVING_LICENSE'));
      out[15]=recruitingYN_(get(r,'HAS_VCA'));out[16]=recruitingYN_(get(r,'HAS_EXPERIENCE_AS_ROOFER'));
      out[17]=get(r,'EXPERIENCE_AS_ROOFER_HOW_LONG');out[18]='Dekarz';
      out[19]=recruitingYN_(get(r,'HAS_EXPERIENCE_WITH_BITUMINAL'));
      out[20]=get(r,'EXP_SURFACE_BURNING_BITUMINAL');
      out[21]=recruitingYN_(get(r,'EXPERIENCE_WITH_PVC'));
      out[22]=get(r,'EXP_SURFACE_BURNING_PVC');
      out[23]=recruitingYN_(get(r,'WORKED_BENELUX'));out[24]=recruitingYN_(get(r,'WORKED_AS_LEADER'));
      out[25]=recruitingYN_(get(r,'ENGLISH_SKILL'))==='TAK' ? get(r,'ENGLISH_SKILL_LEVEL') : '—';
      out[26]=recruitingYN_(get(r,'DUTCH_SKILL'))==='TAK' ? get(r,'DUTCH_SKILL_LEVEL') : '—';
      out[29]=get(r,'FINAL_EVALUATION');out[33]='🆕';out[36]=src.label;out[37]='CF-ND-'+(sid || String(i+2));out[44]=key;
      candidates.push({out:out, date:out[1] instanceof Date ? out[1].getTime():0});
    });
  });
  candidates.sort((a,b)=>b.date-a.date);
  return { target:target, candidates:candidates, duplicates:duplicates.length, invalid:invalid.length, previous:existing.length };
}
function previewRecruitingSync() {
  const s=recruitingRead_();
  const summary={existing:s.previous, newCandidates:s.candidates.length, skippedDuplicates:s.duplicates, invalid:s.invalid, nextBatch:Math.min(RECRUITING_BATCH_LIMIT,s.candidates.length)};
  Logger.log(JSON.stringify(summary)); return summary;
}
function importRecruitingBatch() {
  const lock=LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Inna synchronizacja trwa');
  try {
    const s=recruitingRead_();
    const batch=s.candidates.slice(0,RECRUITING_BATCH_LIMIT).map(x=>x.out);
    if (!batch.length) { Logger.log('Brak nowych rekordów');return 0; }
    const sheet=s.target;
    const start=sheet.getLastRow()+1;
    if (sheet.getMaxColumns()<45) sheet.insertColumnsAfter(sheet.getMaxColumns(),45-sheet.getMaxColumns());
    if (sheet.getMaxRows()<start+batch.length-1) sheet.insertRowsAfter(sheet.getMaxRows(),start+batch.length-1-sheet.getMaxRows());
    // AG i AM są zarezerwowane dla istniejących ARRAYFORMULA; nie zapisujemy tam nawet pustych wartości.\n    sheet.getRange(start,1,batch.length,32).setValues(batch.map(r=>r.slice(0,32)));\n    sheet.getRange(start,34,batch.length,5).setValues(batch.map(r=>r.slice(33,38)));\n    sheet.getRange(start,40,batch.length,6).setValues(batch.map(r=>r.slice(39,45)));
    sheet.getRange(start,2,batch.length,1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(start,5,batch.length,1).setNumberFormat('dd.MM.yyyy');
    Logger.log('Dodano '+batch.length+' rekordów. Istniejących nie zmieniono.');
    return batch.length;
  } finally { lock.releaseLock(); }
}
function installRecruitingTrigger() {
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='importRecruitingBatch').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importRecruitingBatch').timeBased().everyMinutes(15).create();
}
function removeRecruitingTrigger() {
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='importRecruitingBatch').forEach(t => ScriptApp.deleteTrigger(t));
}
