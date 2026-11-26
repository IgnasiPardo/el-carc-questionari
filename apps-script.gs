// El Carç — servei de desament (Google Apps Script).
// Instal·lar dins del full privat: Extensions > Apps Script.
// Executar configura() una vegada, i desplegar com a aplicació web.
// No incloure mai els codis individuals o l'identificador del full al repositori públic.
const WEB_URL = 'https://ignasipardo.github.io/el-carc-questionari/';
const AXIS_COUNTS = [3,4,4,8,3,3,4,3,3,1,4,2];
const PROPERTIES = () => PropertiesService.getScriptProperties();

function configura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || !ss.getSheetByName('Respostes') || !ss.getSheetByName('Accessos')) {
    throw new Error('Executa configura des del projecte Apps Script vinculat al full de respostes d’El Carç.');
  }
  const p = PROPERTIES();
  if (p.getProperty('SHEET_ID') && p.getProperty('SHEET_ID') !== ss.getId()) throw new Error('El projecte ja està vinculat a un altre full.');
  p.setProperty('SHEET_ID', ss.getId());
  const rows = [];
  for (let i=1;i<=3;i++) {
    const key='SOCI_'+i;
    let token=p.getProperty(key);
    if (!token) {
      token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
      p.setProperty(key,token);
    }
    rows.push(['Soci '+i,WEB_URL+'#codi='+token,'Preparat']);
  }
  ss.getSheetByName('Accessos').getRange(2,1,3,3).setValues(rows);
  SpreadsheetApp.flush();
}

function doGet() {
  return page_('Servei de recepció del qüestionari d’El Carç.');
}

function cleanText_(value, limit) {
  if(typeof value !== 'string' || value.length > limit) throw new Error('Text no vàlid');
  return /^\s*[=+\-@]/.test(value) ? "'" + value : value;
}

function score_(value, max) {
  if(value === '?') return 'NS';
  if(!Number.isInteger(value) || value<1 || value>max) throw new Error('Valor no vàlid');
  return value;
}

function verifyPayload_(data) {
  if(!data || !['4','5'].includes(data.version) || typeof data.submissionId !== 'string' ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(data.submissionId)) throw new Error('Enviament no vàlid');
  if(typeof data.participant !== 'string' || !data.participant.trim()) throw new Error('Falta el nom');
  if(!Array.isArray(data.answers) || data.answers.length!==AXIS_COUNTS.length) throw new Error('Eixos incomplets');
  const values=[];
  data.answers.forEach((a,i)=>{
    if(!a || !Array.isArray(a.attention) || a.attention.length!==AXIS_COUNTS[i]) throw new Error('Qüestions incompletes');
    values.push(score_(a.crit,5),score_(a.cap,5));
    a.attention.forEach(v=>values.push(score_(v,3)));
  });
  if(values.length!==66) throw new Error('Valoracions incompletes');
  return {version:data.version,id:data.submissionId,participant:cleanText_(data.participant.trim(),200),values,notes:cleanText_(data.notes||'',10000)};
}

function doPost(e) {
  let lock;
  try {
    const p=PROPERTIES(), token=e && e.parameter && e.parameter.codi;
    if(typeof token!=='string' || !/^[0-9a-f]{64}$/i.test(token)) throw new Error('Accés no vàlid');
    let soci='';
    for(let i=1;i<=3;i++) if(token===p.getProperty('SOCI_'+i)) soci='Soci '+i;
    if(!soci) throw new Error('Accés no vàlid');
    const raw=e.parameter.payload;
    if(typeof raw!=='string' || raw.length>40000) throw new Error('Enviament no vàlid');
    const data=verifyPayload_(JSON.parse(raw));
    const sheetId=p.getProperty('SHEET_ID');if(!sheetId)throw new Error('Servei no configurat');
    lock=LockService.getScriptLock();lock.waitLock(15000);
    const sheet=SpreadsheetApp.openById(sheetId).getSheetByName('Respostes');
    if(!sheet || sheet.getLastColumn()!==72 || sheet.getRange(1,1).getValue()!=='ID enviament') throw new Error('Estructura del full incorrecta');
    const last=sheet.getLastRow();
    const previous=last>1?sheet.getRange(2,1,last-1,5).getValues():[];
    const match=previous.find(row=>row[0]===data.id);
    if(match) {
      if(match[4]!==soci) throw new Error('Identificador no vàlid');
      return page_('Han quedat desades les teves respostes.');
    }
    const row=[data.id,new Date(),data.participant,data.version,soci,...data.values,data.notes];
    const dest=last+1;
    sheet.getRange(dest,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();
    if(sheet.getRange(dest,1).getValue()!==data.id)throw new Error('No s’ha confirmat el desament');
    return page_('Han quedat desades les teves respostes.');
  } catch(err) {
    return page_('No s’han pogut confirmar les respostes. Torna al qüestionari i prova-ho de nou.');
  } finally {
    if(lock && lock.hasLock())lock.releaseLock();
  }
}

function page_(message) {
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ca"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>El Carç</title><style>body{margin:0;background:#f0f1eb;color:#242d26;font:16px/1.5 system-ui,sans-serif;display:grid;place-items:center;min-height:100vh}main{box-sizing:border-box;width:min(90%,560px);padding:36px 24px;background:#faf9f5;border:1px solid #d9dfd5;border-radius:14px;text-align:center}h1{font:normal clamp(22px,5vw,30px)/1.35 Georgia,serif;margin:0}</style></head><body><main><h1>'+message+'</h1></main></body></html>').setTitle('El Carç').addMetaTag('viewport','width=device-width, initial-scale=1');
}
