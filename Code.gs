// ═══════════════════════════════════════════════════════════════
// GESTION DES PRÉSENCES — Google Apps Script OPTIMISÉ
// ═══════════════════════════════════════════════════════════════
//
// Optimisations principales :
//  • bootstrap() : connexion + config + personnes + appels en 1 requête
//  • CacheService pour les données de lecture
//  • 1 seul openById() par requête
//  • Appels stockés 1 ligne = 1 feuille + 1 date + JSON
//  • écritures groupées avec setValues()
//  • suppression groupée / sans deleteRow en boucle
//  • getSeances hors chemin critique de connexion
//
// IMPORTANT : le script accepte encore les anciennes lignes Appels
// (ancienne structure à 6 colonnes) et les convertit automatiquement
// vers la nouvelle structure au prochain enregistrement.
// ═══════════════════════════════════════════════════════════════

var SHEET_ID = '1zWdgb3VbzK7htICZxp6tpP0v86uV3GipOOy3O_TgZao';

var T = {
  FEUILLES:  'Feuilles',
  CONFIG:    'Config',
  PERSONNES: 'Personnes',
  APPELS:    'Appels',
  SEANCES:   'Seances'
};

var CACHE_TTL = 120; // secondes
var SEANCE_MAX = 45000;

// ─────────────────────────────────────────────────────────────
// HTTP
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  return route(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  var p = {};
  try {
    var ct = (e.postData && e.postData.type) ? e.postData.type : '';
    if (ct.indexOf('application/json') !== -1) {
      p = JSON.parse(e.postData.contents || '{}');
    } else {
      p = e.parameter || {};
    }
  } catch (err) {
    p = e.parameter || {};
  }
  return route(p);
}

function route(p) {
  try {
    var data = parseData(p.data);
    var a = String(p.action || '');

    switch (a) {
      case 'bootstrap':     return out(bootstrap(String(p.id || ''), String(p.pwd || '')));
      case 'getFeuille':    return out(getFeuille(p.id));
      case 'createFeuille': return out(createFeuille(p.id, p.name, p.pwd));
      case 'deleteFeuille': return out(deleteFeuille(p.id));
      case 'getConfig':     return out(getConfig(p.id));
      case 'saveConfig':    return out(saveConfig(p.id, data));
      case 'getPersonnes':  return out(getPersonnes(p.id));
      case 'savePersonnes': return out(savePersonnes(p.id, data));
      case 'getAppels':     return out(getAppels(p.id));
      case 'saveAppel':     return out(saveAppel(p.id, p.date, data));
      case 'saveAppels':    return out(saveAppels(p.id, data));
      case 'resetAppels':   return out(resetAppels(p.id));
      case 'getSeances':    return out(getSeances(p.id));
      case 'saveSeance':    return out(saveSeance(p.id, p.date, data));
      default:              return out({ok:false, error:'Action inconnue: ' + a});
    }
  } catch (err) {
    return out({ok:false, error: err && err.message ? err.message : String(err)});
  }
}

function parseData(v) {
  if (v === undefined || v === null || v === '') return undefined;
  if (typeof v !== 'string') return v;
  try { return JSON.parse(v); } catch (e) { return {}; }
}

// ─────────────────────────────────────────────────────────────
// BOOTSTRAP — chemin critique de connexion
// ─────────────────────────────────────────────────────────────
function bootstrap(id, pwd) {
  if (!id) return {ok:false, error:'id manquant'};

  // Les petites métadonnées sont très souvent demandées.
  var meta = readFeuilleCached(id);
  if (!meta) return {ok:false, error:'Feuille introuvable'};
  if (String(meta.pwd) !== String(pwd)) {
    return {ok:false, error:'Mot de passe incorrect'};
  }

  var ss = openSS();
  var cfg = readConfigFromSheet(ss, id);
  var persons = readPersonnesFromSheet(ss, id);
  var records = readAppelsFromSheet(ss, id);

  // On met en cache séparément pour que les refreshs suivants soient rapides.
  cachePut('cfg:' + id, cfg);
  cachePut('persons:' + id, persons);
  cachePut('calls:' + id, records);

  return {
    ok:true,
    data:{
      feuille:{id:meta.id, name:meta.name, pwd:meta.pwd},
      config:cfg,
      persons:persons,
      records:records
    }
  };
}

// ─────────────────────────────────────────────────────────────
// FEUILLES
// ─────────────────────────────────────────────────────────────
function getFeuille(id) {
  var cached = readFeuilleCached(String(id || ''));
  return {ok:true, data:cached};
}

function readFeuilleCached(id) {
  if (!id) return null;
  var key = 'meta:' + id;
  var cached = cacheGet(key);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var ss = openSS();
  var sh = getTab(ss, T.FEUILLES);
  var last = sh.getLastRow();
  if (last < 2) return null;

  var rows = sh.getRange(2,1,last-1,4).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]) === id) {
      var meta = {id:rows[i][0], name:rows[i][1], pwd:rows[i][2]};
      cachePut(key, meta);
      return meta;
    }
  }
  return null;
}

function createFeuille(id, name, pwd) {
  id = String(id || '').trim();
  if (!id || !name) return {ok:false,error:'Nom ou id manquant'};

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = openSS();
    var sh = getTab(ss, T.FEUILLES);
    var last = sh.getLastRow();
    if (last >= 2) {
      var ids = sh.getRange(2,1,last-1,1).getValues();
      for (var i=0; i<ids.length; i++) {
        if (String(ids[i][0]) === id) return {ok:true};
      }
    }
    sh.getRange(sh.getLastRow()+1,1,1,4).setValues([[id,name,pwd,new Date().toISOString()]]);
    cachePut('meta:' + id, {id:id,name:name,pwd:pwd});
    invalidate(id);
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

function deleteFeuille(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = openSS();
    deleteById(ss, T.FEUILLES, id, 0);
    deleteById(ss, T.CONFIG, id, 0);
    deleteById(ss, T.PERSONNES, id, 0);
    deleteById(ss, T.APPELS, id, 0);
    deleteById(ss, T.SEANCES, id, 0);
    invalidate(id);
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────
function getConfig(id) {
  var key = 'cfg:' + id;
  var cached = cacheGet(key);
  if (cached) {
    try { return {ok:true,data:JSON.parse(cached)}; } catch(e) {}
  }
  var ss = openSS();
  var cfg = readConfigFromSheet(ss, String(id || ''));
  cachePut(key, cfg);
  return {ok:true,data:cfg};
}

function readConfigFromSheet(ss, id) {
  var sh = getTab(ss, T.CONFIG);
  var last = sh.getLastRow();
  if (last < 2) return null;

  var rows = sh.getRange(2,1,last-1,7).getValues();
  for (var i=0; i<rows.length; i++) {
    if (String(rows[i][0]) !== id) continue;
    return {
      days: rows[i][1] ? String(rows[i][1]).split(',').map(Number).filter(function(n){return !isNaN(n);}) : [],
      timeStart: fmtTime(rows[i][2]),
      timeEnd: fmtTime(rows[i][3]),
      yearStart: fmtDate(rows[i][4]),
      yearEnd: fmtDate(rows[i][5]),
      vacances: rows[i][6] ? safeJson(rows[i][6], []) : []
    };
  }
  return null;
}

function saveConfig(id, cfg) {
  if (!cfg) return {ok:false,error:'cfg manquant'};
  id = String(id || '');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = openSS();
    var sh = getTab(ss, T.CONFIG);
    var row = [
      id,
      (cfg.days || []).join(','),
      cfg.timeStart || '18:00',
      cfg.timeEnd || '20:00',
      cfg.yearStart || '',
      cfg.yearEnd || '',
      JSON.stringify(cfg.vacances || [])
    ];
    var last = sh.getLastRow();
    var target = 0;
    if (last >= 2) {
      var ids = sh.getRange(2,1,last-1,1).getValues();
      for (var i=0;i<ids.length;i++) {
        if (String(ids[i][0]) === id) { target=i+2; break; }
      }
    }
    if (!target) target=last+1;
    sh.getRange(target,1,1,row.length).setValues([row]);
    cachePut('cfg:' + id, normalizeConfig(cfg));
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────
// PERSONNES
// ─────────────────────────────────────────────────────────────
function getPersonnes(id) {
  var key = 'persons:' + id;
  var cached = cacheGet(key);
  if (cached) {
    try { return {ok:true,data:JSON.parse(cached)}; } catch(e) {}
  }
  var ss = openSS();
  var persons = readPersonnesFromSheet(ss, String(id || ''));
  cachePut(key, persons);
  return {ok:true,data:persons};
}

function readPersonnesFromSheet(ss, id) {
  var sh = getTab(ss, T.PERSONNES);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2,1,last-1,5).getValues();
  var persons=[];
  for (var i=0;i<rows.length;i++) {
    if (String(rows[i][0]) !== id) continue;
    persons.push({
      id:String(rows[i][1]),
      name:String(rows[i][2]),
      startDate:fmtDate(rows[i][3]),
      active:rows[i][4] !== false && rows[i][4] !== 'false'
    });
  }
  return persons;
}

function savePersonnes(id, persons) {
  if (!Array.isArray(persons)) return {ok:false,error:'persons manquant'};
  id=String(id||'');

  var lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss=openSS();
    var sh=getTab(ss,T.PERSONNES);
    var last=sh.getLastRow();
    var old=[];
    if(last>=2) old=sh.getRange(2,1,last-1,5).getValues();

    var keep=[];
    for(var i=0;i<old.length;i++) {
      if(String(old[i][0])!==id) keep.push(old[i]);
    }

    var add=persons.map(function(p){
      return [id,p.id,p.name,p.startDate || '',p.active !== false ? 'true':'false'];
    });
    var all=keep.concat(add);

    replaceDataRows(sh,5,all);
    cachePut('persons:'+id,persons);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────
// APPELS — NOUVELLE STRUCTURE : 1 ligne = 1 feuille/date
// Colonnes : feuille_id | date | data_json | updatedAt
// ─────────────────────────────────────────────────────────────
function getAppels(id) {
  var key='calls:'+id;
  var cached=cacheGet(key);
  if(cached){try{return {ok:true,data:JSON.parse(cached)};}catch(e){}}

  var ss=openSS();
  var records=readAppelsFromSheet(ss,String(id||''));
  cachePut(key,records);
  return {ok:true,data:records};
}

function readAppelsFromSheet(ss,id) {
  var sh=getTab(ss,T.APPELS);
  var last=sh.getLastRow();
  if(last<2)return {};

  var width=Math.max(sh.getLastColumn(),6);
  var rows=sh.getRange(2,1,last-1,width).getValues();
  var records={};

  // Nouvelle structure
  if(width>=4 && isNewAppelsHeader(sh)) {
    for(var i=0;i<rows.length;i++) {
      if(String(rows[i][0])!==id)continue;
      var d=fmtDate(rows[i][1]);if(!d)continue;
      var rec=safeJson(rows[i][2],{});
      if(rec && typeof rec==='object')records[d]=rec;
    }
    return records;
  }

  // Ancienne structure : feuille_id/date/person_id/statut/validated/cancelled
  for(var j=0;j<rows.length;j++) {
    if(String(rows[j][0])!==id)continue;
    var date=fmtDate(rows[j][1]);if(!date)continue;
    var rec2=records[date]||(records[date]={});
    var pid=String(rows[j][2]||'').trim();
    var stat=String(rows[j][3]||'').trim();
    if(pid&&stat)rec2[pid]=stat;
    if(String(rows[j][4]||'').trim()==='true')rec2.__validated=true;
    if(String(rows[j][5]||'').trim()==='true')rec2.__cancelled=true;
  }
  return records;
}

function saveAppel(id,date,rec) {
  if(!id)return {ok:false,error:'id manquant'};
  date=fmtDate(date);
  if(!date)return {ok:false,error:'date manquante'};
  if(!rec||typeof rec!=='object')return {ok:false,error:'rec manquant — data non reçu'};

  var lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss=openSS();
    ensureAppelsSchema(ss);
    var sh=getTab(ss,T.APPELS);
    var last=sh.getLastRow();
    var rows=last>=2?sh.getRange(2,1,last-1,4).getValues():[];
    var target=0;

    for(var i=0;i<rows.length;i++) {
      if(String(rows[i][0])===String(id) && fmtDate(rows[i][1])===date){target=i+2;break;}
    }

    var clean=normalizeRecord(rec);
    var now=new Date().toISOString();
    if(!Object.keys(clean).length){
      if(target)sh.deleteRow(target);
    } else {
      var values=[[id,date,JSON.stringify(clean),now]];
      if(target)sh.getRange(target,1,1,4).setValues(values);
      else sh.getRange(sh.getLastRow()+1,1,1,4).setValues(values);
    }

    var all=getAppelsFreshFromSheet(ss,id);
    cachePut('calls:'+id,all);
    return {ok:true};
  } finally {lock.releaseLock();}
}

function saveAppels(id,recs) {
  if(!id)return {ok:false,error:'id manquant'};
  if(!recs||typeof recs!=='object')return {ok:false,error:'recs manquant — data non reçu'};

  var dates=Object.keys(recs).map(fmtDate).filter(Boolean);
  if(!dates.length)return {ok:true,data:{saved:0,rows:0}};

  var lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss=openSS();
    ensureAppelsSchema(ss);
    var sh=getTab(ss,T.APPELS);
    var last=sh.getLastRow();
    var rows=last>=2?sh.getRange(2,1,last-1,4).getValues():[];
    var wanted={};dates.forEach(function(d){wanted[d]=true;});
    var keep=[];

    for(var i=0;i<rows.length;i++) {
      if(String(rows[i][0])===String(id) && wanted[fmtDate(rows[i][1])])continue;
      keep.push(rows[i]);
    }

    var add=[];
    dates.forEach(function(d){
      var clean=normalizeRecord(recs[d]||{});
      if(Object.keys(clean).length)add.push([id,d,JSON.stringify(clean),new Date().toISOString()]);
    });

    replaceDataRows(sh,4,keep.concat(add));
    var all=getAppelsFreshFromSheet(ss,id);
    cachePut('calls:'+id,all);
    return {ok:true,data:{saved:dates.length,rows:add.length}};
  } finally {lock.releaseLock();}
}

function getAppelsFreshFromSheet(ss,id){
  return readAppelsFromSheet(ss,id);
}

function resetAppels(id) {
  var lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss=openSS();
    var sh=getTab(ss,T.APPELS);
    var last=sh.getLastRow();
    if(last>=2){
      var rows=sh.getRange(2,1,last-1,Math.max(4,sh.getLastColumn())).getValues();
      var keep=[];
      for(var i=0;i<rows.length;i++)if(String(rows[i][0])!==String(id))keep.push(rows[i]);
      replaceDataRows(sh,Math.max(4,sh.getLastColumn()),keep);
    }
    cachePut('calls:'+id,{});
    return {ok:true};
  } finally {lock.releaseLock();}
}

// ─────────────────────────────────────────────────────────────
// SÉANCES
// ─────────────────────────────────────────────────────────────
function getSeances(id) {
  var key='sea:'+id;
  var cached=cacheGet(key);
  if(cached){try{return {ok:true,data:JSON.parse(cached)};}catch(e){}}

  var ss=openSS();
  var sh=getTab(ss,T.SEANCES);
  var last=sh.getLastRow();
  var res={};
  if(last>=2){
    var rows=sh.getRange(2,1,last-1,4).getValues();
    for(var i=0;i<rows.length;i++){
      if(String(rows[i][0])!==String(id))continue;
      var d=fmtDate(rows[i][1]);
      if(!d||!rows[i][2])continue;
      res[d]={html:String(rows[i][2]),updatedAt:fmtStamp(rows[i][3])};
    }
  }
  cachePut(key,res);
  return {ok:true,data:res};
}

function saveSeance(id,date,rec) {
  if(!id)return {ok:false,error:'id manquant'};
  date=fmtDate(date);
  if(!date)return {ok:false,error:'date manquante'};
  if(!rec||typeof rec.html!=='string')return {ok:false,error:'html manquant — data non reçu'};
  var html=rec.html.trim();
  if(html.length>SEANCE_MAX)return {ok:false,error:'Contenu trop long ('+html.length+' / '+SEANCE_MAX+' caractères)'};

  var lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss=openSS();
    var sh=getTab(ss,T.SEANCES);
    var last=sh.getLastRow();
    var target=0;
    if(last>=2){
      var rows=sh.getRange(2,1,last-1,4).getValues();
      for(var i=0;i<rows.length;i++){
        if(String(rows[i][0])===String(id)&&fmtDate(rows[i][1])===date){target=i+2;break;}
      }
    }

    if(!html){
      if(target)sh.deleteRow(target);
      invalidateSeance(id);
      return {ok:true,data:{html:'',updatedAt:''}};
    }

    var stamp=new Date().toISOString();
    if(target)sh.getRange(target,1,1,4).setValues([[id,date,html,stamp]]);
    else sh.getRange(sh.getLastRow()+1,1,1,4).setValues([[id,date,html,stamp]]);
    cacheDelete('sea:'+id);
    return {ok:true,data:{html:html,updatedAt:stamp}};
  } finally {lock.releaseLock();}
}

// ─────────────────────────────────────────────────────────────
// SCHEMA / MIGRATION APPELS
// ─────────────────────────────────────────────────────────────
function ensureAppelsSchema(ss) {
  var sh=getTab(ss,T.APPELS);
  var headers=sh.getRange(1,1,1,Math.max(4,sh.getLastColumn())).getValues()[0];
  if(String(headers[0])==='feuille_id' && String(headers[1])==='date' && String(headers[2])==='data_json')return;

  var last=sh.getLastRow();
  var old=last>=2?sh.getRange(2,1,last-1,Math.max(6,sh.getLastColumn())).getValues():[];
  var grouped={};
  for(var i=0;i<old.length;i++){
    var id=String(old[i][0]||'');
    var d=fmtDate(old[i][1]);
    if(!id||!d)continue;
    var key=id+'\u0000'+d;
    var rec=grouped[key]||(grouped[key]={id:id,date:d,data:{}});
    var pid=String(old[i][2]||'').trim();
    var stat=String(old[i][3]||'').trim();
    if(pid&&stat)rec.data[pid]=stat;
    if(String(old[i][4]||'').trim()==='true')rec.data.__validated=true;
    if(String(old[i][5]||'').trim()==='true')rec.data.__cancelled=true;
  }

  var data=[];
  Object.keys(grouped).forEach(function(k){
    var x=grouped[k];data.push([x.id,x.date,JSON.stringify(x.data),new Date().toISOString()]);
  });

  sh.clearContents();
  sh.getRange(1,1,1,4).setValues([['feuille_id','date','data_json','updatedAt']]);
  if(data.length)sh.getRange(2,1,data.length,4).setValues(data);
}

function isNewAppelsHeader(sh) {
  var h=sh.getRange(1,1,1,Math.max(4,sh.getLastColumn())).getValues()[0];
  return String(h[0])==='feuille_id' && String(h[1])==='date' && String(h[2])==='data_json';
}

function normalizeRecord(rec) {
  var out={};
  Object.keys(rec||{}).forEach(function(k){
    if(k==='__validated'||k==='__cancelled') {
      if(rec[k]===true||rec[k]==='true')out[k]=true;
    } else if(rec[k]!==undefined && rec[k]!==null && String(rec[k])!=='') {
      out[k]=String(rec[k]);
    }
  });
  return out;
}

function normalizeConfig(cfg) {
  cfg=cfg||{};
  return {
    days:Array.isArray(cfg.days)?cfg.days:[],
    timeStart:cfg.timeStart||'18:00',
    timeEnd:cfg.timeEnd||'20:00',
    yearStart:cfg.yearStart||'',
    yearEnd:cfg.yearEnd||'',
    vacances:Array.isArray(cfg.vacances)?cfg.vacances:[]
  };
}

// ─────────────────────────────────────────────────────────────
// SHEETS / UTILITAIRES
// ─────────────────────────────────────────────────────────────
function openSS(){return SpreadsheetApp.openById(SHEET_ID);}

function getTab(ss,name){
  var sh=ss.getSheetByName(name);
  if(!sh){
    sh=ss.insertSheet(name);
    var headers={
      Feuilles:['id','name','pwd','createdAt'],
      Config:['feuille_id','days','timeStart','timeEnd','yearStart','yearEnd','vacances'],
      Personnes:['feuille_id','person_id','name','startDate','active'],
      Appels:['feuille_id','date','data_json','updatedAt'],
      Seances:['feuille_id','date','contenu','updatedAt']
    };
    sh.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
  }
  return sh;
}

function replaceDataRows(sh,width,rows){
  var oldLast=sh.getLastRow();
  if(oldLast>1)sh.getRange(2,1,oldLast-1,width).clearContent();
  if(rows.length)sh.getRange(2,1,rows.length,width).setValues(rows);
}

function deleteById(ss,name,id,col){
  var sh=getTab(ss,name),last=sh.getLastRow();
  if(last<2)return;
  var width=Math.max(sh.getLastColumn(),col+1);
  var rows=sh.getRange(2,1,last-1,width).getValues();
  var keep=[];
  for(var i=0;i<rows.length;i++)if(String(rows[i][col])!==String(id))keep.push(rows[i]);
  replaceDataRows(sh,width,keep);
}

// ─────────────────────────────────────────────────────────────
// CACHE
// ─────────────────────────────────────────────────────────────
function cacheGet(key){
  try{return CacheService.getScriptCache().get(key);}catch(e){return null;}
}

function cachePut(key,value){
  try{
    var s=typeof value==='string'?value:JSON.stringify(value);
    // CacheService limite une entrée. Au-delà, on ne met simplement pas en cache.
    if(s.length<90000)CacheService.getScriptCache().put(key,s,CACHE_TTL);
  }catch(e){}
}

function cacheDelete(key){try{CacheService.getScriptCache().remove(key);}catch(e){}}

function invalidate(id){
  ['meta:','cfg:','persons:','calls:','sea:'].forEach(function(p){cacheDelete(p+id);});
}
function invalidateSeance(id){cacheDelete('sea:'+id);}

// ─────────────────────────────────────────────────────────────
// FORMATAGE
// ─────────────────────────────────────────────────────────────
function safeJson(v,def){
  try{return JSON.parse(String(v));}catch(e){return def;}
}

function fmtStamp(val){
  if(!val)return '';
  return val instanceof Date?val.toISOString():String(val);
}

function fmtTime(val){
  if(!val)return '18:00';
  if(val instanceof Date){
    return String(val.getHours()).padStart(2,'0')+':'+String(val.getMinutes()).padStart(2,'0');
  }
  var s=String(val).trim();
  var m=s.match(/(\d{1,2}):(\d{2})/);
  return m?m[1].padStart(2,'0')+':'+m[2]:'18:00';
}

function fmtDate(val){
  if(!val)return '';
  if(val instanceof Date){
    if(isNaN(val.getTime()))return '';
    return val.getFullYear()+'-'+String(val.getMonth()+1).padStart(2,'0')+'-'+String(val.getDate()).padStart(2,'0');
  }
  var s=String(val).trim();if(!s)return '';
  if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);
  var mFR=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(mFR)return mFR[3]+'-'+mFR[2].padStart(2,'0')+'-'+mFR[1].padStart(2,'0');
  var ts=new Date(s);
  if(!isNaN(ts.getTime()))return ts.getFullYear()+'-'+String(ts.getMonth()+1).padStart(2,'0')+'-'+String(ts.getDate()).padStart(2,'0');
  return s.slice(0,10);
}

function out(data){
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
