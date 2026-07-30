const XLSX = require('xlsx');

const REGION_ALIASES = {
  'Thies':'Thiès', 'Thiés':'Thiès', 'Thiès':'Thiès',
  'Tamba':'Tambacounda', 'Tambacounda':'Tambacounda',
  'Saint Louis':'Saint-Louis', 'Saint Louis ':'Saint-Louis', 'Saint-Louis':'Saint-Louis',
  'BAKEL':'Tambacounda',
};
function normalizeRegion(raw){
  const t = (raw==null?'':String(raw)).trim();
  return REGION_ALIASES[t] || t;
}
function linreg(points){
  const n=points.length;
  const sx=points.reduce((a,p,i)=>a+i,0);
  const sy=points.reduce((a,p)=>a+p,0);
  const sxy=points.reduce((a,p,i)=>a+i*p,0);
  const sxx=points.reduce((a,p,i)=>a+i*i,0);
  const slope=(n*sxy-sx*sy)/(n*sxx-sx*sx);
  const intercept=(sy-slope*sx)/n;
  return {slope, intercept};
}
function round2(n){ return Math.round(n*100)/100; }

function getSheetAOA(wb, name){
  const ws = wb.Sheets[name];
  if(!ws) return null;
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
}

function buildPlaques(aoa){
  const headers = aoa[0];
  const idx = {
    plaque: headers.indexOf('PLAQUE'),
    zone: headers.indexOf('Zone DRV'),
    eql: headers.indexOf('EQL PLAQUES'),
    taux: headers.indexOf("TAUX D'OCCUPATION"),
    pbo: headers.indexOf('Taux de PBO saturé'),
    age: headers.indexOf('Age de la plaque en Mois'),
    commune: headers.indexOf('Commune'),
    arr: headers.indexOf('Arrondissement'),
    dept: headers.indexOf('Département'),
    region: headers.indexOf('Région'),
    dateOuverture: headers.indexOf('Date ouverture'),
  };
  const out = [];
  for(let i=1;i<aoa.length;i++){
    const r = aoa[i];
    if(!r || r[idx.plaque]==null) continue;
    const tauxRaw = Number(r[idx.taux]);
    const pboRaw = Number(r[idx.pbo]);
    out.push({
      plaque: String(r[idx.plaque]),
      zone: r[idx.zone]||'',
      eql: Number(r[idx.eql])||0,
      taux: isFinite(tauxRaw) ? round2(tauxRaw*100) : 0,
      pbo: isFinite(pboRaw) ? round2(pboRaw*100) : 0,
      age: Number(r[idx.age])||0,
      commune: r[idx.commune]||'',
      arr: r[idx.arr]||'',
      dept: r[idx.dept]||'',
      region: (r[idx.region]==null?'':String(r[idx.region])).trim(),
      openYear: excelSerialToYear(r[idx.dateOuverture]),
      openDateLabel: excelSerialToDateLabel(r[idx.dateOuverture]),
      openSerial: Number(r[idx.dateOuverture])||null,
    });
  }
  return out;
}

/* Un serial Excel (nombre de jours depuis 1899-12-30) -> année civile.
   cellDates:false lors du XLSX.read, donc les dates arrivent en nombre. */
function excelSerialToYear(serial){
  const n = Number(serial);
  if(!isFinite(n) || n<=0) return null;
  const ms = Date.UTC(1899,11,30) + n*86400000;
  return new Date(ms).getUTCFullYear();
}

function excelSerialToDateLabel(serial){
  const n = Number(serial);
  if(!isFinite(n) || n<=0) return '—';
  const d = new Date(Date.UTC(1899,11,30) + n*86400000);
  const p = x => String(x).padStart(2,'0');
  return `${p(d.getUTCDate())}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()}`;
}

function buildRegionNaming(plaques){
  const m = {};
  plaques.forEach(p=>{ const r = p.region || '(vide)'; m[r]=(m[r]||0)+1; });
  return m;
}

function buildSynthese(aoa){
  const headers = aoa[0];
  const idx = {region:headers.indexOf('Région'), commune:headers.indexOf('Commune'), eql:headers.indexOf('EQL'), racc:headers.indexOf('Raccordés')};
  const out=[];
  for(let i=1;i<aoa.length;i++){
    const r=aoa[i];
    if(!r || r[idx.commune]==null) continue;
    out.push({
      region: normalizeRegion(r[idx.region]),
      commune: r[idx.commune],
      eql: Number(r[idx.eql])||0,
      racc: Number(r[idx.racc])||0,
    });
  }
  return out;
}

function buildRegionsFromSynthese(syntheseRows, plaqueCountByRegion, pboAvgByRegion){
  const map = {};
  syntheseRows.forEach(row=>{
    if(!map[row.region]) map[row.region] = {region:row.region, eql:0, racc:0};
    map[row.region].eql += row.eql;
    map[row.region].racc += row.racc;
  });
  return Object.values(map).map(r=>({
    region: r.region,
    plaques: plaqueCountByRegion[r.region]||0,
    eql: r.eql,
    racc: r.racc,
    taux: r.eql ? round2(r.racc/r.eql*100) : 0,
    pbo_avg: pboAvgByRegion[r.region]!=null ? round2(pboAvgByRegion[r.region]) : 0,
  })).sort((a,b)=>b.eql-a.eql);
}

function buildTopBottom(syntheseRows){
  const withTaux = syntheseRows.filter(r=>r.eql>=200).map(r=>({
    region:r.region, commune:r.commune, eql:r.eql, racc:r.racc,
    taux: round2(r.racc/r.eql*100),
  }));
  const sorted = [...withTaux].sort((a,b)=>b.taux-a.taux);
  return {
    top10: sorted.slice(0,10),
    bottom10: sorted.slice(-10).sort((a,b)=>a.taux-b.taux),
  };
}

function normalizeWeekLabel(wk){
  const [w,y] = wk.slice(1).split('_');
  return 'S'+w.padStart(2,'0')+'_'+y;
}

function buildTrend(aoa){
  const headers = aoa[0];
  const idx = {semaine:headers.indexOf('Semaine'), eql:headers.indexOf('EQL'), racc:headers.indexOf('Clients Raccordés')};
  const byWeek = {};
  for(let i=1;i<aoa.length;i++){
    const r = aoa[i];
    if(!r || r[idx.semaine]==null) continue;
    const wk = String(r[idx.semaine]).trim();
    if(!/^S\d{1,2}_\d{4}$/.test(wk)) continue;
    if(!byWeek[wk]) byWeek[wk]={eql:0, racc:0};
    byWeek[wk].eql += Number(r[idx.eql])||0;
    byWeek[wk].racc += Number(r[idx.racc])||0;
  }
  let weeks = Object.keys(byWeek).sort((a,b)=>{
    const [aw,ay]=a.slice(1).split('_').map(Number);
    const [bw,by]=b.slice(1).split('_').map(Number);
    return (ay-by) || (aw-bw);
  });
  /* TO_Communes accumule l'historique sur plusieurs années civiles ; le tableau de
     bord ("Évolution du taux d'occupation") est conçu pour une seule année. On ne
     garde donc que les semaines de l'année la plus récente présente dans le fichier. */
  if(weeks.length){
    const latestYear = Number(weeks[weeks.length-1].split('_')[1]);
    weeks = weeks.filter(wk => Number(wk.split('_')[1]) === latestYear);
  }
  return weeks.map(wk=>({
    week: normalizeWeekLabel(wk),
    eql: byWeek[wk].eql,
    racc: byWeek[wk].racc,
    taux: byWeek[wk].eql ? round2(byWeek[wk].racc/byWeek[wk].eql*100) : 0,
  }));
}

function buildAnomalies(plaques, regionNaming){
  const pboSat = plaques.filter(p=>p.pbo>=70).sort((a,b)=>b.pbo-a.pbo).slice(0,40)
    .map(p=>({plaque:p.plaque, region:p.region, commune:p.commune, pbo:p.pbo, taux:p.taux}));
  const lowTaux = [...plaques].sort((a,b)=>a.taux-b.taux).slice(0,10)
    .map(p=>({plaque:p.plaque, region:p.region, commune:p.commune, taux:p.taux, eql:p.eql, age:p.age}));
  const stale = plaques.filter(p=>p.age>=24).sort((a,b)=>a.taux-b.taux).slice(0,50)
    .map(p=>({plaque:p.plaque, region:p.region, commune:p.commune, age:p.age, taux:p.taux, eql:p.eql}));
  return {pbo_sat:pboSat, low_taux:lowTaux, stale:stale, region_naming:regionNaming};
}

function buildHistogram(plaques){
  const labels = ["0-10","10-20","20-30","30-40","40-50","50-60","60-70","70-80","80-90","90-100","100+"];
  const values = new Array(11).fill(0);
  plaques.forEach(p=>{
    let idx = Math.floor(p.age/10);
    if(idx>10) idx=10;
    values[idx]+=1;
  });
  return {labels, values};
}

function buildGlobal(trend, plaquesCount){
  const cur = trend[trend.length-1] || {eql:0,racc:0,taux:0};
  const prev = trend[trend.length-2] || cur;
  return {
    eql_cur: cur.eql, racc_cur: cur.racc, taux_cur: cur.taux,
    eql_prev: prev.eql, racc_prev: prev.racc, taux_prev: prev.taux,
    delta_racc: cur.racc - prev.racc,
    delta_taux: round2(cur.taux - prev.taux),
    total_plaques: plaquesCount,
  };
}

/* ================== Section "Pénétration FTTH / Commune" ==================
   Réplique dynamique du rapport hebdomadaire (auparavant codé en dur).
   Sources vérifiées dans le classeur :
   - TO_Communes  : EQL + Clients Raccordés par commune ET par semaine (long format)
   - Clients Fictifs : "Raccordés sans constitution" = somme par semaine
   - Client       : "Total client" = somme(Nombre de ND) / 2 (confirmé)
   - TO_Plaques   : dates d'ouverture -> plaques déployées par année
   NB : "En cours de fiabilisation" n'a pas de colonne source identifiable dans
   le classeur ; ce chiffre est donc saisi manuellement et mémorisé semaine par
   semaine (voir editTotalClient / localStorage plus bas). */

function buildCommuneWeekMap(communesAOA){
  const headers = communesAOA[0];
  const idx = {commune:headers.indexOf('Commune'), drv:headers.indexOf('Zone DRV'), semaine:headers.indexOf('Semaine'), eql:headers.indexOf('EQL'), racc:headers.indexOf('Clients Raccordés')};
  const map = {};
  for(let i=1;i<communesAOA.length;i++){
    const r = communesAOA[i];
    if(!r || r[idx.commune]==null || r[idx.semaine]==null) continue;
    const wkRaw = String(r[idx.semaine]).trim();
    if(!/^S\d{1,2}_\d{4}$/.test(wkRaw)) continue;
    const wk = normalizeWeekLabel(wkRaw);
    if(!map[wk]) map[wk] = {};
    map[wk][r[idx.commune]] = {
      eql: Number(r[idx.eql])||0,
      racc: Number(r[idx.racc])||0,
      drv: r[idx.drv]||'',
    };
  }
  return map;
}

/* ================== Section "Analyse Statistique Avancée" ==================
   Le classeur contient en réalité ~3 ans d'historique (S01_2024 -> aujourd'hui)
   que le reste du dashboard n'exploite quasiment pas (juste une régression sur
   10 semaines pour la prévision). Ici on exploite tout l'historique pour :
   - détecter les semaines nationales statistiquement anormales (z-score sur
     les variations hebdo, pas un seuil arbitraire) ;
   - comparer chaque commune à SA PROPRE trajectoire historique (et pas à un
     seuil fixe type "< 20%") pour repérer les vrais décrocheurs/accélérateurs ;
   - dégager une saisonnalité par trimestre ;
   - contraster la tendance longue (toute la période) et la tendance courte
     (10 dernières semaines, déjà utilisée en Prévisions IA). */

function weekSortKey(wk){
  const m = /^S(\d{1,2})_(\d{4})$/.exec(wk||'');
  if(!m) return 0;
  return parseInt(m[2],10)*100 + parseInt(m[1],10);
}
function mean(arr){ return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; }
function stdev(arr){
  if(arr.length<2) return 0;
  const m = mean(arr);
  return Math.sqrt(arr.reduce((a,b)=>a+(b-m)*(b-m),0)/(arr.length-1));
}

function buildStatsData(trend, communesAOA){
  const sortedTrend = [...trend].sort((a,b)=>weekSortKey(a.week)-weekSortKey(b.week));

  // Semaines nationales statistiquement anormales : z-score sur les variations
  // hebdomadaires du taux (et non un seuil fixe comme "±1 point").
  const deltas = [];
  for(let i=1;i<sortedTrend.length;i++) deltas.push(sortedTrend[i].taux - sortedTrend[i-1].taux);
  const deltaMean = mean(deltas), deltaStd = stdev(deltas);
  const anomalousWeeks = [];
  for(let i=1;i<sortedTrend.length;i++){
    const d = sortedTrend[i].taux - sortedTrend[i-1].taux;
    const z = deltaStd ? (d-deltaMean)/deltaStd : 0;
    if(Math.abs(z)>=2){
      anomalousWeeks.push({week: sortedTrend[i].week, taux: sortedTrend[i].taux, delta: round2(d), z: round2(z)});
    }
  }
  anomalousWeeks.sort((a,b)=>Math.abs(b.z)-Math.abs(a.z));

  // tendance longue (toute la période dispo) vs tendance courte (10 dernières
  // semaines, celle utilisée pour la prévision en Prévisions IA)
  const fullSeries = sortedTrend.map(t=>t.taux);
  const regFull = fullSeries.length>=2 ? linreg(fullSeries) : {slope:0, intercept:0};
  const recent10 = fullSeries.slice(-10);
  const regRecent = recent10.length>=2 ? linreg(recent10) : {slope:0, intercept:0};

  // saisonnalité par trimestre (S01-13 / S14-26 / S27-39 / S40-52), moyenne
  // du taux national sur toutes les années disponibles
  const quarters = {Q1:[],Q2:[],Q3:[],Q4:[]};
  sortedTrend.forEach(t=>{
    const m = /^S(\d{1,2})_/.exec(t.week);
    if(!m) return;
    const wk = parseInt(m[1],10);
    const q = wk<=13?'Q1':wk<=26?'Q2':wk<=39?'Q3':'Q4';
    quarters[q].push(t.taux);
  });
  const seasonality = Object.entries(quarters).map(([q,vals])=>[q, vals.length?round2(mean(vals)):null, vals.length]);

  // écart de chaque commune par rapport à SA PROPRE trajectoire historique
  const cwMap = buildCommuneWeekMap(communesAOA);
  const communeHistory = {};
  Object.keys(cwMap).forEach(wk=>{
    Object.entries(cwMap[wk]).forEach(([commune,v])=>{
      if(!v.eql) return;
      if(!communeHistory[commune]) communeHistory[commune]=[];
      communeHistory[commune].push({weekKey: weekSortKey(wk), taux: v.racc/v.eql*100});
    });
  });
  const curWeekKey = trend.length ? weekSortKey(trend[trend.length-1].week) : 0;
  const communeZ = [];
  Object.entries(communeHistory).forEach(([commune,hist])=>{
    if(hist.length<8) return; // pas assez d'historique pour être statistiquement solide
    hist.sort((a,b)=>a.weekKey-b.weekKey);
    const cur = hist[hist.length-1];
    if(cur.weekKey!==curWeekKey) return;
    const histWithoutLast = hist.slice(0,-1).map(h=>h.taux);
    const m = mean(histWithoutLast), s = stdev(histWithoutLast);
    if(s===0) return;
    const z = (cur.taux-m)/s;
    communeZ.push({commune, tauxCur: round2(cur.taux), histMean: round2(m), histStd: round2(s), z: round2(z), nSemaines: hist.length});
  });
  communeZ.sort((a,b)=>Math.abs(b.z)-Math.abs(a.z));

  return {
    nWeeksHistory: sortedTrend.length,
    deltaMean: round2(deltaMean), deltaStd: round2(deltaStd),
    anomalousWeeks: anomalousWeeks.slice(0,15),
    regFull, regRecent,
    seasonality,
    communeZ: communeZ.slice(0,25),
    nCommunesAnalysees: communeZ.length,
  };
}


/* somme de la colonne "Clients Fictifs" pour une semaine donnée (accepte
   S8_2026 ou S08_2026 au cas où le padding diffère d'un onglet à l'autre) */
function sumFictifsForWeek(fictifsAOA, weekLabel){
  if(!fictifsAOA) return 0;
  const headers = fictifsAOA[0];
  const idx = {semaine:headers.indexOf('Semaine'), val:headers.indexOf('Clients Fictifs')};
  if(idx.semaine<0 || idx.val<0) return 0;
  const m = /^S0*(\d{1,2})_(\d{4})$/.exec(weekLabel||'');
  const variants = m ? new Set([weekLabel, 'S'+m[1]+'_'+m[2], 'S'+m[1].padStart(2,'0')+'_'+m[2]]) : new Set([weekLabel]);
  let sum = 0;
  for(let i=1;i<fictifsAOA.length;i++){
    const r = fictifsAOA[i];
    if(!r || r[idx.semaine]==null) continue;
    if(variants.has(String(r[idx.semaine]).trim())) sum += Number(r[idx.val])||0;
  }
  return sum;
}

/* "Total client" = somme(Nombre de ND) / 2 — confirmé */
function sumClientTotal(clientAOA){
  if(!clientAOA) return null;
  const headers = clientAOA[0];
  const valIdx = headers.indexOf('Nombre de ND');
  if(valIdx<0) return null;
  let sum = 0;
  for(let i=1;i<clientAOA.length;i++){
    const r = clientAOA[i];
    if(!r) continue;
    sum += Number(r[valIdx])||0;
  }
  return Math.round(sum/2);
}

function buildPlaquesByYear(plaques){
  const byYear = {};
  plaques.forEach(p=>{ if(p.openYear) byYear[p.openYear] = (byYear[p.openYear]||0)+1; });
  return byYear;
}

/* ================== Section "Pénétration FTTH / Plaque" ==================
   Entièrement dérivée de DATA.plaques (onglet TO_Plaques) : rien de manuel ici. */
function buildPlaquesPageData(plaques){
  const withEql = plaques.filter(p=>p.eql>0);

  const top15Taux = [...withEql].sort((a,b)=>b.taux-a.taux).slice(0,15)
    .map(p=>[p.plaque, p.taux]);

  const top15Pbo = [...plaques].sort((a,b)=>b.pbo-a.pbo).slice(0,15)
    .map(p=>[p.plaque, p.pbo]);

  // plaques ouvertes depuis + de 2 ans (24 mois) avec un faible taux de pénétration (<30%)
  const anciennesFaibles = withEql.filter(p=>p.age>=24 && p.taux<30)
    .sort((a,b)=>b.taux-a.taux)
    .map(p=>[p.plaque, p.zone, p.openDateLabel, p.taux]);

  // plaques ouvertes récemment (<= 6 mois), triées par taux croissant
  const recentes = plaques.filter(p=>p.age!=null && p.age<=6 && p.openDateLabel!=='—')
    .sort((a,b)=>a.taux-b.taux)
    .map(p=>[p.plaque, p.zone, p.openDateLabel, p.taux]);

  const drCount = {}, drTaux = {};
  plaques.forEach(p=>{
    const z = p.zone || '(non renseigné)';
    drCount[z] = (drCount[z]||0)+1;
    if(!drTaux[z]) drTaux[z] = {sum:0, n:0};
    if(p.eql>0){ drTaux[z].sum += p.taux; drTaux[z].n += 1; }
  });
  const plaquesParDR = Object.entries(drCount).sort((a,b)=>b[1]-a[1]);
  const tauxMoyenParDR = Object.entries(drTaux)
    .filter(([,v])=>v.n>0)
    .map(([dr,v])=>[dr, round2(v.sum/v.n)])
    .sort((a,b)=>b[1]-a[1]);

  return { top15Taux, top15Pbo, anciennesFaibles, recentes, plaquesParDR, tauxMoyenParDR };
}

/* ================== Section "Client Sans Constitution" ==================
   Entièrement dérivée de l'onglet "Clients Fictifs", au niveau plaque
   (colonnes Zone DRV / Plaque / Semaine / Clients Fictifs). */
function buildPlaqueFIMap(fictifsAOA){
  const map = {};
  if(!fictifsAOA) return map;
  const headers = fictifsAOA[0];
  const idx = {zone:headers.indexOf('Zone DRV'), plaque:headers.indexOf('Plaque'), semaine:headers.indexOf('Semaine'), val:headers.indexOf('Clients Fictifs')};
  if(idx.plaque<0 || idx.semaine<0 || idx.val<0) return map;
  for(let i=1;i<fictifsAOA.length;i++){
    const r = fictifsAOA[i];
    if(!r || r[idx.plaque]==null || r[idx.semaine]==null) continue;
    const wkRaw = String(r[idx.semaine]).trim();
    if(!/^S\d{1,2}_\d{4}$/.test(wkRaw)) continue;
    const wk = normalizeWeekLabel(wkRaw);
    if(!map[wk]) map[wk] = {};
    map[wk][r[idx.plaque]] = { val: Number(r[idx.val])||0, zone: r[idx.zone]||'' };
  }
  return map;
}

/* ================== Section "PBO Zéro Client" ==================
   Sujet ajouté à l'initiative de l'analyse du fichier
   "Fichier_Du_Taux_Pénétration_..." : des PBO (boîtiers de raccordement)
   sont déployés, câblés, mais n'ont RAPPORTÉ AUCUN client raccordé.
   C'est du potentiel de raccordement immédiatement disponible mais inexploité
   (peut être un problème terrain, commercial ou une simple mise à jour à faire),
   d'où l'onglet "Action Requise" du fichier source. Rien de tel n'existait
   jusqu'ici dans le dashboard : les autres pages parlent de saturation
   (trop de clients), jamais de sous-exploitation (zéro client). */
function buildPBOZeroData(wb){
  const aoa = getSheetAOA(wb, 'PBO_0_Client');
  if(!aoa || aoa.length<2) return null;
  const headers = aoa[0];
  const idx = {
    nro: headers.indexOf('NRO'), plaque: headers.indexOf('Plaque'), pbo: headers.indexOf('PBO'),
    brins: headers.indexOf('Total Brins'), occ: headers.indexOf('Brins Occupés'), dr: headers.indexOf('DR'),
    action: headers.indexOf('Action Requise'), commentaire: headers.indexOf('Commentaire'),
  };
  const rows = [];
  for(let i=1;i<aoa.length;i++){
    const r = aoa[i];
    if(!r || r[idx.plaque]==null) continue;
    rows.push({
      nro: r[idx.nro]||'', plaque: r[idx.plaque]||'', pbo: r[idx.pbo]||'',
      brins: Number(r[idx.brins])||0, dr: r[idx.dr]||'(non renseigné)',
      action: r[idx.action]||'', commentaire: r[idx.commentaire]||'',
    });
  }
  const drMap = {};
  rows.forEach(r=>{
    if(!drMap[r.dr]) drMap[r.dr] = {count:0, brins:0};
    drMap[r.dr].count += 1; drMap[r.dr].brins += r.brins;
  });
  const parDR = Object.entries(drMap).map(([dr,v])=>[dr, v.count, v.brins]).sort((a,b)=>b[1]-a[1]);

  const plaqueMap = {};
  rows.forEach(r=>{
    if(!plaqueMap[r.plaque]) plaqueMap[r.plaque] = {count:0, brins:0, dr:r.dr};
    plaqueMap[r.plaque].count += 1; plaqueMap[r.plaque].brins += r.brins;
  });
  const topPlaques = Object.entries(plaqueMap).map(([p,v])=>[p, v.count, v.brins, v.dr])
    .sort((a,b)=>b[2]-a[2]).slice(0,15);

  return {
    rows, parDR, topPlaques,
    totalPBO: rows.length,
    totalBrins: rows.reduce((s,r)=>s+r.brins,0),
    plaquesConcernees: Object.keys(plaqueMap).length,
  };
}

/* ================== Section "Analyse du taux de pénétration FTTH < 50%" ==================
   Sourcée depuis le 2e classeur (Fichier_Du_Taux_Pénétration_...xlsx), feuilles
   "Pénétration < à 50%" (liste plaque par plaque) et "Données chiffrées par
   commune" (pour identifier les communes les plus faibles). */
function buildPenetrationBelow50Data(wb){
  const aoa = getSheetAOA(wb, 'Pénétration < à 50%');
  if(!aoa || aoa.length<2) return null;
  const headers = aoa[0];
  const idx = {
    plaque: headers.indexOf('Plaques'), archi: headers.indexOf('ARCHI'), dr: headers.indexOf('Zone DR'),
    zone: headers.indexOf('Zone de couverture'), s28: headers.indexOf('Taux de pénétration S28_26'),
    s29: headers.indexOf('Taux de pénétration S29_26'), evo: headers.indexOf("Taux d'évolution"),
    age: headers.indexOf('Age'),
  };
  const rows = [];
  for(let i=1;i<aoa.length;i++){
    const r = aoa[i];
    if(!r || r[idx.plaque]==null) continue;
    rows.push({
      plaque: r[idx.plaque]||'', archi: r[idx.archi]||'', dr: r[idx.dr]||'',
      zone: r[idx.zone]||'', s28: round2((Number(r[idx.s28])||0)*100), s29: round2((Number(r[idx.s29])||0)*100),
      ecart: round2((Number(r[idx.evo])||0)*100), age: Number(r[idx.age])||0,
    });
  }
  rows.sort((a,b)=>a.plaque.localeCompare(b.plaque));

  const zonesCouverture = [...new Set(rows.map(r=>r.zone))].filter(Boolean).sort((a,b)=>a.localeCompare(b));
  const zonesDR = [...new Set(rows.map(r=>r.dr))].filter(Boolean).sort((a,b)=>a.localeCompare(b));

  // top 6 des plus fortes baisses (évolution la plus négative) parmi les plaques < 50%
  const worstEvolution = [...rows].sort((a,b)=>a.ecart-b.ecart).slice(0,6)
    .map(r=>[r.plaque, r.ecart]);

  // top 5 communes les plus faibles (< 50%), depuis la feuille "Données chiffrées par commune"
  let worstCommunes = [];
  const communesAOA2 = getSheetAOA(wb, 'Données chiffrées par commune');
  if(communesAOA2 && communesAOA2.length>1){
    const ch = communesAOA2[0];
    const cidx = {
      commune: ch.indexOf('COMMUNE'), s28: ch.indexOf('Taux de pénétration S28_2026'), s29: ch.indexOf('Taux de pénétration S29_2026'),
    };
    const communeRows = [];
    for(let i=1;i<communesAOA2.length;i++){
      const r = communesAOA2[i];
      if(!r || r[cidx.commune]==null || r[cidx.s29]==null) continue;
      const s29 = Number(r[cidx.s29])*100;
      if(s29>=50) continue;
      communeRows.push({commune:r[cidx.commune], s28: round2(Number(r[cidx.s28])*100), s29: round2(s29)});
    }
    worstCommunes = communeRows.sort((a,b)=>a.s29-b.s29).slice(0,5).sort((a,b)=>b.s29-a.s29);
  }

  return { rows, zonesCouverture, zonesDR, worstEvolution, worstCommunes, total: rows.length };
}

function buildSansConstitutionData(fictifsAOA, curWeek, prevWeek){
  const fiMap = buildPlaqueFIMap(fictifsAOA);
  const curMap = fiMap[curWeek] || {};
  const prevMap = fiMap[prevWeek] || {};

  // clients sans constitution par zone DR (semaine courante)
  const drSums = {};
  Object.values(curMap).forEach(({val,zone})=>{
    const z = zone || '(non renseigné)';
    drSums[z] = (drSums[z]||0)+val;
  });
  const parDR = Object.entries(drSums).sort((a,b)=>b[1]-a[1]);

  // top 10 plaques (semaine courante)
  const top10 = Object.entries(curMap).map(([plaque,v])=>[plaque, v.val])
    .sort((a,b)=>b[1]-a[1]).slice(0,10);

  // évolution plaque par plaque entre les deux semaines
  const plaqueNames = new Set([...Object.keys(curMap), ...Object.keys(prevMap)]);
  const evolution = [...plaqueNames].map(plaque=>{
    const c = curMap[plaque], p = prevMap[plaque];
    const cur = c?c.val:0, prev = p?p.val:0;
    return { plaque, zone: (c&&c.zone)||(p&&p.zone)||'', cur, prev, ecart: cur-prev };
  }).sort((a,b)=>b.cur-a.cur);

  // Innovation : détection auto des hausses anormales (fiabilisation à prioriser).
  // Une plaque est signalée si elle gagne au moins 50 FI ET que ça représente
  // une hausse d'au moins 40% par rapport à la semaine précédente (ou apparaît
  // franchement du jour au lendemain avec un volume déjà significatif).
  const alerts = evolution.filter(r=>{
    if(r.ecart<50) return false;
    if(r.prev===0) return r.cur>=50;
    return (r.ecart/r.prev) >= 0.4;
  }).sort((a,b)=>b.ecart-a.ecart);

  return { curWeek, prevWeek, parDR, top10, evolution, alerts };
}

function buildPenetrationData(plaques, communesAOA, fictifsAOA, clientAOA, curWeek, prevWeek, globalCur, globalPrev){
  const cwMap = buildCommuneWeekMap(communesAOA);
  const curMap = cwMap[curWeek] || {};
  const prevMap = cwMap[prevWeek] || {};

  const sansConstCur = sumFictifsForWeek(fictifsAOA, curWeek);
  const sansConstPrev = sumFictifsForWeek(fictifsAOA, prevWeek);
  const raccInclCur = (globalCur.racc||0) + sansConstCur;
  const raccInclPrev = (globalPrev.racc||0) + sansConstPrev;
  const tauxPenCur = globalCur.eql ? round2(raccInclCur/globalCur.eql*100) : 0;
  const tauxPenPrev = globalPrev.eql ? round2(raccInclPrev/globalPrev.eql*100) : 0;

  // tableau commune par commune (S-1 vs S), trié alphabétiquement
  const communeNames = new Set([...Object.keys(curMap), ...Object.keys(prevMap)]);
  const tauxRows = [...communeNames].map(nom=>{
    const c = curMap[nom], p = prevMap[nom];
    const tauxC = c && c.eql ? c.racc/c.eql*100 : null;
    const tauxP = p && p.eql ? p.racc/p.eql*100 : null;
    return {
      nom, eqlCur: c?c.eql:0, raccCur: c?c.racc:0,
      tauxCur: tauxC, tauxPrev: tauxP,
      evolution: (tauxC!=null && tauxP!=null) ? tauxC-tauxP : null,
    };
  }).sort((a,b)=>a.nom.localeCompare(b.nom,'fr'));

  // communes à fort taux de pénétration (semaine courante, EQL significatif)
  const forte = tauxRows.filter(r=>r.tauxCur!=null && r.eqlCur>=200)
    .sort((a,b)=>b.tauxCur-a.tauxCur).slice(0,5)
    .map(r=>[r.nom, round2(r.tauxCur)]);

  // répartition des communes par tranche de taux (le nombre total de communes
  // suit TO_Communes[Semaine]=cur, sans condition d'EQL ; les tranches
  // excluent les communes sans EQL, un taux étant alors indéfini)
  const communesFibrees = Object.keys(curMap).length;
  const withTaux = tauxRows.filter(r=>r.tauxCur!=null && r.eqlCur>0);
  const above40 = withTaux.filter(r=>r.tauxCur>40).length;
  const between20_40 = withTaux.filter(r=>r.tauxCur>=20 && r.tauxCur<=40).length;
  const below20 = withTaux.filter(r=>r.tauxCur<20).length;

  // classement des DR (semaine courante)
  const drMap = {};
  Object.values(curMap).forEach((c,i)=>{});
  Object.keys(curMap).forEach(nom=>{
    const c = curMap[nom];
    const dr = c.drv || '(non renseigné)';
    if(!drMap[dr]) drMap[dr] = {eql:0, racc:0};
    drMap[dr].eql += c.eql; drMap[dr].racc += c.racc;
  });
  const drRanking = Object.entries(drMap).map(([dr,v])=>[dr, v.eql, v.racc]).sort((a,b)=>b[2]-a[2]);
  const drTotal = ['Total', drRanking.reduce((s,r)=>s+r[1],0), drRanking.reduce((s,r)=>s+r[2],0)];

  // top 5 communes par clients raccordés (semaine courante)
  const top5 = tauxRows.filter(r=>r.raccCur>0).sort((a,b)=>b.raccCur-a.raccCur).slice(0,5)
    .map(r=>[r.nom, r.raccCur]);

  // plaques déployées par année (depuis TO_Plaques)
  const plaquesByYear = buildPlaquesByYear(plaques);

  const totalClientCur = sumClientTotal(clientAOA); // manuel/mémorisé — voir refreshPenSection
  // "En cours de fiabilisation" = Total_client - Total_Clients_Semaine (Clients_Raccordés + NB_FI) — formule DAX fournie
  // En cours de fiabilisation = Total client - Total clients raccordés - Raccordés sans constitution — formule DAX fournie
  const enCoursFiabilisationCur = totalClientCur!=null ? (totalClientCur - raccInclCur - sansConstCur) : null;

  // Taux_Evolution_Taux_Pénétration_Global = (Suivante-Précédente)/Précédente — formule DAX fournie
  const tauxEvolution = tauxPenPrev ? round2((tauxPenCur-tauxPenPrev)/tauxPenPrev*100) : null;

  return {
    curWeek, prevWeek,
    tauxCur: tauxPenCur, tauxPrev: tauxPenPrev, tauxEvolution,
    eqlCur: globalCur.eql, eqlPrev: globalPrev.eql,
    totalClientCur, enCoursFiabilisationCur,
    totalClientPrev: null, enCoursFiabilisationPrev: null, // renseignés côté serveur via l'historique (voir index.js)
    raccInclCur, raccInclPrev,
    sansConstCur, sansConstPrev,
    // Pourcentage_FI = NB_FI / Clients_Raccordés (racc brut, hors FI) — formule DAX fournie
    sansConstPctCur: globalCur.racc ? round2(sansConstCur/globalCur.racc*100) : 0,
    sansConstPctPrev: globalPrev.racc ? round2(sansConstPrev/globalPrev.racc*100) : 0,
    tauxRows, forte, top5, drRanking, drTotal,
    communesFibrees, above40, between20_40, below20,
    plaquesTotal: plaques.length, plaquesByYear,
  };
}

function buildDataFromWorkbook(wb){
  const plaquesAOA = getSheetAOA(wb, 'TO_Plaques');
  const synthAOA = getSheetAOA(wb, 'SYNTHESE');
  const communesAOA = getSheetAOA(wb, 'TO_Communes');
  if(!plaquesAOA || !synthAOA || !communesAOA){
    throw new Error("Le classeur doit contenir les feuilles 'TO_Plaques', 'SYNTHESE' et 'TO_Communes'.");
  }
  const plaques = buildPlaques(plaquesAOA);
  if(!plaques.length) throw new Error("Aucune plaque trouvée dans l'onglet TO_Plaques.");
  const regionNaming = buildRegionNaming(plaques);

  const plaqueCountByRegion = {}, pboSumByRegion = {}, pboCountByRegion = {};
  plaques.forEach(p=>{
    const rn = normalizeRegion(p.region);
    plaqueCountByRegion[rn] = (plaqueCountByRegion[rn]||0)+1;
    pboSumByRegion[rn] = (pboSumByRegion[rn]||0)+p.pbo;
    pboCountByRegion[rn] = (pboCountByRegion[rn]||0)+1;
  });
  const pboAvgByRegion = {};
  Object.keys(pboSumByRegion).forEach(k=>{ pboAvgByRegion[k]=pboSumByRegion[k]/pboCountByRegion[k]; });

  const syntheseRows = buildSynthese(synthAOA);
  const regions = buildRegionsFromSynthese(syntheseRows, plaqueCountByRegion, pboAvgByRegion);
  const {top10, bottom10} = buildTopBottom(syntheseRows);
  const trend = buildTrend(communesAOA);
  const anomalies = buildAnomalies(plaques, regionNaming);
  const hist = buildHistogram(plaques);
  const global = buildGlobal(trend, plaques.length);

  const week = trend.length ? trend[trend.length-1].week : (DATA.week||'S00_2026');
  const m = /^S(\d{1,2})_(\d{4})$/.exec(week);
  const weekLabel = m ? `Semaine ${parseInt(m[1],10)}, ${m[2]}` : week;

  const fictifsAOA = getSheetAOA(wb, 'Clients Fictifs');
  const clientAOA = getSheetAOA(wb, 'Client');
  const curTrend = trend[trend.length-1] || {eql:0,racc:0};
  const prevTrend = trend[trend.length-2] || curTrend;
  const prevWeek = trend.length>=2 ? trend[trend.length-2].week : week;
  const penetration = buildPenetrationData(plaques, communesAOA, fictifsAOA, clientAOA, week, prevWeek, curTrend, prevTrend);
  const pen_plaques = buildPlaquesPageData(plaques);
  const sans_constitution = buildSansConstitutionData(fictifsAOA, week, prevWeek);
  const stats = buildStatsData(trend, communesAOA);

  return {
    week, week_label: weekLabel,
    global, regions, trend, top10, bottom10,
    anomalies, plaques, penetration, pen_plaques, sans_constitution, stats,
    hist_labels: hist.labels, hist_values: hist.values,
    source_file: (typeof importedFileName!=='undefined' && importedFileName) || `Suivi_Taux_d_occupation_de_la_fibre_${week}.xlsx`,
  };
}

module.exports = { buildDataFromWorkbook, buildPBOZeroData, buildPenetrationBelow50Data, round2 };
