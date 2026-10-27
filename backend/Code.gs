/**
 * Centre d’activités l’Orchidbleue — Sondage du temps des Fêtes 2026–2027.
 * Déployer en application Web depuis un classeur Google Sheets (Extensions → Apps Script).
 * Les noms sont NOMINATIFS : restreindre l’accès au classeur aux personnes autorisées.
 */

const NOM_FEUILLE_REPONSES = 'Réponses';
const NOM_FEUILLE_BILAN = 'Bilan';
const JOURS_OUVERTS = Object.freeze([
  ['2026-12-21', 'Lun. 21 décembre'],
  ['2026-12-22', 'Mar. 22 décembre'],
  ['2026-12-23', 'Mer. 23 décembre'],
  ['2026-12-28', 'Lun. 28 décembre'],
  ['2026-12-29', 'Mar. 29 décembre'],
  ['2026-12-30', 'Mer. 30 décembre']
]);
const ENTETES = ['Identifiant de réponse','Dernière modification','Personne responsable','Participant·e',
  'Aucun besoin','Journées sélectionnées'].concat(JOURS_OUVERTS.map(function(j){return j[1];}));

function initialiserSondage() {
  const classeur = obtenirClasseur_();
  // Conserve les anciennes réponses dans un onglet archivé si les colonnes ont changé.
  const feuilleReponses = preparerFeuilleReponses_(classeur, true);
  normaliserAucunBesoin_(feuilleReponses);
  let bilan = classeur.getSheetByName(NOM_FEUILLE_BILAN);
  if (!bilan) bilan = classeur.insertSheet(NOM_FEUILLE_BILAN);
  bilan.getRange('A1:C1').breakApart();
  bilan.clear();
  bilan.getRange('A1:C1').merge();
  bilan.getRange('A1').setValue('L’ORCHIDBLEUE · BILAN DES FÊTES');
  bilan.getRange('A1:C1').setBackground('#245e80').setFontColor('white').setFontSize(14).setFontWeight('bold');
  bilan.setRowHeight(1,44);
  bilan.getRange('A2').setValue('Réponses reçues');
  bilan.getRange('B2').setFormula("=COUNTA('Réponses'!A2:A)");
  bilan.getRange('A3').setValue('Sans besoin de service');
  bilan.getRange('B3').setFormula("=SUM('Réponses'!E2:E)");
  bilan.getRange('A5:C5').setValues([['JOUR','BESOINS EXPRIMÉS','STATUT']]);
  bilan.getRange('A5:C5').setBackground('#deedf3').setFontColor('#245e80').setFontWeight('bold');
  JOURS_OUVERTS.forEach(function(jour,index){
    const ligne=6+index;
    const colonne=lettreColonne_(7+index);
    bilan.getRange(ligne,1).setValue(jour[1]);
    bilan.getRange(ligne,2).setFormula("=SUM('Réponses'!"+colonne+'2:'+colonne+')');
    bilan.getRange(ligne,3).setValue('Offert');
    if(index%2===1) bilan.getRange(ligne,1,1,3).setBackground('#f1f7f9');
  });
  bilan.getRange('A13').setValue('Dates non offertes : 24, 25, 26 et 31 décembre; 1er et 2 janvier.');
  bilan.getRange('A14').setValue('Une même personne peut sélectionner plusieurs journées.');
  bilan.getRange('A15').setValue('Les réponses sont nominatives et réservées à la planification.');
  bilan.getRange('A13:A15').setFontColor('#728e9e').setFontSize(10);
  bilan.setColumnWidth(1,295);bilan.setColumnWidth(2,175);bilan.setColumnWidth(3,165);
  bilan.setFrozenRows(5);
  bilan.getRange('B2:B11').setHorizontalAlignment('center');
  bilan.getRange('B2:B3').setFontSize(18).setFontWeight('bold').setFontColor('#245e80');
  bilan.getCharts().forEach(function(c){bilan.removeChart(c);});
  const graphique=bilan.newChart().asBarChart().addRange(bilan.getRange('A5:B11'))
    .setNumHeaders(1).setOption('title','Besoins par journée')
    .setOption('legend',{position:'none'}).setOption('colors',['#3c91be'])
    .setOption('hAxis',{minValue:0,title:'Nombre de participant·e·s'})
    .setPosition(5,5,0,0).build();
  bilan.insertChart(graphique);
  classeur.setActiveSheet(bilan);
  SpreadsheetApp.flush();
  return 'Feuilles « Réponses » et « Bilan » prêtes.';
}

function doGet() { return ContentService.createTextOutput('Service du sondage Orchidbleue actif.'); }
function doPost(e) {
  let nonce='';
  try {
    const brut=e&&e.parameter?e.parameter.payload:'';
    if(typeof brut!=='string'||brut.length>4500) throw new Error('Requête invalide');
    const payload=JSON.parse(brut);
    nonce=typeof payload.nonce==='string'?payload.nonce:'';
    if(!/^[a-zA-Z0-9_-]{16,100}$/.test(nonce)) throw new Error('Requête invalide');
    if(typeof payload.id!=='string'||!/^[a-zA-Z0-9_-]{16,80}$/.test(payload.id)) throw new Error('Identifiant invalide');
    const responsable=validerNom_(payload.responsable);
    const participant=validerNom_(payload.participant);
    if(!Array.isArray(payload.dates)||payload.dates.length>JOURS_OUVERTS.length) throw new Error('Dates invalides');
    if(typeof payload.aucunBesoin!=='boolean'||payload.aucunBesoin!==(payload.dates.length===0)) throw new Error('Choix incohérent');
    const possibles=JOURS_OUVERTS.map(function(j){return j[0];});
    if(new Set(payload.dates).size!==payload.dates.length||payload.dates.some(function(d){return typeof d!=='string'||possibles.indexOf(d)===-1;})) throw new Error('Date non offerte');

    const verrou=LockService.getScriptLock();
    verrou.waitLock(15000);
    try {
      const feuille=preparerFeuilleReponses_(obtenirClasseur_(), false);
      const dates=new Set(payload.dates);
      const ligne=[payload.id,new Date(),pourFeuille_(responsable),pourFeuille_(participant),
        payload.aucunBesoin?1:0,payload.dates.join(', ')]
        .concat(JOURS_OUVERTS.map(function(j){return dates.has(j[0])?1:0;}));
      const total=feuille.getLastRow();
      let cible=-1;
      if(total>1) {
        const ids=feuille.getRange(2,1,total-1,1).getValues();
        for(let i=0;i<ids.length;i++) if(ids[i][0]===payload.id){cible=i+2;break;}
      }
      if(cible===-1) feuille.appendRow(ligne);
      else feuille.getRange(cible,1,1,ligne.length).setValues([ligne]);
      SpreadsheetApp.flush();
    } finally { verrou.releaseLock(); }
    return reponseIframe_(nonce,true);
  } catch(erreur) {
    console.error('Échec du sondage : '+erreur.message);
    return reponseIframe_(nonce,false);
  }
}
function validerNom_(valeur) {
  if(typeof valeur!=='string') throw new Error('Nom manquant');
  const nom=valeur.trim().replace(/\s+/g,' ');
  if(nom.length<2||nom.length>100||/[<>\u0000-\u001F]/.test(nom)) throw new Error('Nom invalide');
  return nom;
}
function pourFeuille_(valeur) {
  // Empêche l'interprétation de texte libre comme formule dans Google Sheets.
  return /^[=+@-]/.test(valeur)?"'"+valeur:valeur;
}
function obtenirClasseur_() {
  const classeur=SpreadsheetApp.getActiveSpreadsheet();
  if(!classeur) throw new Error('Créez ce script à partir d’un Google Sheets.');
  return classeur;
}
function preparerFeuilleReponses_(classeur, autoriserArchivage) {
  let feuille=classeur.getSheetByName(NOM_FEUILLE_REPONSES);
  if(!feuille) feuille=classeur.insertSheet(NOM_FEUILLE_REPONSES);
  if(feuille.getLastRow()>0) {
    const existants=feuille.getRange(1,1,1,Math.min(feuille.getLastColumn(),ENTETES.length)).getValues()[0];
    const conforme=ENTETES.every(function(t,i){return existants[i]===t;});
    if(conforme) return feuille;

    // Archive les anciennes colonnes au lieu d'effacer les réponses du premier test.
    if(feuille.getLastRow()>1) {
      if(!autoriserArchivage) {
        throw new Error('Ancienne version détectée : exécutez initialiserSondage depuis Apps Script.');
      }
      const date=Utilities.formatDate(new Date(),classeur.getSpreadsheetTimeZone(),'yyyyMMdd-HHmmss');
      let base='Réponses - archive '+date;
      let nom=base;
      let n=2;
      while(classeur.getSheetByName(nom)) nom=base+'-'+n++;
      feuille.setName(nom);
      feuille=classeur.insertSheet(NOM_FEUILLE_REPONSES);
    } else {
      // L'ancienne feuille n'avait que les en-têtes et aucune donnée.
      feuille.clear();
    }
  }
  feuille.getRange(1,1,1,ENTETES.length).setValues([ENTETES]);
  feuille.getRange(1,1,1,ENTETES.length).setBackground('#245e80').setFontColor('white').setFontWeight('bold');
  feuille.setFrozenRows(1);
  feuille.setColumnWidth(1,290);feuille.setColumnWidth(2,178);
  feuille.setColumnWidth(3,235);feuille.setColumnWidth(4,235);
  feuille.setColumnWidth(5,120);feuille.setColumnWidth(6,310);
  feuille.getRange('B:B').setNumberFormat('dd mmm yyyy HH:mm:ss');
  feuille.getRange('E1').setNote('1 = aucun besoin de service; 0 = journées sélectionnées. Ce format numérique évite les erreurs de calcul entre différentes langues de Google Sheets.');
  return feuille;
}

function normaliserAucunBesoin_(feuille) {
  // Compatibilité avec les réponses nominatives antérieures : TRUE/FALSE -> 1/0.
  // Ne touche pas aux noms, identifiants ni aux dates enregistrées.
  const derniereLigne=feuille.getLastRow();
  if(derniereLigne<=1) return;
  const valeurs=feuille.getRange(2,5,derniereLigne-1,1).getValues();
  valeurs.forEach(function(ligne,index){
    if(typeof ligne[0]==='boolean') feuille.getRange(index+2,5).setValue(ligne[0]?1:0);
  });
}

function lettreColonne_(n) {
  let nom='';while(n>0){n--;nom=String.fromCharCode(65+n%26)+nom;n=Math.floor(n/26);}return nom;
}
function reponseIframe_(nonce,ok) {
  const message=JSON.stringify({type:'orchidbleue-survey-result',nonce:nonce,ok:ok}).replace(/</g,'\\u003c');
  const html='<!doctype html><html lang="fr"><head><meta charset="utf-8"></head><body>'+
    '<script>window.top.postMessage('+message+',"*");</script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Confirmation du sondage')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
