/**
 * Conti Chalet Benvenuto — collegamento tra l'app e questo Foglio Google.
 *
 * COME SI INSTALLA (una volta sola):
 * 1. Nel Foglio Google: Estensioni > Apps Script. Cancella tutto e incolla questo file.
 * 2. Qui sotto cambia il PIN con uno tuo (lo scriverai nell'app la prima volta).
 * 3. Salva (icona del dischetto), poi in alto scegli la funzione "prepara" e premi Esegui.
 *    Google chiede i permessi: Rivedi autorizzazioni > il tuo account > Avanzate > Vai a ... > Consenti.
 * 4. Esegui il deployment > Nuovo deployment > tipo "App web".
 *    Esegui come: Me.  Chi ha accesso: Chiunque.  Premi Esegui il deployment e copia l'URL dell'app web.
 * 5. Apri l'app, incolla l'URL e il PIN. Fatto.
 *
 * I dati stanno solo in questo Foglio. Senza il PIN nessuno può leggerli o modificarli tramite l'app.
 */
const PIN = 'CAMBIAMI';

const COLONNE = {
  movimenti: ['id','data','dataFine','tipo','piano','canale','categoria','descrizione','importo','caparra','saldato','lenzuola','pulizie','note','esercizio','aggiornato'],
  esercizi:  ['id','inizio','fondoCassa','dataFondoCassa','utiliPT','utiliPB','quotaPT','provvPct','provvPagate','provvGiacomo','provvPapa','cassaReale','dataCassaReale','chiuso','pulizieDal'],
  config:    ['id','valore']
};
const DATE = ['data','dataFine','inizio','dataFondoCassa','dataCassaReale','pulizieDal'];
const TESTO = ['id','tipo','piano','canale','categoria','descrizione','note','esercizio','aggiornato','valore'];

function prepara() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(COLONNE).forEach(nome => {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    const cols = COLONNE[nome];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#EFE6D5');
    sh.setFrozenRows(1);
    // colonne di testo: niente conversioni automatiche (es. "2026-27" non deve diventare una data)
    cols.forEach((c, i) => { if (TESTO.indexOf(c) >= 0) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@'); });
  });
  const vuoto = ss.getSheetByName('Foglio1') || ss.getSheetByName('Sheet1');
  if (vuoto && ss.getSheets().length > 1 && vuoto.getLastRow() === 0) ss.deleteSheet(vuoto);
}

function doGet() {
  return ContentService.createTextOutput('Conti Chalet Benvenuto: collegamento attivo.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return risposta({ errore: 'richiesta non valida' }); }
  if (!PIN || PIN === 'CAMBIAMI' || String(req.pin) !== String(PIN)) return risposta({ errore: 'pin' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.azione === 'salva') salva(req.foglio, req.riga);
    else if (req.azione === 'elimina') elimina(req.foglio, req.id);
    else if (req.azione === 'salvaTanti') (req.righe || []).forEach(r => salva(req.foglio, r));
    return risposta(tutto());
  } catch (err) {
    return risposta({ errore: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function risposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function foglio(nome) {
  if (!COLONNE[nome]) throw new Error('foglio sconosciuto: ' + nome);
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('manca il foglio "' + nome + '": esegui la funzione prepara');
  return sh;
}

function leggi(nome) {
  const sh = foglio(nome), cols = COLONNE[nome], n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return sh.getRange(2, 1, n, cols.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      cols.forEach((c, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        o[c] = v;
      });
      return o;
    });
}

function tutto() {
  const config = {};
  leggi('config').forEach(r => { try { config[r.id] = JSON.parse(r.valore); } catch (e) { config[r.id] = r.valore; } });
  return {
    movimenti: leggi('movimenti'),
    esercizi: leggi('esercizi'),
    config: config,
    foglio: SpreadsheetApp.getActive().getUrl()
  };
}

function salva(nome, riga) {
  if (!riga || !riga.id) throw new Error('manca id');
  const sh = foglio(nome), cols = COLONNE[nome];
  if (nome === 'config' && typeof riga.valore !== 'string') riga.valore = JSON.stringify(riga.valore);
  const valori = cols.map(c => {
    const v = riga[c];
    if (v === undefined || v === null) return '';
    if (DATE.indexOf(c) >= 0 && /^\d{4}-\d{2}-\d{2}$/.test(v)) { const p = v.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    return v;
  });
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(String(riga.id));
  if (i >= 0) sh.getRange(i + 2, 1, 1, cols.length).setValues([valori]);
  else sh.appendRow(valori);
  DATE.forEach(c => { const k = cols.indexOf(c); if (k >= 0) sh.getRange(2, k + 1, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('dd/mm/yyyy'); });
}

function elimina(nome, id) {
  const sh = foglio(nome), n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}
