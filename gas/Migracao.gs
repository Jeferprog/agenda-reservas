/**
 * Migração ÚNICA: copia os usuários da coleção `users` do Firestore
 * para a Planilha Google (aba "Usuarios").
 *
 * Este arquivo é AUTOSSUFICIENTE — não depende do Code.gs. Pode ser colado
 * sozinho num projeto Apps Script só para rodar a migração.
 *
 * COMO USAR (rodar uma única vez):
 *   1. Configure a propriedade do script USERS_SHEET_ID (ver gas/README.md).
 *   2. Garanta os escopos do manifesto (gas/appsscript.json), incluindo
 *      "datastore".
 *   3. No editor do Apps Script, selecione a função
 *      "migrarUsuariosDoFirestore" e clique em ▶ Executar. Autorize.
 *   4. Veja o resultado em "Registros de execução" (Logs).
 *   5. Confira a planilha. Depois exclua a coleção `users` do Firestore.
 */

const FIRESTORE_PROJECT_ID = 'minhaagenda-5fc2d';

function migrarUsuariosDoFirestore() {
  const sh = mig_getSheet_();

  // usuários já presentes na planilha (para não duplicar)
  const existing = {};
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) { // pula cabeçalho
    const nm = String(values[i][1] || '').toLowerCase();
    if (nm) existing[nm] = true;
  }

  const docs = mig_fetchFirestore_('users');
  let add = 0, skip = 0, semNome = 0;

  docs.forEach(d => {
    const f = d.fields || {};
    const username = mig_str_(f.username) || mig_str_(f.name); // 'name' = legado
    if (!username) { semNome++; return; }
    if (existing[username.toLowerCase()]) { skip++; return; }

    const password = mig_str_(f.password);
    let role = mig_str_(f.role) || 'user';
    role = (role === 'admin') ? 'admin' : 'user';
    let color = mig_str_(f.customColor) || mig_str_(f.color);
    if (!color) color = mig_colorFor_(username);

    sh.appendRow([Utilities.getUuid(), username, password, role, color]);
    existing[username.toLowerCase()] = true;
    add++;
  });

  const msg = `Migração concluída: ${add} adicionado(s), ${skip} já existia(m), ` +
              `${semNome} sem nome (ignorado(s)). Total lido do Firestore: ${docs.length}.`;
  Logger.log(msg);
  return msg;
}

/** Abre a aba "Usuarios" da planilha definida em USERS_SHEET_ID. */
function mig_getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('USERS_SHEET_ID');
  if (!id) throw new Error('Configure a propriedade do script USERS_SHEET_ID.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName('Usuarios');
  if (!sh) sh = ss.insertSheet('Usuarios');
  if (sh.getLastRow() === 0) sh.appendRow(['id', 'username', 'password', 'role', 'color']);
  return sh;
}

/** Cor consistente a partir do nome (mesma paleta do app). */
function mig_colorFor_(username) {
  const palette = ['#005c46', '#f58220', '#d62828', '#003049', '#f77f00',
                   '#2a9d8f', '#264653', '#e76f51', '#6610f2', '#e83e8c'];
  let hash = 0;
  for (let i = 0; i < username.length; i++) hash = username.charCodeAt(i) + ((hash << 5) - hash);
  return palette[Math.abs(hash) % palette.length];
}

/** Lê todos os documentos de uma coleção via Firestore REST (acesso admin do dono). */
function mig_fetchFirestore_(collection) {
  const token = ScriptApp.getOAuthToken();
  let docs = [], pageToken = '';
  do {
    let url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
              '/databases/(default)/documents/' + collection + '?pageSize=300';
    if (pageToken) url += '&pageToken=' + encodeURIComponent(pageToken);

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Erro ao ler o Firestore (HTTP ' + resp.getResponseCode() + '): ' + resp.getContentText());
    }
    const data = JSON.parse(resp.getContentText());
    if (data.documents) docs = docs.concat(data.documents);
    pageToken = data.nextPageToken || '';
  } while (pageToken);
  return docs;
}

/** Extrai o valor textual de um campo no formato REST do Firestore. */
function mig_str_(field) {
  if (!field) return '';
  if (field.stringValue != null) return String(field.stringValue);
  if (field.integerValue != null) return String(field.integerValue);
  return '';
}
