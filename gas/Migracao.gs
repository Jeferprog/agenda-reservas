/**
 * Migração ÚNICA: copia os usuários da coleção `users` do Firestore
 * para a Planilha Google (aba "Usuarios").
 *
 * COMO USAR (rodar uma única vez):
 *   1. Cole este arquivo no MESMO projeto Apps Script do Code.gs.
 *   2. Garanta que USERS_SHEET_ID já está configurado (ver gas/README.md).
 *   3. No editor do Apps Script, selecione a função
 *      "migrarUsuariosDoFirestore" e clique em ▶ Executar.
 *      (autorize os escopos solicitados — ver README, seção Migração).
 *   4. Veja o resultado em "Registros de execução" (Logs).
 *   5. Confira a planilha. Depois exclua a coleção `users` do Firestore.
 *
 * Reaproveita getSheet_(), readUsers_() e colorFor_() do Code.gs.
 */

const FIRESTORE_PROJECT_ID = 'minhaagenda-5fc2d';

function migrarUsuariosDoFirestore() {
  // usuários já presentes na planilha (para não duplicar)
  const existing = {};
  readUsers_().forEach(u => existing[u.username.toLowerCase()] = true);

  const docs = fetchFirestoreCollection_('users');
  const sh = getSheet_();
  let add = 0, skip = 0, semNome = 0;

  docs.forEach(d => {
    const f = d.fields || {};
    const username = strVal_(f.username) || strVal_(f.name); // 'name' = legado
    if (!username) { semNome++; return; }

    if (existing[username.toLowerCase()]) { skip++; return; }

    const password = strVal_(f.password);
    let role = strVal_(f.role) || 'user';
    role = (role === 'admin') ? 'admin' : 'user';
    let color = strVal_(f.customColor) || strVal_(f.color);
    if (!color) color = colorFor_(username);

    sh.appendRow([Utilities.getUuid(), username, password, role, color]);
    existing[username.toLowerCase()] = true;
    add++;
  });

  const msg = `Migração concluída: ${add} adicionado(s), ${skip} já existia(m), ${semNome} sem nome (ignorado(s)). Total lido do Firestore: ${docs.length}.`;
  Logger.log(msg);
  return msg;
}

/** Lê todos os documentos de uma coleção via Firestore REST (acesso admin do dono). */
function fetchFirestoreCollection_(collection) {
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

/** Extrai o valor textual de um campo do formato REST do Firestore. */
function strVal_(field) {
  if (!field) return '';
  if (field.stringValue != null) return String(field.stringValue);
  if (field.integerValue != null) return String(field.integerValue);
  return '';
}
