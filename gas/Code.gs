/**
 * Agenda Cresol Cooperar — Backend (Google Apps Script)
 * ------------------------------------------------------------------
 * Login e gestão de usuários acontecem NO SERVIDOR.
 * As senhas ficam numa Planilha Google privada e NUNCA são enviadas
 * ao navegador. O cliente recebe apenas um token de sessão e os dados
 * públicos do usuário (nome, papel, cor).
 *
 * Locais (places) e reservas (reservations) continuam no Firestore,
 * acessados pelo navegador — não são dados sensíveis.
 *
 * SETUP (ver gas/README.md para o passo a passo completo):
 *   1. Crie uma Planilha Google privada (aba "Usuarios").
 *   2. Em "Configurações do projeto > Propriedades do script", crie:
 *        USERS_SHEET_ID = <id da planilha (parte da URL)>
 *   3. Implante como App da Web: "Executar como: Eu" e acesso
 *      restrito à sua organização.
 */

const SHEET_NAME = 'Usuarios';
const SESSION_TTL_SECONDS = 12 * 60 * 60; // 12 horas
const USER_PALETTE = ['#005c46', '#f58220', '#d62828', '#003049', '#f77f00',
                      '#2a9d8f', '#264653', '#e76f51', '#6610f2', '#e83e8c'];

/** Serve o app (arquivo HTML deve se chamar "AgendaCresolCooperar" no projeto GAS). */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('AgendaCresolCooperar')
    .setTitle('Agenda Cresol Cooperar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/* ============================ PLANILHA ============================ */

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('USERS_SHEET_ID');
  if (!id) throw new Error('Configure a propriedade do script USERS_SHEET_ID.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['id', 'username', 'password', 'role', 'color']);
  return sh;
}

function colorFor_(username) {
  let hash = 0;
  for (let i = 0; i < username.length; i++) hash = username.charCodeAt(i) + ((hash << 5) - hash);
  return USER_PALETTE[Math.abs(hash) % USER_PALETTE.length];
}

/** Lê todos os usuários da planilha (uso interno — inclui senha). */
function readUsers_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  return values.slice(1) // pula o cabeçalho
    .map((r, i) => ({
      rowIndex: i + 2, // linha real na planilha (1-based + cabeçalho)
      id: String(r[0]),
      username: String(r[1] == null ? '' : r[1]).trim(),
      password: String(r[2] == null ? '' : r[2]),
      role: normalizeRole_(r[3]),
      color: String(r[4] || '')
    }))
    .filter(u => u.username !== '')
    .map(u => { if (!u.color) u.color = colorFor_(u.username); return u; });
}

/** Versão segura (sem senha) para enviar ao cliente. */
function publicUser_(u) {
  return { id: u.id, username: u.username, role: u.role, color: u.color };
}

/* ============================ SESSÕES ============================ */
/* Sessão "stateless": o token é assinado (HMAC) com um segredo do projeto.
   Não depende de cache (que expira/é descartado), então o login do admin
   permanece válido até o vencimento (SESSION_TTL_SECONDS). */

function normalizeRole_(r) {
  return String(r || '').trim().toLowerCase() === 'admin' ? 'admin' : 'user';
}

function getSessionSecret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SESSION_SECRET', s); }
  return s;
}

function signPart_(data) {
  const sig = Utilities.computeHmacSha256Signature(data, getSessionSecret_());
  return Utilities.base64EncodeWebSafe(sig);
}

function createSession_(u) {
  const payload = { username: u.username, role: normalizeRole_(u.role), exp: Date.now() + SESSION_TTL_SECONDS * 1000 };
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  return body + '.' + signPart_(body);
}

function getSession_(token) {
  if (!token || String(token).indexOf('.') < 0) return null;
  const parts = String(token).split('.');
  const body = parts[0], sig = parts[1];
  if (!body || !sig || signPart_(body) !== sig) return null; // assinatura inválida
  let payload;
  try {
    payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString());
  } catch (e) { return null; }
  if (!payload || !payload.exp || payload.exp < Date.now()) return null; // expirada
  return payload;
}

function requireAdmin_(token) {
  const s = getSession_(token);
  if (!s) throw new Error('Sessão expirada. Faça login novamente.');
  if (normalizeRole_(s.role) !== 'admin') throw new Error('Ação permitida apenas para administradores.');
  return s;
}

/* ===================== API (google.script.run) ===================== */

/** Valida login no servidor. Retorna token + dados públicos do usuário. */
function login(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();
  if (!username || !password) return { ok: false, message: 'Informe usuário e senha.' };

  const u = readUsers_().find(x => String(x.username).trim().toLowerCase() === username.toLowerCase());
  if (!u) return { ok: false, message: 'Usuário não encontrado.' };
  if (String(u.password).trim() !== password) return { ok: false, message: 'Senha incorreta.' };

  return { ok: true, token: createSession_(u), user: publicUser_(u) };
}

function logout(token) {
  // Token stateless: nada a limpar no servidor; o cliente apaga o token.
  return { ok: true };
}

/**
 * Autoatendimento com trava de primeiro acesso.
 * - 1º acesso (senha ainda é o hash legado): redefine só com usuário + nova senha.
 * - Conta já com senha própria: exige a senha ATUAL para trocar (evita que
 *   terceiros redefinam a senha de quem já configurou a sua).
 */
function redefinirMinhaSenha(username, novaSenha, senhaAtual) {
  username = String(username || '').trim();
  novaSenha = String(novaSenha || '').trim();
  senhaAtual = String(senhaAtual || '').trim();
  if (!username || !novaSenha) return { ok: false, message: 'Informe usuário e nova senha.' };
  if (novaSenha.length < 4) return { ok: false, message: 'A nova senha deve ter ao menos 4 caracteres.' };

  const u = readUsers_().find(x => String(x.username).trim().toLowerCase() === username.toLowerCase());
  if (!u) return { ok: false, message: 'Usuário não encontrado.' };

  const atual = String(u.password).trim();
  // Senha legada (hash) tem ~101 caracteres; senha própria é bem mais curta.
  const ehPrimeiroAcesso = atual.length >= 50;

  if (!ehPrimeiroAcesso) {
    if (!senhaAtual) return { ok: false, message: 'Esta conta já tem senha. Informe a senha atual para trocar.' };
    if (senhaAtual !== atual) return { ok: false, message: 'Senha atual incorreta.' };
  }

  try {
    getSheet_().getRange(u.rowIndex, 3).setValue(novaSenha);
  } catch (err) {
    return { ok: false, message: 'Falha ao gravar na planilha: ' + (err && err.message ? err.message : err) +
      ' — verifique se a implantação executa como o dono e se a planilha é editável.' };
  }
  return { ok: true, message: 'Senha redefinida! Agora faça login com a nova senha.' };
}

/**
 * DIAGNÓSTICO — rode esta função direto no editor do Apps Script
 * (selecione "diagnosticarUsuarios" e clique em ▶ Executar) e veja o
 * resultado em "Registros de execução". Mostra como o papel (role) está
 * gravado na planilha e se conta como admin.
 */
function diagnosticarUsuarios() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  Logger.log('Planilha USERS_SHEET_ID: ' + PropertiesService.getScriptProperties().getProperty('USERS_SHEET_ID'));
  Logger.log('Cabeçalho: ' + JSON.stringify(values[0]));
  let n = 0;
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[1]) continue;
    n++;
    const ehAdmin = String(r[3] || '').trim().toLowerCase() === 'admin';
    const user = String(r[1]);
    const senha = String(r[2] == null ? '' : r[2]);
    Logger.log('username="' + user + '" | role_na_planilha="' + r[3] + '" | admin? ' + ehAdmin
      + ' | user_com_espaco=' + (user !== user.trim())
      + ' | senha_len=' + senha.length + ' | senha_com_espaco=' + (senha !== senha.trim()));
  }
  Logger.log('Total de usuários: ' + n);
  return n;
}

/**
 * INSPECIONAR — mostra o início da senha armazenada de um usuário, para
 * identificar se é texto puro ou um hash. Ajuste USERNAME e rode no editor.
 */
function amostraSenha() {
  const USERNAME = 'Magda'; // <-- ajuste para o usuário que quer inspecionar
  const u = readUsers_().find(x => String(x.username).trim().toLowerCase() === USERNAME.trim().toLowerCase());
  if (!u) { Logger.log('Usuário não encontrado: ' + USERNAME); return; }
  Logger.log('Usuário: ' + u.username + ' | tamanho da senha: ' + u.password.length);
  Logger.log('Início: ' + u.password.substring(0, 30) + '...');
}

/**
 * REDEFINIR SENHA (texto puro) — desbloqueia o login.
 * Ajuste USERNAME e NOVA_SENHA abaixo e rode no editor (▶). Não precisa implantar.
 * Depois entre no app com esse usuário/senha; como admin, use o 🔑 para
 * redefinir a senha dos demais colaboradores.
 */
function redefinirSenha() {
  const USERNAME = 'Magda';          // <-- ajuste
  const NOVA_SENHA = 'Cresol@2024';  // <-- ajuste (texto puro)
  const u = readUsers_().find(x => String(x.username).trim().toLowerCase() === USERNAME.trim().toLowerCase());
  if (!u) throw new Error('Usuário não encontrado: ' + USERNAME);
  getSheet_().getRange(u.rowIndex, 3).setValue(String(NOVA_SENHA));
  Logger.log('Senha de ' + u.username + ' redefinida para texto puro. Já pode logar.');
  return 'OK';
}

/**
 * Mapa público de cores (username -> color), SEM senhas.
 * Usado pelo calendário (inclusive a visão pública) para exibir a cor
 * atual de cada usuário. Os nomes já aparecem nas reservas do calendário,
 * então isto não expõe nenhum dado novo.
 */
function getColorMap() {
  return readUsers_().map(u => ({ username: u.username, color: u.color }));
}

/** Lista usuários (sem senha) — somente admin. */
function listUsers(token) {
  requireAdmin_(token);
  return readUsers_().map(publicUser_);
}

function addUser(token, username, password, role) {
  requireAdmin_(token);
  username = String(username || '').trim();
  password = String(password || '').trim();
  role = (role === 'admin') ? 'admin' : 'user';
  if (!username || !password) throw new Error('Informe nome e senha.');

  if (readUsers_().some(x => x.username.toLowerCase() === username.toLowerCase()))
    throw new Error('Usuário já existe.');

  try {
    getSheet_().appendRow([Utilities.getUuid(), username, password, role, colorFor_(username)]);
  } catch (err) {
    throw new Error('Falha ao gravar na planilha: ' + (err && err.message ? err.message : err) +
      ' — verifique se a implantação executa como VOCÊ (dono) e se a planilha USERS_SHEET_ID é editável pela sua conta.');
  }
  return { ok: true };
}

function updateUserPassword(token, id, newPassword) {
  requireAdmin_(token);
  newPassword = String(newPassword || '').trim();
  if (!newPassword) throw new Error('Senha inválida.');

  const u = readUsers_().find(x => x.id === String(id));
  if (!u) throw new Error('Usuário não encontrado.');
  getSheet_().getRange(u.rowIndex, 3).setValue(newPassword);
  return { ok: true };
}

function updateUserColor(token, id, color) {
  requireAdmin_(token);
  const u = readUsers_().find(x => x.id === String(id));
  if (!u) throw new Error('Usuário não encontrado.');
  getSheet_().getRange(u.rowIndex, 5).setValue(String(color || ''));
  return { ok: true };
}

function deleteUser(token, id) {
  requireAdmin_(token);
  const u = readUsers_().find(x => x.id === String(id));
  if (!u) throw new Error('Usuário não encontrado.');
  getSheet_().deleteRow(u.rowIndex);
  return { ok: true };
}

/* ================= ESPELHO NA PLANILHA (dual-write) =================
   Transição do banco: locais e reservas continuam no Firestore (leitura
   em tempo real), mas cada gravação é espelhada aqui, nas abas "Locais" e
   "Reservas", usando o ID do documento do Firestore como chave. */

function requireSession_(token) {
  const s = getSession_(token);
  if (!s) throw new Error('Sessão expirada. Faça login novamente.');
  return s;
}

function getTab_(name, header) {
  const id = PropertiesService.getScriptProperties().getProperty('USERS_SHEET_ID');
  if (!id) throw new Error('Configure a propriedade USERS_SHEET_ID.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}
function locaisSheet_()  { return getTab_('Locais',  ['firestoreId', 'name']); }
function reservasSheet_() { return getTab_('Reservas', ['firestoreId', 'date', 'placeName', 'period', 'userName', 'userColor']); }

function findRowByKey_(sh, key) {
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(key)) return i + 1;
  return -1;
}

/* ---- Locais ---- */
function mirrorSavePlace(token, firestoreId, name) {
  requireAdmin_(token);
  const sh = locaisSheet_();
  const data = [String(firestoreId), String(name == null ? '' : name)];
  const row = findRowByKey_(sh, firestoreId);
  if (row < 0) sh.appendRow(data); else sh.getRange(row, 1, 1, data.length).setValues([data]);
  return { ok: true };
}
function mirrorDeletePlace(token, firestoreId) {
  requireAdmin_(token);
  const sh = locaisSheet_();
  const row = findRowByKey_(sh, firestoreId);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

/* ---- Reservas ---- */
function mirrorSaveReserva(token, firestoreId, d) {
  requireSession_(token);
  d = d || {};
  const sh = reservasSheet_();
  const data = [String(firestoreId), String(d.date || ''), String(d.placeName || ''),
                String(d.period || ''), String(d.userName || ''), String(d.userColor || '')];
  const row = findRowByKey_(sh, firestoreId);
  if (row < 0) sh.appendRow(data); else sh.getRange(row, 1, 1, data.length).setValues([data]);
  return { ok: true };
}
function mirrorDeleteReserva(token, firestoreId) {
  requireSession_(token);
  const sh = reservasSheet_();
  const row = findRowByKey_(sh, firestoreId);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

/* ---- Backfill / ressincronização completa (admin) ----
   Recebe TODOS os locais e reservas atuais do Firestore (do cliente) e
   regrava as abas do zero. Use para popular a planilha e corrigir diferenças. */
function espelharTudo(token, places, reservas) {
  requireAdmin_(token);
  places = places || [];
  reservas = reservas || [];

  const l = locaisSheet_();
  if (l.getLastRow() > 1) l.getRange(2, 1, l.getLastRow() - 1, l.getLastColumn()).clearContent();
  if (places.length) {
    l.getRange(2, 1, places.length, 2).setValues(places.map(p => [String(p.id), String(p.name || '')]));
  }

  const r = reservasSheet_();
  if (r.getLastRow() > 1) r.getRange(2, 1, r.getLastRow() - 1, r.getLastColumn()).clearContent();
  if (reservas.length) {
    r.getRange(2, 1, reservas.length, 6).setValues(reservas.map(x => [
      String(x.id), String(x.date || ''), String(x.placeName || ''),
      String(x.period || ''), String(x.userName || ''), String(x.userColor || '')
    ]));
  }
  return { ok: true, locais: places.length, reservas: reservas.length };
}
