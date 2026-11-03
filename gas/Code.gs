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
      username: String(r[1] == null ? '' : r[1]),
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

  const u = readUsers_().find(x => x.username.toLowerCase() === username.toLowerCase());
  if (!u || u.password !== password) return { ok: false, message: 'Usuário ou senha inválidos.' };

  return { ok: true, token: createSession_(u), user: publicUser_(u) };
}

function logout(token) {
  // Token stateless: nada a limpar no servidor; o cliente apaga o token.
  return { ok: true };
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
    Logger.log('username="' + r[1] + '" | role_na_planilha="' + r[3] + '" | conta_como_admin? ' + ehAdmin);
  }
  Logger.log('Total de usuários: ' + n);
  return n;
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
