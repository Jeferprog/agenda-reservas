# Agenda Cresol Cooperar — Implantação (Google Apps Script)

O login e o gerenciamento de usuários rodam **no servidor** (Apps Script).
As senhas ficam numa **Planilha Google privada** e **nunca** são enviadas ao
navegador. Locais e reservas continuam no **Firestore** (calendário em tempo
real e visão pública).

## Arquivos

- `gas/Code.gs` — backend (login, sessão por token, CRUD de usuários na planilha).
- `AgendaCresolCooperar.html` — front-end (servido pelo Apps Script).
- `firestore.rules` — regras do Firestore (na raiz do repo).

---

## Passo a passo

### 1. Planilha de usuários
1. Crie uma **Planilha Google** privada (não compartilhe publicamente).
2. Copie o **ID** da planilha (parte da URL entre `/d/` e `/edit`).
3. A aba `Usuarios` e o cabeçalho são criados automaticamente no primeiro uso.
   - Colunas: `id | username | password | role | color`
   - `role` deve ser `admin` ou `user`.
   - Para criar o **primeiro admin**, adicione uma linha manualmente
     (ex.: `id` qualquer, `username`, `senha`, `admin`, cor opcional).

### 2. Projeto Apps Script
1. Crie/abra o projeto GAS.
2. Cole o conteúdo de `gas/Code.gs` no arquivo de script (`Código.gs`).
3. Crie um arquivo HTML chamado **`AgendaCresolCooperar`** (sem extensão no
   editor) e cole o conteúdo de `AgendaCresolCooperar.html`.
4. Em **Configurações do projeto → Propriedades do script**, adicione:
   - `USERS_SHEET_ID` = *(ID da planilha do passo 1)*

### 3. Implantar como App da Web
- **Implantar → Nova implantação → App da Web**
- **Executar como:** Eu (proprietário)
- **Quem tem acesso:** *somente sua organização* (recomendado — restringe o
  acesso a contas internas e impede chamadas externas ao servidor).

### 4. Firebase (locais e reservas)
1. **Authentication → Sign-in method →** ative o provedor **Anônimo**.
2. Publique as regras de `firestore.rules`:
   - `firebase deploy --only firestore:rules` **ou**
   - Console → Firestore → Regras → colar → Publicar.
   - ⚠️ Publique as regras **só depois** de implantar esta versão do app
     (que faz login anônimo), senão o calendário para de carregar.

### 5. Migração e limpeza (importante)
1. Mova os usuários que existiam na coleção `users` do Firestore para a
   planilha (colunas acima).
2. **Exclua a coleção `users` do Firestore** — ela contém senhas em texto puro.
   As regras já bloqueiam o acesso a ela, mas o dado deve ser removido.

---

## Como funciona a sessão
- No login, o servidor valida usuário/senha na planilha e devolve um **token**
  (guardado em `sessionStorage`) válido por 6 horas.
- Ações de administrador (criar/editar/excluir usuário, trocar cor) enviam o
  token; o servidor confere se a sessão é de um **admin** antes de executar.
- As senhas nunca trafegam para o navegador.

## Observações
- Cores de usuário são gravadas em cada reserva no momento da criação, então a
  cor aparece no calendário sem precisar ler a lista de usuários no cliente.
  Mudar a cor de um usuário passa a valer para as **novas** reservas.
- A verificação "local já ocupado" para não-admins continua no cliente
  (Firestore), como na versão original.
