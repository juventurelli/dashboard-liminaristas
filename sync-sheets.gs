// ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
// 1. Crie um token em: https://github.com/settings/tokens/new
//    Selecione: repo → Contents → Read & Write
// 2. Cole o token abaixo
const GITHUB_TOKEN  = 'SEU_TOKEN_AQUI';
const REPO_OWNER    = 'juventurelli';
const REPO_NAME     = 'dashboard-liminaristas';
const FILE_PATH     = 'data.json';
const BRANCH        = 'main';

// Nomes exatos das abas na planilha
const ABA_S1 = 'Liminaristas';          // ← ajuste se o nome for diferente
const ABA_S2 = 'Req. Adiamento Gravidez'; // ← ajuste se o nome for diferente

// Colunas de etapas que existem na aba s1
const ETAPAS = ['VDBP','CI','INSPSAU','EAP','TACF','PPO','PPRM','PHC','Habilitacao a Matricula'];
// ─────────────────────────────────────────────────────────────────────────────

function syncToGitHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s1 = lerS1(ss);
  const s2 = lerS2(ss);
  const json = JSON.stringify({ s1, s2 }, null, 2);
  atualizarGitHub(json);
}

// ── Lê a aba Liminaristas ────────────────────────────────────────────────────
function lerS1(ss) {
  const sheet = ss.getSheetByName(ABA_S1);
  if (!sheet) { Logger.log('Aba não encontrada: ' + ABA_S1); return []; }

  const [headers, ...rows] = sheet.getDataRange().getValues();
  const h = headers.map(v => v.toString().trim());

  return rows
    .filter(r => val(r, h, 'Nome'))
    .map(r => ({
      insc:             val(r, h, 'Inscrição')       || val(r, h, 'Insc')  || val(r, h, 'insc'),
      nome:             val(r, h, 'Nome'),
      exame_orig:       val(r, h, 'Exame Orig')      || val(r, h, 'Exame Original'),
      exame_part:       val(r, h, 'Exame Part')      || val(r, h, 'Exame Participação') || val(r, h, 'Exame Part.'),
      especialidade:    val(r, h, 'Especialidade'),
      sigla:            val(r, h, 'Sigla'),
      processo:         val(r, h, 'Processo'),
      nas_vagas:        val(r, h, 'Nas Vagas')       || val(r, h, 'NasVagas'),
      status:           val(r, h, 'Status'),
      direito_matricula:val(r, h, 'Direito Matrícula')|| val(r, h, 'Dir. Matrícula') || val(r, h, 'Direito Matricula'),
      telefone:         val(r, h, 'Telefone'),
      email:            val(r, h, 'Email')            || val(r, h, 'E-mail'),
      etapas: Object.fromEntries(
        ETAPAS.map(k => [k, val(r, h, k)])
      )
    }));
}

// ── Lê a aba Req. Adiamento Gravidez ────────────────────────────────────────
function lerS2(ss) {
  const sheet = ss.getSheetByName(ABA_S2);
  if (!sheet) { Logger.log('Aba não encontrada: ' + ABA_S2); return []; }

  const [headers, ...rows] = sheet.getDataRange().getValues();
  const h = headers.map(v => v.toString().trim());

  return rows
    .filter(r => val(r, h, 'Nome'))
    .map(r => ({
      nome:               val(r, h, 'Nome'),
      insc:               val(r, h, 'Inscrição')         || val(r, h, 'Insc'),
      exame:              val(r, h, 'Exame'),
      especialidade:      val(r, h, 'Especialidade'),
      email:              val(r, h, 'Email')              || val(r, h, 'E-mail'),
      direito_reconducao: val(r, h, 'Direito Recondução') || val(r, h, 'Dir. Recondução'),
      situacao_final:     val(r, h, 'Situação Final')     || val(r, h, 'Situacao Final'),
      nas_vagas:          val(r, h, 'Nas Vagas'),
      acao:               val(r, h, 'Ação')               || val(r, h, 'Acao'),
      exame_convocar:     val(r, h, 'Exame Convocar')     || val(r, h, 'Convocar para'),
      status:             val(r, h, 'Status')
    }));
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function val(row, headers, key) {
  const i = headers.indexOf(key);
  if (i < 0) return null;
  const v = row[i];
  if (v === null || v === undefined || v === '') return null;
  return v.toString().trim();
}

function atualizarGitHub(content) {
  const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
  const opts = { headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github.v3+json' }, muteHttpExceptions: true };

  // Busca SHA do arquivo atual (necessário para atualizar)
  const get = UrlFetchApp.fetch(url, opts);
  const sha = get.getResponseCode() === 200 ? JSON.parse(get.getContentText()).sha : null;

  const body = {
    message: 'Sync via Apps Script — ' + new Date().toISOString(),
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: BRANCH,
    ...(sha && { sha })
  };

  const put = UrlFetchApp.fetch(url, { ...opts, method: 'put', contentType: 'application/json', payload: JSON.stringify(body) });
  const code = put.getResponseCode();
  Logger.log(code === 200 || code === 201 ? '✅ GitHub atualizado com sucesso' : '❌ Erro: ' + put.getContentText());
}

// ── Configura gatilho de tempo (rode esta função UMA vez) ────────────────────
function configurarGatilho() {
  // Remove gatilhos antigos da mesma função
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncToGitHub')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Cria gatilho a cada 30 minutos
  ScriptApp.newTrigger('syncToGitHub')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('✅ Gatilho criado: syncToGitHub a cada 30 minutos');
}
