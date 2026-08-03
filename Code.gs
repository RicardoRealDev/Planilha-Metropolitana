/**
 * CODE.GS COMPLETO — MESA DA REUNIÃO
 * Google Sheets + votação + dashboard + mapa + Zoom/Cloudflare + Plano B.
 *
 * CORREÇÃO PRINCIPAL DESTA VERSÃO
 * Presença final = participantes ativos no Zoom + presenças lançadas no Plano B.
 * Uma fonte não apaga a outra durante a mesma sessão da reunião.
 *
 * INSTALAÇÃO
 * 1. Apague TODO o conteúdo do Code.gs atual.
 * 2. Cole este arquivo inteiro no Code.gs.
 * 3. Não mantenha patches ou funções duplicadas em outros arquivos .gs.
 * 4. Salve o projeto.
 * 5. Recarregue a planilha.
 * 6. Em Implantar > Gerenciar implantações, publique uma nova versão do Web App.
 * 7. Confirme que o Cloudflare usa a URL /exec da implantação atual.
 */

const FASE1_CONFIG = {
  headerRow: 4,
  firstDataRow: 5,
  conselhos: {
    'Araguaína': 'Araguaina',
    'Palmas': 'Palmas',
    'Gurupi': 'Gurupi'
  },
  pautas: {
    'Vice-Presidência': 'Vice',
    'Regimento Interno': 'Reg'
  },
  defaultCandidate: 'Clayzer'
};

const FASE2_CONFIG = {
  aliasSheet: 'ALIASES_ZOOM',
  logZoomSheet: '_LOG_ZOOM',
  logZoomLiveSheet: '_LOG_ZOOM_AO_VIVO',
  zoomActiveSheet: '_ZOOM_ATIVOS',
  pendenciasSheet: '_PENDENCIAS_ZOOM',
  debugSheet: '_DEBUG_PIPEDREAM',
  debugMatchSheet: '_DEBUG_MATCH_ZOOM',
  errosSheet: '_ERROS_ZOOM',
  voteLogSheet: '_LOG_VOTOS',
  snapshotSheet: '_SNAPSHOT_VOTACAO',
  activeMeetingProperty: 'FASE2_ZOOM_REUNIAO_ATIVA',
  zoomSecretProperty: 'FASE2_ZOOM_SECRET_TOKEN',
  zoomUrlTokenProperty: 'FASE2_ZOOM_URL_TOKEN',
  spreadsheetIdProperty: 'FASE2_SPREADSHEET_ID',
  contextoProperty: 'FASE1_CONTEXTO',
  ultimaImportacaoProperty: 'FASE2_ULTIMA_IMPORTACAO',
  presencePrefixes: ['Inst', 'Vice', 'Reg']
};

/** =========================
 *  MENU E PAINEL
 *  ========================= */

function onOpen() {
  rememberSpreadsheetId_();

  SpreadsheetApp.getUi()
    .createMenu('REUNIÃO')
    .addItem('Abrir painel da reunião', 'abrirPainelReuniao')
    .addItem('Abrir painel grande', 'abrirPainelReuniaoGrande')
    .addSeparator()
    .addItem('Abrir votação', 'menuAbrirVotacao')
    .addItem('Todos presentes favoráveis', 'menuTodosFavoraveis')
    .addItem('Conferir resultado', 'menuConferirResultado')
    .addSeparator()
    .addItem('Atualizar dashboard da planilha', 'menuAtualizarDashboardPlanilha')
    .addItem('Abrir Dashboard_Reuniao', 'abrirAbaDashboardReuniao')
    .addSeparator()
    .addItem('Configurar Zoom ao vivo', 'menuConfigurarZoomAoVivo')
    .addItem('Ativar presença ao vivo', 'menuAtivarPresencaAoVivo')
    .addItem('Ver status presença ao vivo', 'menuStatusPresencaAoVivo')
    .addItem('Resetar somente Zoom ao vivo', 'menuResetarPresencaAoVivo')
    .addSeparator()
    .addItem('Abrir aba ALIASES_ZOOM', 'abrirAbaAliasesZoom')
    .addItem('Abrir DEBUG PIPEDREAM', 'abrirAbaDebugPipedream')
    .addItem('Abrir DEBUG MATCH ZOOM', 'abrirAbaDebugMatchZoom')
    .addToUi();
}

function abrirPainelReuniao() {
  rememberSpreadsheetId_();
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Mesa da Reunião')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirPainelReuniaoGrande() {
  rememberSpreadsheetId_();
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(1320)
    .setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Mesa da Reunião');
}

function menuAbrirVotacao() {
  const res = abrirVotacao(getContexto_());
  SpreadsheetApp.getUi().alert(res.mensagem);
}

function menuTodosFavoraveis() {
  const res = todosFavoraveis(getContexto_());
  SpreadsheetApp.getUi().alert(res.mensagem);
}

function menuConferirResultado() {
  const res = obterResumo(getContexto_());
  SpreadsheetApp.getUi().alert(res.textoAta || res.mensagem);
}

function menuAtualizarDashboardPlanilha() {
  const res = atualizarDashboardPlanilha(getContexto_());
  SpreadsheetApp.getUi().alert(res.mensagem);
}

function abrirAbaDashboardReuniao() {
  const ss = getSS_();
  let sheet = ss.getSheetByName('Dashboard_Reuniao');
  if (!sheet) {
    atualizarDashboardPlanilha(getContexto_());
    sheet = ss.getSheetByName('Dashboard_Reuniao');
  }
  sheet.showSheet();
  ss.setActiveSheet(sheet);
}

function menuConfigurarZoomAoVivo() {
  const ui = SpreadsheetApp.getUi();
  rememberSpreadsheetId_();

  const secretResp = ui.prompt(
    'Configurar Zoom ao vivo',
    'Cole o Secret Token do app Zoom. Se o Cloudflare já faz a validação, ele pode ser mantido como backup.',
    ui.ButtonSet.OK_CANCEL
  );
  if (secretResp.getSelectedButton() !== ui.Button.OK) return;

  const secret = String(secretResp.getResponseText() || '').trim();
  if (secret) getProps_().setProperty(FASE2_CONFIG.zoomSecretProperty, secret);

  garantirTokenPublicoWebhook_();

  ui.alert(
    'Configuração salva.\n\n' +
    'Use esta URL na variável APPS_SCRIPT_URL do Cloudflare Worker:\n\n' +
    obterUrlWebhookZoom_() + '\n\n' +
    'No Zoom Marketplace, mantenha a URL pública do Cloudflare Worker.'
  );
}

function menuAtivarPresencaAoVivo() {
  const ui = SpreadsheetApp.getUi();
  const contexto = getContexto_();
  const resp = ui.prompt(
    'Ativar presença ao vivo',
    'Informe o Meeting ID do Zoom para ' + contexto.conselho + '. Pode colar o link ou apenas o número.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const meetingId = extrairMeetingId_(resp.getResponseText());
  if (!meetingId) {
    ui.alert('Meeting ID não identificado.');
    return;
  }

  const res = ativarPresencaAoVivo(Object.assign({}, contexto, { meetingId: meetingId }));
  ui.alert(res.mensagem);
}

function menuStatusPresencaAoVivo() {
  SpreadsheetApp.getUi().alert(formatarStatusAoVivo_(obterStatusPresencaAoVivo()));
}

function menuResetarPresencaAoVivo() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Resetar somente o Zoom ao vivo',
    'Os participantes ativos do Zoom serão zerados. As presenças do Plano B serão preservadas. Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  ui.alert(resetarPresencaAoVivo().mensagem);
}

function abrirAbaAliasesZoom() {
  const sheet = ensureAliasesSheet_();
  sheet.showSheet();
  getSS_().setActiveSheet(sheet);
}

function abrirAbaDebugPipedream() {
  const sheet = ensureSheetWithHeaders_(
    FASE2_CONFIG.debugSheet,
    ['Recebido em', 'Token recebido', 'Evento', 'Meeting ID detectado', 'Nome detectado', 'Body bruto'],
    false
  );
  sheet.showSheet();
  getSS_().setActiveSheet(sheet);
}

function abrirAbaDebugMatchZoom() {
  const sheet = ensureSheetWithHeaders_(
    FASE2_CONFIG.debugMatchSheet,
    ['Data/hora', 'Aba', 'Linha', 'Nome no Zoom', 'Tipo de match', 'Status aplicado', 'Ente', 'Representante votante', 'Titular', 'Suplente'],
    false
  );
  sheet.showSheet();
  getSS_().setActiveSheet(sheet);
}

/** =========================
 *  FUNÇÕES CHAMADAS PELO SIDEBAR
 *  ========================= */

function obterEstadoInicial() {
  rememberSpreadsheetId_();
  return {
    conselhos: Object.keys(FASE1_CONFIG.conselhos),
    pautas: Object.keys(FASE1_CONFIG.pautas),
    votosVice: [FASE1_CONFIG.defaultCandidate, 'Abstenção', 'Branco', 'Ausente', 'Impedido'],
    votosReg: ['Sim', 'Não', 'Abstenção', 'Ausente', 'Impedido'],
    contexto: getContexto_(),
    zoom: obterStatusPresencaAoVivo()
  };
}

function salvarContexto(payload) {
  rememberSpreadsheetId_();
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);
  getProps_().setProperty(FASE2_CONFIG.contextoProperty, JSON.stringify(contexto));
  return { ok: true, mensagem: 'Contexto salvo.', contexto: contexto };
}

function configurarZoomAoVivo(payload) {
  payload = payload || {};
  rememberSpreadsheetId_();

  if (payload.secretToken) {
    getProps_().setProperty(FASE2_CONFIG.zoomSecretProperty, String(payload.secretToken).trim());
  }
  if (payload.publicToken) {
    getProps_().setProperty(FASE2_CONFIG.zoomUrlTokenProperty, String(payload.publicToken).trim());
  }

  garantirTokenPublicoWebhook_();

  return {
    ok: true,
    url: obterUrlWebhookZoom_(),
    mensagem: 'Configuração salva. Use esta URL no Cloudflare Worker, na variável APPS_SCRIPT_URL.'
  };
}

function ativarPresencaAoVivo(payload) {
  rememberSpreadsheetId_();
  payload = payload || {};

  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);
  salvarContexto(contexto);

  const meetingId = extrairMeetingId_(payload.meetingId);
  if (!meetingId) throw new Error('Informe o Meeting ID do Zoom.');

  const registro = {
    ativo: true,
    conselho: contexto.conselho,
    pauta: contexto.pauta,
    candidato: contexto.candidato || FASE1_CONFIG.defaultCandidate,
    meetingId: meetingId,
    iniciadoEm: new Date().toISOString(),
    operador: getUserEmail_()
  };

  getProps_().setProperty(FASE2_CONFIG.activeMeetingProperty, JSON.stringify(registro));
  ensureZoomActiveSheet_();

  // Nova sessão: usa uma nova chave de Plano B e não carrega presenças manuais antigas.
  atualizarPresencaPorAtivosZoom_(registro);

  return {
    ok: true,
    mensagem: 'Presença ao vivo ativada para ' + contexto.conselho + '. Meeting ID: ' + meetingId + '.',
    reuniao: registro
  };
}

function desativarPresencaAoVivo() {
  const active = getReuniaoAoVivoAtiva_();

  if (active && active.meetingId) {
    marcarTodosZoomComoInativos_(active.meetingId, 'DESATIVADA_MANUALMENTE');

    // Recalcula usando apenas o Plano B da sessão. A presença manual é preservada.
    atualizarPresencaPorAtivosZoom_(active);
  }

  getProps_().deleteProperty(FASE2_CONFIG.activeMeetingProperty);

  return {
    ok: true,
    mensagem: 'Presença ao vivo desativada. As presenças lançadas no Plano B foram mantidas.'
  };
}

function resetarPresencaAoVivo() {
  const active = getReuniaoAoVivoAtiva_();
  if (!active || !active.meetingId) {
    return { ok: false, mensagem: 'Nenhuma reunião ao vivo ativa para resetar.' };
  }

  marcarTodosZoomComoInativos_(active.meetingId, 'RESET_MANUAL');
  const atualizacao = atualizarPresencaPorAtivosZoom_(active);

  appendLog_(FASE2_CONFIG.logZoomLiveSheet, [
    'Recebido em', 'Evento Zoom', 'Meeting ID', 'Tópico', 'Nome no Zoom', 'E-mail', 'Evento em', 'Conselho ativo', 'Resultado'
  ], [[
    new Date(), 'RESET_MANUAL', active.meetingId, '', '', '', new Date(), active.conselho,
    'Zoom zerado; Plano B preservado'
  ]]);

  return {
    ok: true,
    atualizacao: atualizacao,
    mensagem: 'Participantes do Zoom zerados. As presenças do Plano B foram preservadas.'
  };
}

function obterStatusPresencaAoVivo() {
  rememberSpreadsheetId_();

  const active = getReuniaoAoVivoAtiva_();
  const publicToken = getProps_().getProperty(FASE2_CONFIG.zoomUrlTokenProperty) || '';
  const secretOk = !!getProps_().getProperty(FASE2_CONFIG.zoomSecretProperty);
  const ativos = active && active.meetingId ? getZoomActiveParticipants_(active.meetingId) : [];
  const pendencias = readPendenciasZoom_();

  let manuais = [];
  if (active) {
    const contexto = contextoDaReuniaoAtiva_(active);
    manuais = lerPresencasManuaisSessao_(contexto, active);
  }

  return {
    ok: true,
    configurado: !!publicToken,
    secretConfigurado: secretOk,
    url: obterUrlWebhookZoom_(),
    reuniaoAtiva: active,
    ativos: ativos,
    totalAtivos: ativos.length,
    presencasManuais: manuais,
    totalManuais: manuais.length,
    pendencias: pendencias
  };
}

/** =========================
 *  WEB APP / CLOUDFLARE / ZOOM
 *  ========================= */

function doGet() {
  rememberSpreadsheetId_();
  return jsonOutput_({
    ok: true,
    servico: 'Mesa da Reunião — Webhook Zoom via Cloudflare',
    planilha: getSS_().getName(),
    dica: 'O Cloudflare Worker deve enviar POST com JSON do Zoom para esta URL.'
  });
}

function doPost(e) {
  const recebidoEm = new Date();
  let raw = '';

  try {
    rememberSpreadsheetId_();

    raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const bodyOriginal = safeJsonParse_(raw);
    const body = unwrapZoomBody_(bodyOriginal);

    appendLog_(FASE2_CONFIG.debugSheet, [
      'Recebido em', 'Token recebido', 'Evento', 'Meeting ID detectado', 'Nome detectado', 'Body bruto'
    ], [[
      recebidoEm,
      e && e.parameter ? String(e.parameter.token || '') : '',
      String(body.event || ''),
      detectarMeetingId_(body),
      detectarNomeParticipante_(body),
      raw
    ]]);

    if (!validarTokenPublicoWebhook_(e)) {
      appendLog_(FASE2_CONFIG.errosSheet, ['Data/hora', 'Erro', 'Conteúdo recebido'], [[
        new Date(), 'Token público inválido no Apps Script', raw
      ]]);
      return jsonOutput_({ ok: false, erro: 'Token público inválido no Apps Script' });
    }

    const resultado = processarEventoZoom_(body);
    return jsonOutput_({ ok: true, origem: 'Apps Script', resultado: resultado });

  } catch (err) {
    appendLog_(FASE2_CONFIG.errosSheet, ['Data/hora', 'Erro', 'Conteúdo recebido'], [[
      new Date(), String(err && err.stack ? err.stack : err), raw
    ]]);
    return jsonOutput_({ ok: false, erro: String(err && err.message ? err.message : err) });
  }
}

function processarEventoZoom_(body) {
  body = body || {};
  const event = String(body.event || '').trim();

  if (event === 'endpoint.url_validation') {
    return responderValidacaoZoom_(body);
  }

  const obj = body.payload && body.payload.object ? body.payload.object : {};
  const participant = obj.participant || body.participant || {};

  const meetingId = normalizarMeetingId_(
    obj.id || obj.uuid || body.id || body.meeting_id || ''
  );

  const topic = String(obj.topic || body.topic || '');
  const nome = String(
    participant.user_name ||
    participant.participant_user_name ||
    participant.name ||
    participant.display_name ||
    body.user_name ||
    ''
  ).trim();

  const email = String(
    participant.email ||
    participant.user_email ||
    body.email ||
    ''
  ).trim();

  const eventoData = dataEventoZoom_(body.event_ts);
  const active = getReuniaoAoVivoAtiva_();

  appendLog_(FASE2_CONFIG.logZoomLiveSheet, [
    'Recebido em', 'Evento Zoom', 'Meeting ID', 'Tópico', 'Nome no Zoom', 'E-mail', 'Evento em', 'Conselho ativo', 'Resultado'
  ], [[
    new Date(), event, meetingId, topic, nome, email, eventoData,
    active ? active.conselho : '', 'Recebido pelo Apps Script'
  ]]);

  if (!active || !active.meetingId) {
    return {
      ok: true,
      ignorado: true,
      motivo: 'Nenhuma reunião ativa configurada na planilha.',
      event: event,
      meetingId: meetingId,
      nome: nome
    };
  }

  const meetingAtivo = normalizarMeetingId_(active.meetingId);
  if (meetingId && meetingId !== meetingAtivo) {
    return {
      ok: true,
      ignorado: true,
      motivo: 'Evento de outro Meeting ID',
      meetingRecebido: meetingId,
      meetingAtivo: meetingAtivo,
      event: event,
      nome: nome
    };
  }

  if (event === 'meeting.started') {
    marcarTodosZoomComoInativos_(meetingAtivo, 'MEETING_STARTED_RESET');
    const atualizacaoInicio = atualizarPresencaPorAtivosZoom_(active);
    return {
      ok: true,
      event: event,
      mensagem: 'Reunião iniciada. Ativos antigos do Zoom foram limpos e o Plano B foi preservado.',
      atualizacao: atualizacaoInicio
    };
  }

  if (event === 'meeting.ended') {
    marcarTodosZoomComoInativos_(meetingAtivo, 'MEETING_ENDED');
    const atualizacaoFim = atualizarPresencaPorAtivosZoom_(active);
    return {
      ok: true,
      event: event,
      mensagem: 'Reunião encerrada. O Zoom foi zerado e o Plano B foi preservado.',
      atualizacao: atualizacaoFim
    };
  }

  if (event === 'meeting.participant_joined') {
    if (!nome) return { ok: true, ignorado: true, motivo: 'Evento de entrada sem nome', event: event };
    registrarParticipanteZoomAtivo_(active, nome, email, 'ENTROU', eventoData);
    return {
      ok: true,
      event: event,
      acao: 'ENTROU',
      participante: nome,
      meetingId: meetingId,
      atualizacao: atualizarPresencaPorAtivosZoom_(active)
    };
  }

  if (event === 'meeting.participant_left') {
    if (!nome) return { ok: true, ignorado: true, motivo: 'Evento de saída sem nome', event: event };
    registrarParticipanteZoomAtivo_(active, nome, email, 'SAIU', eventoData);
    return {
      ok: true,
      event: event,
      acao: 'SAIU',
      participante: nome,
      meetingId: meetingId,
      atualizacao: atualizarPresencaPorAtivosZoom_(active)
    };
  }

  return {
    ok: true,
    ignorado: true,
    motivo: 'Evento recebido, mas não tratado',
    event: event,
    meetingId: meetingId,
    nome: nome
  };
}

/** =========================
 *  CONTROLE DE PARTICIPANTES ATIVOS DO ZOOM
 *  ========================= */

function registrarParticipanteZoomAtivo_(active, nome, email, acao, eventoData) {
  const sheet = ensureZoomActiveSheet_();
  const meetingId = normalizarMeetingId_(active.meetingId);
  const nomeNorm = norm_(nome);
  const last = sheet.getLastRow();
  let row = null;

  if (last >= 2) {
    const values = sheet.getRange(2, 1, last - 1, 10).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (normalizarMeetingId_(values[i][0]) === meetingId && norm_(values[i][1]) === nomeNorm) {
        row = i + 2;
        break;
      }
    }
  }

  if (!row) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 10).setValues([[
      meetingId, nome, email || '', active.conselho, 0, 'Inativo', '', '', '', ''
    ]]);
  }

  const joinCount = Number(sheet.getRange(row, 5).getValue() || 0);
  const now = new Date();

  if (acao === 'ENTROU') {
    sheet.getRange(row, 3).setValue(email || sheet.getRange(row, 3).getValue() || '');
    sheet.getRange(row, 4).setValue(active.conselho || '');
    sheet.getRange(row, 5).setValue(joinCount + 1);
    sheet.getRange(row, 6).setValue('Ativo');
    sheet.getRange(row, 7).setValue(eventoData || now);
    sheet.getRange(row, 9).setValue('ENTROU');
  } else if (acao === 'SAIU') {
    sheet.getRange(row, 6).setValue('Inativo');
    sheet.getRange(row, 8).setValue(eventoData || now);
    sheet.getRange(row, 9).setValue('SAIU');
  }

  sheet.getRange(row, 10).setValue(now);
  SpreadsheetApp.flush();
}

function ensureZoomActiveSheet_() {
  return ensureSheetWithHeaders_(FASE2_CONFIG.zoomActiveSheet, [
    'Meeting ID', 'Nome no Zoom', 'E-mail', 'Conselho', 'Qtd entradas', 'Status',
    'Última entrada', 'Última saída', 'Último evento', 'Atualizado em'
  ], true);
}

function getZoomActiveParticipants_(meetingId) {
  const sheet = ensureZoomActiveSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];

  const mid = normalizarMeetingId_(meetingId);
  const values = sheet.getRange(2, 1, last - 1, 10).getDisplayValues();
  const out = [];
  const seen = {};

  values.forEach(function(r) {
    if (normalizarMeetingId_(r[0]) === mid && norm_(r[5]) === 'ativo') {
      const nome = String(r[1] || '').trim();
      const key = norm_(nome);
      if (nome && !seen[key]) {
        seen[key] = true;
        out.push(nome);
      }
    }
  });

  return out;
}

function marcarTodosZoomComoInativos_(meetingId, evento) {
  const sheet = ensureZoomActiveSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;

  const mid = normalizarMeetingId_(meetingId);
  const values = sheet.getRange(2, 1, last - 1, 10).getDisplayValues();
  const agora = new Date();

  values.forEach(function(r, i) {
    if (normalizarMeetingId_(r[0]) === mid) {
      const row = i + 2;
      sheet.getRange(row, 6).setValue('Inativo');
      sheet.getRange(row, 8).setValue(agora);
      sheet.getRange(row, 9).setValue(evento || 'INATIVO');
      sheet.getRange(row, 10).setValue(agora);
    }
  });

  SpreadsheetApp.flush();
}

function getReuniaoAoVivoAtiva_() {
  const raw = getProps_().getProperty(FASE2_CONFIG.activeMeetingProperty);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function contextoDaReuniaoAtiva_(active) {
  return {
    conselho: active.conselho,
    pauta: active.pauta || 'Regimento Interno',
    candidato: active.candidato || FASE1_CONFIG.defaultCandidate
  };
}

/** =========================
 *  PRESENÇA HÍBRIDA: ZOOM + PLANO B
 *  ========================= */

function chavePresencaManualSessao_(contexto, active) {
  const conselho = String(contexto && contexto.conselho || '').trim();
  const meetingId = active && active.meetingId
    ? normalizarMeetingId_(active.meetingId)
    : 'SEM_MEETING';
  const sessao = active && active.iniciadoEm
    ? String(active.iniciadoEm)
    : meetingId;

  const identificador = (conselho + '_' + meetingId + '_' + sessao)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .slice(0, 180);

  return 'FASE2_PRESENCA_MANUAL_' + identificador;
}

function lerPresencasManuaisSessao_(contexto, active) {
  const raw = getProps_().getProperty(chavePresencaManualSessao_(contexto, active));
  if (!raw) return [];

  try {
    const lista = JSON.parse(raw);
    return Array.isArray(lista) ? lista.filter(Boolean) : [];
  } catch (e) {
    return [];
  }
}

function salvarPresencasManuaisSessao_(contexto, active, participantes) {
  const lista = unirParticipantesPresenca_(participantes || []);
  getProps_().setProperty(
    chavePresencaManualSessao_(contexto, active),
    JSON.stringify(lista)
  );
  return lista;
}

function unirParticipantesPresenca_() {
  const saida = [];
  const vistos = {};

  Array.prototype.slice.call(arguments).forEach(function(lista) {
    (lista || []).forEach(function(nome) {
      const texto = String(nome || '').trim();
      const chave = norm_(texto);
      if (!texto || !chave || vistos[chave]) return;
      vistos[chave] = true;
      saida.push(texto);
    });
  });

  return saida;
}

function reuniaoAtivaDoMesmoConselho_(active, contexto) {
  return !!(
    active &&
    active.meetingId &&
    contexto &&
    norm_(active.conselho) === norm_(contexto.conselho)
  );
}

function importarPresencaZoomColada(payload) {
  payload = payload || {};

  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);
  salvarContexto(contexto);

  const texto = String(payload.listaZoom || '').trim();
  if (!texto) throw new Error('Cole a lista de participantes antes de importar.');

  const novosManuais = parseListaParticipantesZoom_(texto);
  if (!novosManuais.length) throw new Error('Não encontrei nomes válidos na lista colada.');

  ensureAliasesSheet_();
  const aliases = readAliasesZoom_();
  const abas = getAbasPresencaConselho_(contexto);
  if (!abas.length) {
    throw new Error('Não encontrei abas de presença/votação para o conselho selecionado.');
  }

  const activeOriginal = getReuniaoAoVivoAtiva_();
  const active = reuniaoAtivaDoMesmoConselho_(activeOriginal, contexto)
    ? activeOriginal
    : null;

  const manuaisAnteriores = lerPresencasManuaisSessao_(contexto, active);
  const participantesManuais = salvarPresencasManuaisSessao_(
    contexto,
    active,
    unirParticipantesPresenca_(manuaisAnteriores, novosManuais)
  );

  const participantesAoVivo = active
    ? getZoomActiveParticipants_(active.meetingId)
    : [];

  const participantesCombinados = unirParticipantesPresenca_(
    participantesAoVivo,
    participantesManuais
  );

  appendLog_(FASE2_CONFIG.logZoomSheet, [
    'Data/hora', 'Conselho', 'Evento', 'Nome no Zoom', 'Operador'
  ], novosManuais.map(function(nome) {
    return [new Date(), contexto.conselho, 'PRESENCA_MANUAL_ADICIONADA', nome, getUserEmail_()];
  }));

  const matchedNorms = {};
  const detalhesAbas = [];

  abas.forEach(function(sheet) {
    detalhesAbas.push(
      atualizarPresencaEmAba_(sheet, participantesCombinados, aliases, matchedNorms)
    );
  });

  const pendencias = participantesCombinados.filter(function(nome) {
    return !matchedNorms[norm_(nome)];
  });

  writePendenciasZoom_(contexto, pendencias);
  salvarUltimaImportacao_({
    conselho: contexto.conselho,
    participantes: participantesCombinados,
    participantesAoVivo: participantesAoVivo,
    participantesManuais: participantesManuais,
    pendencias: pendencias,
    detalhesAbas: detalhesAbas,
    origem: active ? 'ZOOM_AO_VIVO_MAIS_PLANO_B' : 'PLANO_B'
  });

  return {
    ok: true,
    participantes: participantesCombinados.length,
    participantesAoVivo: participantesAoVivo.length,
    participantesManuais: participantesManuais.length,
    adicionadosAgora: novosManuais.length,
    pendencias: pendencias,
    detalhesAbas: detalhesAbas,
    mensagem:
      'Presença manual adicionada sem alterar a presença ao vivo. ' +
      'Ao vivo: ' + participantesAoVivo.length + '. ' +
      'Plano B: ' + participantesManuais.length + '. ' +
      'Total combinado: ' + participantesCombinados.length + '. ' +
      'Pendências: ' + pendencias.length + '.'
  };
}

function atualizarPresencaPorAtivosZoom_(active) {
  if (!active || !active.meetingId) {
    return {
      participantesAtivos: 0,
      participantesManuais: 0,
      participantesCombinados: 0,
      pendencias: 0,
      detalhesAbas: []
    };
  }

  const contexto = contextoDaReuniaoAtiva_(active);
  const participantesAoVivo = getZoomActiveParticipants_(active.meetingId);
  const participantesManuais = lerPresencasManuaisSessao_(contexto, active);
  const participantesCombinados = unirParticipantesPresenca_(
    participantesAoVivo,
    participantesManuais
  );

  const aliases = readAliasesZoom_();
  const matchedNorms = {};
  const abas = getAbasPresencaConselho_(contexto);
  const detalhesAbas = [];

  abas.forEach(function(sheet) {
    detalhesAbas.push(
      atualizarPresencaEmAba_(sheet, participantesCombinados, aliases, matchedNorms)
    );
  });

  const pendencias = participantesCombinados.filter(function(nome) {
    return !matchedNorms[norm_(nome)];
  });

  writePendenciasZoom_(contexto, pendencias);
  salvarUltimaImportacao_({
    conselho: contexto.conselho,
    participantes: participantesCombinados,
    participantesAoVivo: participantesAoVivo,
    participantesManuais: participantesManuais,
    pendencias: pendencias,
    detalhesAbas: detalhesAbas,
    origem: 'ZOOM_AO_VIVO_MAIS_PLANO_B'
  });

  return {
    participantesAtivos: participantesAoVivo.length,
    participantesManuais: participantesManuais.length,
    participantesCombinados: participantesCombinados.length,
    pendencias: pendencias.length,
    detalhesAbas: detalhesAbas
  };
}

function obterUltimaImportacaoZoom() {
  const raw = getProps_().getProperty(FASE2_CONFIG.ultimaImportacaoProperty);
  if (!raw) {
    return {
      ok: true,
      mensagem: 'Nenhuma importação ou atualização de presença foi executada ainda.',
      pendencias: []
    };
  }

  try {
    return Object.assign({ ok: true }, JSON.parse(raw));
  } catch (e) {
    return { ok: false, mensagem: 'Não foi possível ler a última importação.', pendencias: [] };
  }
}

function limparPresencasManuaisSessaoAtual() {
  const active = getReuniaoAoVivoAtiva_();
  const contexto = active ? contextoDaReuniaoAtiva_(active) : getContexto_();

  getProps_().deleteProperty(chavePresencaManualSessao_(contexto, active));

  if (active && active.meetingId) {
    atualizarPresencaPorAtivosZoom_(active);
  } else {
    const aliases = readAliasesZoom_();
    const matchedNorms = {};
    getAbasPresencaConselho_(contexto).forEach(function(sheet) {
      atualizarPresencaEmAba_(sheet, [], aliases, matchedNorms);
    });
  }

  return {
    ok: true,
    mensagem: 'Presenças manuais da sessão apagadas. A presença ao vivo foi mantida.'
  };
}

function salvarUltimaImportacao_(dados) {
  const registro = Object.assign({ data: new Date().toISOString() }, dados || {});
  getProps_().setProperty(FASE2_CONFIG.ultimaImportacaoProperty, JSON.stringify(registro));
}

/** =========================
 *  ATUALIZAÇÃO DE PRESENÇA NAS ABAS
 *  ========================= */

function getAbasPresencaConselho_(contexto) {
  const ss = getSS_();
  const sufixo = FASE1_CONFIG.conselhos[contexto.conselho];
  if (!sufixo) return [];

  return FASE2_CONFIG.presencePrefixes
    .map(function(prefixo) { return ss.getSheetByName(prefixo + '_' + sufixo); })
    .filter(function(sheet) { return !!sheet; });
}

function atualizarPresencaEmAba_(sheet, participantes, aliases, matchedNorms) {
  const cols = getColunasPresenca_(sheet);
  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  const numRows = Math.max(0, lastRow - FASE1_CONFIG.firstDataRow + 1);

  if (numRows === 0) {
    return { aba: sheet.getName(), presentes: 0, ausentes: 0, atualizada: false };
  }

  const values = sheet
    .getRange(FASE1_CONFIG.firstDataRow, 1, numRows, sheet.getLastColumn())
    .getDisplayValues();

  const presencas = [];
  const debugRows = [];
  let presentes = 0;
  let ausentes = 0;

  values.forEach(function(rowValues, idx) {
    const rowNumber = FASE1_CONFIG.firstDataRow + idx;
    const match = encontrarParticipanteNaLinha_(rowValues, cols, participantes, aliases);

    const ente = cols.ente ? String(rowValues[cols.ente - 1] || '').trim() : '';
    const representante = cols.representante ? String(rowValues[cols.representante - 1] || '').trim() : '';
    const titular = cols.titular ? String(rowValues[cols.titular - 1] || '').trim() : '';
    const suplente = cols.suplente ? String(rowValues[cols.suplente - 1] || '').trim() : '';

    if (match && match.nomeZoom) {
      const status = match.condicao === 'suplente' ? 'Presente suplente' : 'Presente titular';
      presencas.push([status]);
      presentes++;
      matchedNorms[norm_(match.nomeZoom)] = true;

      debugRows.push([
        new Date(), sheet.getName(), rowNumber, match.nomeZoom, match.via, status,
        ente, representante, titular, suplente
      ]);
    } else {
      presencas.push(['Ausente']);
      ausentes++;
    }
  });

  sheet
    .getRange(FASE1_CONFIG.firstDataRow, cols.presenca, numRows, 1)
    .setValues(presencas);

  if (debugRows.length) {
    appendLog_(FASE2_CONFIG.debugMatchSheet, [
      'Data/hora', 'Aba', 'Linha', 'Nome no Zoom', 'Tipo de match', 'Status aplicado',
      'Ente', 'Representante votante', 'Titular', 'Suplente'
    ], debugRows);
  }

  SpreadsheetApp.flush();
  return { aba: sheet.getName(), presentes: presentes, ausentes: ausentes, atualizada: true };
}

function getColunasPresenca_(sheet) {
  const headers = sheet
    .getRange(FASE1_CONFIG.headerRow, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  const map = mapHeaders_(headers);

  function find(labels, required) {
    for (let i = 0; i < labels.length; i++) {
      const key = norm_(labels[i]);
      if (map[key]) return map[key];
    }
    if (required) {
      throw new Error('Coluna não encontrada na aba ' + sheet.getName() + ': ' + labels[0]);
    }
    return null;
  }

  return {
    segmento: find(['Segmento'], false),
    ente: find(['Ente/Órgão ou Município', 'Ente/Órgão', 'Município', 'Ente'], true),
    representante: find(['Representante votante', 'Representante', 'Nome', 'Membro'], false),
    titular: find(['Titular', 'Prefeito(a)/Titular', 'Prefeito/Titular', 'Representante titular', 'Titular/Prefeito'], false),
    suplente: find(['Suplente', 'Representante suplente'], false),
    presenca: find(['Situação de presença', 'Comparecimento'], true)
  };
}

function encontrarParticipanteNaLinha_(rowValues, cols, participantes, aliases) {
  const ente = cols.ente ? String(rowValues[cols.ente - 1] || '').trim() : '';
  const representante = cols.representante ? String(rowValues[cols.representante - 1] || '').trim() : '';
  const titular = cols.titular ? String(rowValues[cols.titular - 1] || '').trim() : '';
  const suplente = cols.suplente ? String(rowValues[cols.suplente - 1] || '').trim() : '';

  for (let i = 0; i < participantes.length; i++) {
    const nomeZoom = String(participantes[i] || '').trim();
    if (!nomeZoom) continue;

    const alias = findAliasForZoomName_(nomeZoom, aliases);

    if (alias) {
      const aliasOficial = String(alias.nomeOficial || '').trim();
      const aliasEnte = String(alias.ente || '').trim();
      const condicaoAlias = norm_(alias.condicao).indexOf('supl') >= 0 ? 'suplente' : 'titular';

      if (aliasEnte && isNameMatch_(aliasEnte, ente)) {
        return { nomeZoom: nomeZoom, condicao: condicaoAlias, via: 'alias-ente' };
      }
      if (aliasOficial && suplente && isNameMatch_(aliasOficial, suplente)) {
        return { nomeZoom: nomeZoom, condicao: 'suplente', via: 'alias-suplente' };
      }
      if (aliasOficial && titular && isNameMatch_(aliasOficial, titular)) {
        return { nomeZoom: nomeZoom, condicao: condicaoAlias, via: 'alias-titular' };
      }
      if (aliasOficial && representante && isNameMatch_(aliasOficial, representante)) {
        return { nomeZoom: nomeZoom, condicao: condicaoAlias, via: 'alias-representante' };
      }
    }

    if (titular && isNameMatch_(nomeZoom, titular)) {
      return { nomeZoom: nomeZoom, condicao: 'titular', via: 'titular' };
    }
    if (suplente && isNameMatch_(nomeZoom, suplente)) {
      return { nomeZoom: nomeZoom, condicao: 'suplente', via: 'suplente' };
    }
    if (representante && isNameMatch_(nomeZoom, representante)) {
      return { nomeZoom: nomeZoom, condicao: 'titular', via: 'representante' };
    }
    if (ente && isEntityNameMatch_(nomeZoom, ente)) {
      return { nomeZoom: nomeZoom, condicao: 'titular', via: 'ente' };
    }
  }

  return null;
}

/** =========================
 *  VOTAÇÃO
 *  ========================= */

function abrirVotacao(payload) {
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);
  salvarContexto(contexto);

  const sheet = getAbaVotacao_(contexto);
  const cols = getColunas_(sheet);
  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  const numRows = Math.max(0, lastRow - FASE1_CONFIG.firstDataRow + 1);
  if (numRows === 0) throw new Error('Não encontrei linhas de membros na aba ' + sheet.getName());

  const data = sheet
    .getRange(FASE1_CONFIG.firstDataRow, 1, numRows, sheet.getLastColumn())
    .getDisplayValues();

  const now = new Date();
  const rows = [];
  let aptos = 0;
  let ausentes = 0;

  data.forEach(function(r, i) {
    const rowNumber = FASE1_CONFIG.firstDataRow + i;
    const status = r[cols.presenca - 1] || '';
    const isPresente = statusEhPresente_(status);

    if (isPresente) aptos++;
    else ausentes++;

    rows.push([
      now,
      contexto.conselho,
      contexto.pauta,
      sheet.getName(),
      rowNumber,
      cols.segmento ? r[cols.segmento - 1] || '' : '',
      r[cols.ente - 1] || '',
      r[cols.representante - 1] || '',
      r[cols.peso - 1] || '',
      status,
      r[cols.voto - 1] || '',
      getUserEmail_()
    ]);
  });

  appendLog_(FASE2_CONFIG.snapshotSheet, [
    'Data/hora', 'Conselho', 'Pauta', 'Aba', 'Linha', 'Segmento', 'Ente',
    'Representante votante', 'Peso', 'Situação de presença', 'Voto no momento', 'Operador'
  ], rows);

  return {
    ok: true,
    aptos: aptos,
    ausentes: ausentes,
    mensagem: 'Votação aberta para ' + contexto.pauta + '. Aptos: ' + aptos + '. Ausentes: ' + ausentes + '.'
  };
}

function todosFavoraveis(payload) {
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);
  const sheet = getAbaVotacao_(contexto);
  const cols = getColunas_(sheet);
  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  const numRows = Math.max(0, lastRow - FASE1_CONFIG.firstDataRow + 1);
  if (numRows === 0) throw new Error('Não encontrei linhas de membros na aba ' + sheet.getName());

  const presencas = sheet.getRange(FASE1_CONFIG.firstDataRow, cols.presenca, numRows, 1).getDisplayValues();
  const entes = sheet.getRange(FASE1_CONFIG.firstDataRow, cols.ente, numRows, 1).getDisplayValues();
  const reps = sheet.getRange(FASE1_CONFIG.firstDataRow, cols.representante, numRows, 1).getDisplayValues();
  const votosAntigos = sheet.getRange(FASE1_CONFIG.firstDataRow, cols.voto, numRows, 1).getDisplayValues();

  const votos = [];
  const logs = [];
  const now = new Date();
  let favoraveis = 0;
  let ausentes = 0;

  for (let i = 0; i < numRows; i++) {
    const presente = statusEhPresente_(presencas[i][0]);
    let voto = 'Ausente';

    if (presente) {
      voto = contexto.pauta === 'Vice-Presidência'
        ? (contexto.candidato || FASE1_CONFIG.defaultCandidate)
        : 'Sim';
      favoraveis++;
    } else {
      ausentes++;
    }

    votos.push([voto]);
    logs.push([
      now, contexto.conselho, contexto.pauta, sheet.getName(), FASE1_CONFIG.firstDataRow + i,
      entes[i][0] || '', reps[i][0] || '', 'TODOS_FAVORAVEIS',
      votosAntigos[i][0] || '', voto, '', getUserEmail_()
    ]);
  }

  sheet.getRange(FASE1_CONFIG.firstDataRow, cols.voto, numRows, 1).setValues(votos);
  appendLog_(FASE2_CONFIG.voteLogSheet, [
    'Data/hora', 'Conselho', 'Pauta', 'Aba', 'Linha', 'Ente', 'Representante',
    'Ação', 'Voto anterior', 'Voto lançado', 'Observação', 'Operador'
  ], logs);
  SpreadsheetApp.flush();

  return {
    ok: true,
    favoraveis: favoraveis,
    ausentes: ausentes,
    mensagem: 'Votos lançados. Favoráveis: ' + favoraveis + '. Ausentes: ' + ausentes + '.',
    resumo: obterResumo(contexto)
  };
}

function buscarRegistros(payload) {
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);

  const termo = norm_(payload && payload.busca);
  if (!termo) return { ok: true, resultados: [] };

  const sheet = getAbaVotacao_(contexto);
  const cols = getColunas_(sheet);
  const colsPresenca = getColunasPresenca_(sheet);
  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  const numRows = Math.max(0, lastRow - FASE1_CONFIG.firstDataRow + 1);
  if (!numRows) return { ok: true, resultados: [] };

  const values = sheet
    .getRange(FASE1_CONFIG.firstDataRow, 1, numRows, sheet.getLastColumn())
    .getDisplayValues();

  const resultados = [];

  values.forEach(function(r, i) {
    const ente = String(r[cols.ente - 1] || '').trim();
    const representante = String(r[cols.representante - 1] || '').trim();
    const titular = colsPresenca.titular ? String(r[colsPresenca.titular - 1] || '').trim() : '';
    const suplente = colsPresenca.suplente ? String(r[colsPresenca.suplente - 1] || '').trim() : '';

    const alvo = norm_([ente, representante, titular, suplente].join(' '));
    if (alvo.indexOf(termo) < 0) return;

    resultados.push({
      row: FASE1_CONFIG.firstDataRow + i,
      ente: ente,
      representante: representante || titular || suplente,
      titular: titular,
      suplente: suplente,
      presenca: String(r[cols.presenca - 1] || '').trim(),
      votoAtual: String(r[cols.voto - 1] || '').trim(),
      peso: String(r[cols.peso - 1] || '').trim(),
      segmento: cols.segmento ? String(r[cols.segmento - 1] || '').trim() : ''
    });
  });

  return { ok: true, resultados: resultados.slice(0, 30) };
}

function registrarExcecao(payload) {
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);

  const sheet = getAbaVotacao_(contexto);
  const cols = getColunas_(sheet);
  const row = Number(payload && payload.row);
  const voto = String(payload && payload.voto || '').trim();
  const observacao = String(payload && payload.observacao || '').trim();

  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  if (!row || row < FASE1_CONFIG.firstDataRow || row > lastRow) {
    throw new Error('Linha inválida para registrar o voto.');
  }
  if (!voto) throw new Error('Informe o voto.');

  const oldVote = sheet.getRange(row, cols.voto).getDisplayValue();
  sheet.getRange(row, cols.voto).setValue(voto);
  if (cols.observacoes && observacao) {
    sheet.getRange(row, cols.observacoes).setValue(observacao);
  }

  const ente = sheet.getRange(row, cols.ente).getDisplayValue();
  const representante = sheet.getRange(row, cols.representante).getDisplayValue();

  appendLog_(FASE2_CONFIG.voteLogSheet, [
    'Data/hora', 'Conselho', 'Pauta', 'Aba', 'Linha', 'Ente', 'Representante',
    'Ação', 'Voto anterior', 'Voto lançado', 'Observação', 'Operador'
  ], [[
    new Date(), contexto.conselho, contexto.pauta, sheet.getName(), row,
    ente, representante, 'EXCECAO', oldVote, voto, observacao, getUserEmail_()
  ]]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    mensagem: 'Voto registrado para ' + (ente || representante) + ': ' + voto + '.',
    resumo: obterResumo(contexto)
  };
}

function obterResumo(payload) {
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);

  const sheet = getAbaVotacao_(contexto);
  let resumo = lerResumo_(sheet);

  if (!resumo.linhas.length) {
    resumo = montarResumoCalculado_(contexto, lerLinhasBaseDashboard_(sheet));
  }

  const textoAta = gerarTextoAta_(contexto, resumo);

  return {
    ok: true,
    aba: sheet.getName(),
    resumo: resumo,
    textoAta: textoAta,
    mensagem: 'Resumo atualizado.'
  };
}

/** =========================
 *  DASHBOARD E MAPA
 *  ========================= */

function obterDashboardReuniao(payload) {
  rememberSpreadsheetId_();
  const contexto = normalizarPayload_(payload);
  validarContexto_(contexto);

  const sheet = getAbaVotacao_(contexto);
  const linhasBase = lerLinhasBaseDashboard_(sheet);
  const resumoPlanilha = lerResumo_(sheet).linhas.length
    ? lerResumo_(sheet)
    : montarResumoCalculado_(contexto, linhasBase);

  const votosMap = {};
  const segmentosMap = {};
  const totais = {
    membros: linhasBase.length,
    pesoTotal: 0,
    presentes: 0,
    presentesTitulares: 0,
    presentesSuplentes: 0,
    ausentes: 0,
    pesoPresente: 0,
    votosValidos: 0,
    pesoFavoravel: 0,
    pesoContrario: 0,
    abstencoesPeso: 0,
    impedidosPeso: 0,
    ausentesPeso: 0,
    votados: 0
  };

  linhasBase.forEach(function(linha) {
    const segmento = linha.segmento || 'Sem segmento';
    const peso = numero_(linha.peso);

    totais.pesoTotal += peso;
    if (linha.presente) {
      totais.presentes++;
      totais.pesoPresente += peso;
      if (linha.suplente) totais.presentesSuplentes++;
      else totais.presentesTitulares++;
    } else {
      totais.ausentes++;
    }

    if (!segmentosMap[segmento]) {
      segmentosMap[segmento] = {
        segmento: segmento,
        membros: 0,
        presentes: 0,
        ausentes: 0,
        pesoTotal: 0,
        pesoPresente: 0
      };
    }

    segmentosMap[segmento].membros++;
    segmentosMap[segmento].pesoTotal += peso;
    if (linha.presente) {
      segmentosMap[segmento].presentes++;
      segmentosMap[segmento].pesoPresente += peso;
    } else {
      segmentosMap[segmento].ausentes++;
    }

    const classe = classificarVotoDashboard_(contexto, linha.voto, norm_(linha.voto));
    if (!votosMap[classe.label]) {
      votosMap[classe.label] = { label: classe.label, peso: 0, quantidade: 0, tipo: classe.tipo };
    }
    votosMap[classe.label].peso += peso;
    votosMap[classe.label].quantidade++;

    if (classe.contaComoVotado) totais.votados++;
    if (classe.valido) totais.votosValidos += peso;
    if (classe.tipo === 'favoravel' || classe.tipo === 'candidato-principal') totais.pesoFavoravel += peso;
    if (classe.tipo === 'contrario') totais.pesoContrario += peso;
    if (classe.tipo === 'abstencao') totais.abstencoesPeso += peso;
    if (classe.tipo === 'impedido') totais.impedidosPeso += peso;
    if (classe.tipo === 'ausente') totais.ausentesPeso += peso;
  });

  const votosSerie = ordenarVotosDashboard_(contexto, votosMap);
  const segmentos = Object.keys(segmentosMap)
    .map(function(k) { return segmentosMap[k]; })
    .sort(function(a, b) { return b.pesoPresente - a.pesoPresente; });

  const quorumInfo = obterQuorumDashboard_(contexto, resumoPlanilha, totais, votosMap);
  const resultadoAutomatico = obterResultadoOficial_(resumoPlanilha) ||
    calcularResultadoVisual_(contexto, resumoPlanilha, totais, quorumInfo);
  const zoom = obterStatusPresencaAoVivo();

  return {
    ok: true,
    atualizadoEm: new Date().toISOString(),
    fonte: 'Dados oficiais lidos da própria aba da votação',
    contexto: contexto,
    aba: sheet.getName(),
    resultado: resultadoAutomatico || 'Sem resultado',
    resumo: resumoPlanilha,
    totais: totais,
    quorum: quorumInfo,
    votosSerie: votosSerie,
    segmentos: segmentos,
    zoom: {
      ativo: !!(zoom && zoom.reuniaoAtiva && zoom.reuniaoAtiva.meetingId),
      meetingId: zoom && zoom.reuniaoAtiva ? zoom.reuniaoAtiva.meetingId : '',
      totalAtivos: zoom ? zoom.totalAtivos : 0,
      totalManuais: zoom ? zoom.totalManuais : 0,
      pendencias: zoom ? zoom.pendencias || [] : []
    },
    ultimosVotos: getUltimosVotosDashboard_(contexto, 8),
    municipios: montarPresencaMunicipiosMapa_(linhasBase),
    mapaAviso: ''
  };
}

function obterDashboardReuniaoComMapa(payload) {
  try {
    return obterDashboardReuniao(payload);
  } catch (err) {
    const contexto = normalizarPayload_(payload);
    return {
      ok: false,
      contexto: contexto,
      municipios: [],
      mapaAviso: 'Não foi possível ler a presença municipal: ' +
        String(err && err.message ? err.message : err)
    };
  }
}

function lerLinhasBaseDashboard_(sheet) {
  const cols = getColunas_(sheet);
  const colsPresenca = getColunasPresenca_(sheet);
  const lastRow = getUltimaLinhaDados_(sheet, cols.ente);
  const numRows = Math.max(0, lastRow - FASE1_CONFIG.firstDataRow + 1);
  if (!numRows) return [];

  const values = sheet
    .getRange(FASE1_CONFIG.firstDataRow, 1, numRows, sheet.getLastColumn())
    .getDisplayValues();

  return values.map(function(r, i) {
    const presenca = String(r[cols.presenca - 1] || '').trim();
    const representante = String(r[cols.representante - 1] || '').trim();
    const titular = colsPresenca.titular ? String(r[colsPresenca.titular - 1] || '').trim() : '';
    const suplenteNome = colsPresenca.suplente ? String(r[colsPresenca.suplente - 1] || '').trim() : '';

    return {
      row: FASE1_CONFIG.firstDataRow + i,
      segmento: cols.segmento ? String(r[cols.segmento - 1] || '').trim() : '',
      ente: String(r[cols.ente - 1] || '').trim(),
      representante: representante || (statusEhSuplente_(presenca) ? suplenteNome : titular),
      titular: titular,
      suplenteNome: suplenteNome,
      peso: r[cols.peso - 1] || '',
      presenca: presenca,
      presente: statusEhPresente_(presenca),
      suplente: statusEhSuplente_(presenca),
      voto: String(r[cols.voto - 1] || '').trim(),
      observacoes: cols.observacoes ? String(r[cols.observacoes - 1] || '').trim() : ''
    };
  });
}

function montarPresencaMunicipiosMapa_(linhasBase) {
  const porMunicipio = {};

  (linhasBase || []).forEach(function(linha) {
    const segmentoNorm = norm_(linha && linha.segmento || '');

    if (segmentoNorm.indexOf('estado') >= 0 && segmentoNorm.indexOf('municip') < 0) return;

    const nome = limparNomeMunicipioMapa_(linha && linha.ente || '');
    const chave = norm_(nome);
    if (!chave) return;

    const registro = {
      nome: nome,
      ente: String(linha.ente || '').trim(),
      segmento: String(linha.segmento || '').trim(),
      representante: String(linha.representante || '').trim(),
      titular: String(linha.titular || '').trim(),
      suplenteNome: String(linha.suplenteNome || '').trim(),
      peso: linha.peso || '',
      presenca: String(linha.presenca || '').trim(),
      presente: !!linha.presente,
      suplente: !!linha.suplente,
      voto: String(linha.voto || '').trim()
    };

    if (!porMunicipio[chave] || registro.presente) {
      porMunicipio[chave] = registro;
    }
  });

  return Object.keys(porMunicipio)
    .map(function(chave) { return porMunicipio[chave]; })
    .sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

function limparNomeMunicipioMapa_(valor) {
  return String(valor || '')
    .replace(/^\s*(munic[ií]pio|prefeitura municipal|prefeitura)\s+(de|do|da|dos|das)\s+/i, '')
    .replace(/\s*[-–—]\s*TO\s*$/i, '')
    .replace(/\s*\/\s*TO\s*$/i, '')
    .trim();
}

function classificarVotoDashboard_(contexto, voto, votoNorm) {
  const candidato = contexto.candidato || FASE1_CONFIG.defaultCandidate;

  if (!votoNorm || votoNorm === 'ausente') {
    return { label: 'Ausente', tipo: 'ausente', valido: false, contaComoVotado: false };
  }
  if (votoNorm === 'impedido') {
    return { label: 'Impedido', tipo: 'impedido', valido: false, contaComoVotado: true };
  }
  if (votoNorm === 'abstencao') {
    return { label: 'Abstenção', tipo: 'abstencao', valido: false, contaComoVotado: true };
  }
  if (votoNorm === 'branco') {
    return { label: 'Branco', tipo: 'branco', valido: false, contaComoVotado: true };
  }

  if (contexto.pauta === 'Regimento Interno') {
    if (votoNorm === 'sim') {
      return { label: 'Sim', tipo: 'favoravel', valido: true, contaComoVotado: true };
    }
    if (votoNorm === 'nao') {
      return { label: 'Não', tipo: 'contrario', valido: true, contaComoVotado: true };
    }
    return { label: voto || 'Outro', tipo: 'outro', valido: true, contaComoVotado: true };
  }

  if (votoNorm === norm_(candidato)) {
    return { label: candidato, tipo: 'candidato-principal', valido: true, contaComoVotado: true };
  }

  return { label: voto || 'Outro candidato', tipo: 'candidato', valido: true, contaComoVotado: true };
}

function ordenarVotosDashboard_(contexto, mapa) {
  const ordemReg = ['Sim', 'Não', 'Abstenção', 'Impedido', 'Ausente'];
  const ordemVice = [contexto.candidato || FASE1_CONFIG.defaultCandidate, 'Branco', 'Abstenção', 'Impedido', 'Ausente'];
  const ordem = contexto.pauta === 'Regimento Interno' ? ordemReg : ordemVice;

  return Object.keys(mapa || {})
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) {
      const ia = ordem.map(norm_).indexOf(norm_(a.label));
      const ib = ordem.map(norm_).indexOf(norm_(b.label));
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return -1;
      if (ib >= 0) return 1;
      return b.peso - a.peso;
    });
}

function obterQuorumDashboard_(contexto, resumo, totais, votosMap) {
  const por = resumo && resumo.porIndicador ? resumo.porIndicador : {};
  let quorumTexto = '';

  Object.keys(por).some(function(k) {
    if (norm_(k).indexOf('quorum') >= 0) {
      quorumTexto = String(por[k] || '');
      return true;
    }
    return false;
  });

  let quorumExigido = numero_(quorumTexto);
  let baseAprovacao = 0;

  if (contexto.pauta === 'Regimento Interno') {
    if (!quorumExigido) quorumExigido = Math.ceil((totais.pesoTotal * 2) / 3);
    baseAprovacao = totais.pesoFavoravel;
  } else {
    if (!quorumExigido) quorumExigido = Math.floor(totais.votosValidos / 2) + 1;
    baseAprovacao = maiorPesoVoto_(votosMap);
  }

  return {
    texto: quorumTexto,
    exigido: quorumExigido,
    baseAprovacao: baseAprovacao,
    percentualAprovacao: quorumExigido
      ? Math.min(100, Math.round((baseAprovacao / quorumExigido) * 100))
      : 0,
    percentualPresenca: totais.pesoTotal
      ? Math.round((totais.pesoPresente / totais.pesoTotal) * 100)
      : 0
  };
}

function maiorPesoVoto_(mapa) {
  let maior = 0;
  Object.keys(mapa || {}).forEach(function(k) {
    const labelNorm = norm_(k);
    if (['ausente', 'abstencao', 'impedido', 'branco'].indexOf(labelNorm) >= 0) return;
    maior = Math.max(maior, numero_(mapa[k].peso));
  });
  return maior;
}

function obterResultadoOficial_(resumo) {
  if (!resumo || !resumo.porIndicador) return '';

  const por = resumo.porIndicador;
  const chaves = Object.keys(por);
  for (let i = 0; i < chaves.length; i++) {
    const n = norm_(chaves[i]);
    if (n === 'resultado automatico' || n.indexOf('resultado automatico') >= 0 || n === 'resultado') {
      const valor = String(por[chaves[i]] || '').trim();
      if (valor) return valor;
    }
  }
  return '';
}

function calcularResultadoVisual_(contexto, resumo, totais, quorum) {
  if (!totais.votados) return 'Sem resultado';

  if (contexto.pauta === 'Regimento Interno') {
    if (quorum && quorum.exigido && totais.pesoFavoravel >= quorum.exigido) return 'Aprovado';
    if (totais.pesoFavoravel > totais.pesoContrario) return 'Em apuração';
    return 'Rejeitado';
  }

  if (quorum && quorum.exigido && quorum.baseAprovacao >= quorum.exigido) return 'Eleito';
  return 'Em apuração';
}

function getUltimosVotosDashboard_(contexto, limit) {
  const sheet = getSS_().getSheetByName(FASE2_CONFIG.voteLogSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = mapHeaders_(headers);
  const last = sheet.getLastRow();
  const start = Math.max(2, last - 100 + 1);
  const values = sheet.getRange(start, 1, last - start + 1, sheet.getLastColumn()).getDisplayValues();
  const out = [];

  for (let i = values.length - 1; i >= 0 && out.length < limit; i--) {
    const r = values[i];
    const conselho = valorPorCabecalho_(r, map, 'Conselho');
    const pauta = valorPorCabecalho_(r, map, 'Pauta');

    if (conselho && norm_(conselho) !== norm_(contexto.conselho)) continue;
    if (pauta && norm_(pauta) !== norm_(contexto.pauta)) continue;

    out.push({
      dataHora: valorPorCabecalho_(r, map, 'Data/hora'),
      ente: valorPorCabecalho_(r, map, 'Ente'),
      representante: valorPorCabecalho_(r, map, 'Representante'),
      acao: valorPorCabecalho_(r, map, 'Ação'),
      voto: valorPorCabecalho_(r, map, 'Voto lançado')
    });
  }

  return out;
}

function atualizarDashboardPlanilha(payload) {
  const data = obterDashboardReuniao(payload || getContexto_());
  const ss = getSS_();
  let sheet = ss.getSheetByName('Dashboard_Reuniao');
  if (!sheet) sheet = ss.insertSheet('Dashboard_Reuniao');

  sheet.showSheet();
  sheet.clear();
  sheet.clearFormats();
  sheet.getCharts().forEach(function(chart) { sheet.removeChart(chart); });

  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 10, 120);
  sheet.setRowHeights(1, 60, 26);

  const azul = '#003f88';
  const azulClaro = '#e8f1ff';
  const cinza = '#f3f4f6';
  const borda = '#d1d5db';

  sheet.getRange('A1:J2').merge()
    .setValue('DASHBOARD DA REUNIÃO\n' + data.contexto.conselho + ' · ' + data.contexto.pauta)
    .setFontWeight('bold')
    .setFontSize(17)
    .setFontColor('#ffffff')
    .setBackground(azul)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true);

  sheet.getRange('A3:J3').merge()
    .setValue(
      'Fonte: aba ' + data.aba + ' · Atualizado em ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    )
    .setFontSize(10)
    .setFontColor('#374151')
    .setBackground(azulClaro)
    .setHorizontalAlignment('center');

  const cards = [
    ['A5:B7', 'Presentes', data.totais.presentes],
    ['C5:D7', 'Peso presente', data.totais.pesoPresente],
    ['E5:F7', 'Favoráveis', data.totais.pesoFavoravel],
    ['G5:H7', 'Zoom ativos', data.zoom.totalAtivos],
    ['I5:J7', 'Resultado', data.resultado]
  ];

  cards.forEach(function(c) {
    const rg = sheet.getRange(c[0]);
    rg.merge()
      .setValue(c[1] + '\n' + c[2])
      .setBackground('#ffffff')
      .setBorder(true, true, true, true, true, true, borda, SpreadsheetApp.BorderStyle.SOLID)
      .setWrap(true)
      .setFontWeight('bold')
      .setFontSize(13)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  });

  const votosStart = 10;
  sheet.getRange(votosStart, 1).setValue('Distribuição de votos').setFontWeight('bold').setFontSize(13);
  sheet.getRange(votosStart + 1, 1, 1, 3)
    .setValues([['Voto', 'Peso', 'Quantidade']])
    .setFontWeight('bold')
    .setBackground(cinza);

  const votosRows = data.votosSerie.length
    ? data.votosSerie.map(function(v) { return [v.label, v.peso, v.quantidade]; })
    : [['Sem votos', 0, 0]];

  sheet.getRange(votosStart + 2, 1, votosRows.length, 3).setValues(votosRows);

  try {
    const chartVotos = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(votosStart + 1, 1, votosRows.length + 1, 2))
      .setPosition(votosStart, 5, 0, 0)
      .setOption('title', 'Distribuição de votos')
      .setOption('pieHole', 0.45)
      .setOption('legend', { position: 'right' })
      .build();
    sheet.insertChart(chartVotos);
  } catch (e) {}

  const segStart = votosStart + votosRows.length + 5;
  sheet.getRange(segStart, 1).setValue('Presença por segmento').setFontWeight('bold').setFontSize(13);
  sheet.getRange(segStart + 1, 1, 1, 4)
    .setValues([['Segmento', 'Membros', 'Presentes', 'Ausentes']])
    .setFontWeight('bold')
    .setBackground(cinza);

  const segRows = data.segmentos.length
    ? data.segmentos.map(function(s) { return [s.segmento, s.membros, s.presentes, s.ausentes]; })
    : [['Sem dados', 0, 0, 0]];

  sheet.getRange(segStart + 2, 1, segRows.length, 4).setValues(segRows);

  try {
    const chartSeg = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(segStart + 1, 1, segRows.length + 1, 4))
      .setPosition(segStart, 6, 0, 0)
      .setOption('title', 'Presença por segmento')
      .setOption('legend', { position: 'bottom' })
      .build();
    sheet.insertChart(chartSeg);
  } catch (e) {}

  const mapStart = segStart + segRows.length + 5;
  sheet.getRange(mapStart, 1).setValue('Presença municipal').setFontWeight('bold').setFontSize(13);
  sheet.getRange(mapStart + 1, 1, 1, 4)
    .setValues([['Município', 'Presença', 'Representante', 'Voto']])
    .setFontWeight('bold')
    .setBackground(cinza);

  const municipioRows = data.municipios.length
    ? data.municipios.map(function(m) { return [m.nome, m.presenca, m.representante, m.voto]; })
    : [['Sem dados', '-', '-', '-']];

  sheet.getRange(mapStart + 2, 1, municipioRows.length, 4).setValues(municipioRows);

  const logStart = mapStart + municipioRows.length + 5;
  sheet.getRange(logStart, 1).setValue('Últimos votos').setFontWeight('bold').setFontSize(13);
  sheet.getRange(logStart + 1, 1, 1, 5)
    .setValues([['Hora', 'Ente', 'Representante', 'Ação', 'Voto']])
    .setFontWeight('bold')
    .setBackground(cinza);

  const logs = data.ultimosVotos.length
    ? data.ultimosVotos.map(function(v) { return [v.dataHora, v.ente, v.representante, v.acao, v.voto]; })
    : [['-', '-', '-', '-', '-']];

  sheet.getRange(logStart + 2, 1, logs.length, 5).setValues(logs);
  sheet.autoResizeColumns(1, 10);
  ss.setActiveSheet(sheet);

  return { ok: true, mensagem: 'Dashboard_Reuniao atualizado.', dashboard: data };
}

/** =========================
 *  ALIASES E PENDÊNCIAS
 *  ========================= */

function ensureAliasesSheet_() {
  const sheet = ensureSheetWithHeaders_(FASE2_CONFIG.aliasSheet, [
    'Nome no Zoom', 'Nome oficial', 'Ente', 'Condição', 'Observação'
  ], false);
  sheet.autoResizeColumns(1, 5);
  return sheet;
}

function readAliasesZoom_() {
  const sheet = ensureAliasesSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];

  const values = sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
  return values.map(function(r) {
    return {
      nomeZoom: String(r[0] || '').trim(),
      nomeOficial: String(r[1] || '').trim(),
      ente: String(r[2] || '').trim(),
      condicao: String(r[3] || '').trim(),
      observacao: String(r[4] || '').trim()
    };
  }).filter(function(a) {
    return a.nomeZoom && norm_(a.nomeZoom).indexOf('ex.:') !== 0;
  });
}

function findAliasForZoomName_(nomeZoom, aliases) {
  const nz = norm_(nomeZoom);
  if (!nz) return null;

  return (aliases || []).find(function(a) {
    const az = norm_(a.nomeZoom);
    if (!az) return false;
    if (nz === az) return true;

    const zoomTokens = meaningfulTokens_(nz);
    const aliasTokens = meaningfulTokens_(az);
    if (zoomTokens.length < 2 || aliasTokens.length < 2) return false;

    const setAlias = {};
    aliasTokens.forEach(function(t) { setAlias[t] = true; });
    const matches = zoomTokens.filter(function(t) { return setAlias[t]; }).length;

    return matches >= 2 && matches / Math.min(zoomTokens.length, aliasTokens.length) >= 0.7;
  }) || null;
}

function writePendenciasZoom_(contexto, pendencias) {
  const sheet = ensureSheetWithHeaders_(FASE2_CONFIG.pendenciasSheet, [
    'Data/hora', 'Conselho', 'Nome no Zoom não identificado', 'Status', 'Orientação'
  ], true);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 5).setValues([[
    'Data/hora', 'Conselho', 'Nome no Zoom não identificado', 'Status', 'Orientação'
  ]]);

  if (pendencias && pendencias.length) {
    const rows = pendencias.map(function(nome) {
      return [new Date(), contexto.conselho, nome, 'Pendente', 'Cadastrar equivalência na aba ALIASES_ZOOM.'];
    });
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function readPendenciasZoom_() {
  const sheet = getSS_().getSheetByName(FASE2_CONFIG.pendenciasSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet
    .getRange(2, 3, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .map(function(r) { return r[0]; })
    .filter(Boolean);
}

/** =========================
 *  RESUMO E ATA
 *  ========================= */

function lerResumo_(sheet) {
  const headerRow = sheet
    .getRange(FASE1_CONFIG.headerRow, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  let indicadorCol = null;
  let resultadoCol = null;

  for (let i = 0; i < headerRow.length; i++) {
    if (norm_(headerRow[i]) === 'indicador') indicadorCol = i + 1;
    if (norm_(headerRow[i]) === 'resultado') resultadoCol = i + 1;
  }

  if (!indicadorCol || !resultadoCol) return { linhas: [], porIndicador: {} };

  const last = Math.min(sheet.getLastRow(), 100);
  if (last <= FASE1_CONFIG.headerRow) return { linhas: [], porIndicador: {} };

  const values = sheet
    .getRange(FASE1_CONFIG.headerRow + 1, indicadorCol, last - FASE1_CONFIG.headerRow, 2)
    .getDisplayValues();

  const linhas = [];
  const porIndicador = {};

  values.forEach(function(r) {
    const indicador = String(r[0] || '').trim();
    const resultado = String(r[1] || '').trim();
    if (indicador) {
      linhas.push({ indicador: indicador, resultado: resultado });
      porIndicador[indicador] = resultado;
    }
  });

  return { linhas: linhas, porIndicador: porIndicador };
}

function montarResumoCalculado_(contexto, linhasBase) {
  let pesoTotal = 0;
  let pesoPresente = 0;
  let favoraveis = 0;
  let contrarios = 0;
  let abstencoes = 0;
  let impedimentos = 0;
  let brancos = 0;
  let validos = 0;
  const candidatos = {};

  (linhasBase || []).forEach(function(linha) {
    const peso = numero_(linha.peso);
    const votoNorm = norm_(linha.voto);
    pesoTotal += peso;
    if (linha.presente) pesoPresente += peso;

    if (contexto.pauta === 'Regimento Interno') {
      if (votoNorm === 'sim') favoraveis += peso;
      else if (votoNorm === 'nao') contrarios += peso;
      else if (votoNorm === 'abstencao') abstencoes += peso;
      else if (votoNorm === 'impedido') impedimentos += peso;
    } else {
      if (votoNorm === 'abstencao') abstencoes += peso;
      else if (votoNorm === 'impedido') impedimentos += peso;
      else if (votoNorm === 'branco') brancos += peso;
      else if (votoNorm && votoNorm !== 'ausente') {
        validos += peso;
        const label = linha.voto;
        candidatos[label] = (candidatos[label] || 0) + peso;
      }
    }
  });

  const linhas = [];

  if (contexto.pauta === 'Regimento Interno') {
    const quorum = Math.ceil((pesoTotal * 2) / 3);
    const resultado = favoraveis >= quorum ? 'Aprovado' : (favoraveis || contrarios ? 'Rejeitado' : 'Sem resultado');
    linhas.push(
      { indicador: 'Total de votos favoráveis', resultado: formatarNumero_(favoraveis) },
      { indicador: 'Total de votos contrários', resultado: formatarNumero_(contrarios) },
      { indicador: 'Total de abstenções', resultado: formatarNumero_(abstencoes) },
      { indicador: 'Total de impedimentos', resultado: formatarNumero_(impedimentos) },
      { indicador: 'Quórum exigido de 2/3 da totalidade', resultado: formatarNumero_(quorum) },
      { indicador: 'Resultado automático', resultado: resultado }
    );
  } else {
    Object.keys(candidatos).forEach(function(nome) {
      linhas.push({ indicador: nome, resultado: formatarNumero_(candidatos[nome]) });
    });
    linhas.push(
      { indicador: 'Total de votos válidos', resultado: formatarNumero_(validos) },
      { indicador: 'Abstenções', resultado: formatarNumero_(abstencoes) },
      { indicador: 'Votos brancos', resultado: formatarNumero_(brancos) },
      { indicador: 'Impedimentos', resultado: formatarNumero_(impedimentos) }
    );
  }

  const porIndicador = {};
  linhas.forEach(function(l) { porIndicador[l.indicador] = l.resultado; });
  return { linhas: linhas, porIndicador: porIndicador };
}

function gerarTextoAta_(contexto, resumo) {
  const r = resumo.porIndicador || {};

  if (contexto.pauta === 'Regimento Interno') {
    const fav = r['Total de votos favoráveis'] || '0';
    const contra = r['Total de votos contrários'] || '0';
    const abst = r['Total de abstenções'] || '0';
    const imp = r['Total de impedimentos'] || '0';
    const quorum = r['Quórum exigido de 2/3 da totalidade'] || '';
    const resultado = r['Resultado automático'] || '';
    const aprovado = norm_(resultado).indexOf('aprovado') >= 0;

    return 'Realizada a votação, a minuta do Regimento Interno recebeu ' + fav +
      ' votos ponderados favoráveis, ' + contra + ' votos ponderados contrários, ' +
      abst + ' abstenções e ' + imp + ' impedimentos. ' +
      (quorum ? 'O quórum qualificado exigido era de ' + quorum + ' votos ponderados. ' : '') +
      'Diante da apuração, o Regimento Interno foi ' +
      (aprovado ? 'aprovado' : 'rejeitado') +
      ', para fins de proclamação pela Presidência.';
  }

  if (contexto.pauta === 'Vice-Presidência') {
    const validos = r['Total de votos válidos'] || '';
    const abst = r['Abstenções'] || '0';
    const brancos = r['Votos brancos'] || '0';
    const imp = r['Impedimentos'] || '0';
    const candidato = contexto.candidato || FASE1_CONFIG.defaultCandidate;
    let votosCandidato = '';

    resumo.linhas.forEach(function(l) {
      if (norm_(l.indicador) === norm_(candidato)) votosCandidato = l.resultado;
    });

    return 'Realizada a votação para a Vice-Presidência, o candidato ' + candidato +
      ' recebeu ' + (votosCandidato || validos || '___') + ' votos ponderados. ' +
      'Foram registradas ' + abst + ' abstenções, ' + brancos +
      ' votos brancos e ' + imp + ' impedimentos. Diante do resultado, fica proclamado ' +
      'o resultado da eleição, conforme apuração da Secretaria-Geral.';
  }

  return '';
}

/** =========================
 *  MATCHING E PARSERS
 *  ========================= */

function unwrapZoomBody_(body) {
  if (body && body.event && body.payload) return body;

  if (body && body.body) {
    if (typeof body.body === 'string') {
      const parsed = safeJsonParse_(body.body);
      if (parsed && parsed.event) return parsed;
    }
    if (typeof body.body === 'object' && body.body.event) return body.body;
  }

  return body || {};
}

function safeJsonParse_(s) {
  try {
    return JSON.parse(s || '{}');
  } catch (e) {
    return {};
  }
}

function detectarMeetingId_(body) {
  body = unwrapZoomBody_(body || {});
  const obj = body.payload && body.payload.object ? body.payload.object : {};
  return normalizarMeetingId_(obj.id || obj.uuid || body.id || body.meeting_id || '');
}

function detectarNomeParticipante_(body) {
  body = unwrapZoomBody_(body || {});
  const obj = body.payload && body.payload.object ? body.payload.object : {};
  const p = obj.participant || body.participant || {};
  return String(
    p.user_name || p.participant_user_name || p.name || p.display_name || body.user_name || ''
  ).trim();
}

function parseListaParticipantesZoom_(texto) {
  const raw = String(texto || '')
    .replace(/\t/g, ' ')
    .split(/\r?\n|;|,/)
    .map(function(x) { return x.replace(/^[-•*\d.)\s]+/, '').trim(); })
    .filter(Boolean);

  const seen = {};
  const out = [];

  raw.forEach(function(nome) {
    const limpo = nome
      .replace(/\s*\(.*?\)\s*/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const key = norm_(limpo);

    if (limpo.length > 1 && !seen[key]) {
      seen[key] = true;
      out.push(limpo);
    }
  });

  return out;
}

function isEntityNameMatch_(nomeZoom, ente) {
  const nz = norm_(nomeZoom);
  const ne = norm_(ente);
  if (!nz || !ne || ne.length < 3) return false;

  const genericos = [
    'governo do estado', 'estado', 'tocantins', 'governo',
    'representante', 'conselho', 'secretaria'
  ];
  if (genericos.indexOf(ne) >= 0) return false;

  if (nz === ne) return true;

  const temMarcadorInstitucional = /prefeitura|municipio|gabinete|camara|secretaria|vereador|prefeito/.test(nz);
  if (temMarcadorInstitucional && nz.indexOf(ne) >= 0) return true;

  const enteTokens = meaningfulTokens_(ne);
  const zoomTokens = meaningfulTokens_(nz);
  if (!enteTokens.length || !zoomTokens.length) return false;

  const setZoom = {};
  zoomTokens.forEach(function(t) { setZoom[t] = true; });
  const matches = enteTokens.filter(function(t) { return setZoom[t]; }).length;

  return matches === enteTokens.length && enteTokens.length >= 1 && temMarcadorInstitucional;
}

function isNameMatch_(a, b) {
  const na = norm_(a);
  const nb = norm_(b);
  if (!na || !nb) return false;
  if (na === nb) return true;

  if (na.length >= 6 && nb.indexOf(na) >= 0) return true;
  if (nb.length >= 6 && na.indexOf(nb) >= 0) return true;

  const ta = meaningfulTokens_(na);
  const tb = meaningfulTokens_(nb);
  if (!ta.length || !tb.length) return false;

  const setB = {};
  tb.forEach(function(t) { setB[t] = true; });
  const matches = ta.filter(function(t) { return setB[t]; }).length;
  const menor = Math.min(ta.length, tb.length);

  if (menor === 1) {
    return matches === 1 && Math.max(na.length, nb.length) >= 6;
  }

  return matches >= 2 && matches / menor >= 0.7;
}

function meaningfulTokens_(value) {
  const stop = {
    de: true, da: true, do: true, das: true, dos: true, e: true,
    prefeitura: true, municipal: true, municipio: true, gabinete: true,
    camara: true, secretaria: true, governo: true, estado: true,
    tocantins: true, to: true, dr: true, dra: true, sr: true, sra: true
  };

  return norm_(value)
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(function(t) { return t && t.length >= 2 && !stop[t]; });
}

/** =========================
 *  UTILITÁRIOS DE PLANILHA
 *  ========================= */

function getSS_() {
  const props = getProps_();
  let id = props.getProperty(FASE2_CONFIG.spreadsheetIdProperty);

  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      id = active.getId();
      props.setProperty(FASE2_CONFIG.spreadsheetIdProperty, id);
      return active;
    }
  } catch (e) {}

  if (id) return SpreadsheetApp.openById(id);

  throw new Error(
    'Não consegui identificar a planilha. Abra a planilha e use REUNIÃO > Abrir painel da reunião uma vez.'
  );
}

function rememberSpreadsheetId_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) getProps_().setProperty(FASE2_CONFIG.spreadsheetIdProperty, ss.getId());
  } catch (e) {}
}

function getProps_() {
  return PropertiesService.getDocumentProperties();
}

function getUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || 'usuário não identificado';
  } catch (e) {
    return 'usuário não identificado';
  }
}

function getContexto_() {
  const raw = getProps_().getProperty(FASE2_CONFIG.contextoProperty);
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (e) {}
  }

  return {
    conselho: 'Palmas',
    pauta: 'Regimento Interno',
    candidato: FASE1_CONFIG.defaultCandidate
  };
}

function normalizarPayload_(payload) {
  payload = payload || {};
  const salvo = getContexto_();

  return {
    conselho: payload.conselho || salvo.conselho || 'Palmas',
    pauta: payload.pauta || salvo.pauta || 'Regimento Interno',
    candidato: payload.candidato || salvo.candidato || FASE1_CONFIG.defaultCandidate
  };
}

function validarContexto_(contexto) {
  if (!FASE1_CONFIG.conselhos[contexto.conselho]) {
    throw new Error('Conselho inválido: ' + contexto.conselho);
  }
  if (!FASE1_CONFIG.pautas[contexto.pauta]) {
    throw new Error('Pauta inválida: ' + contexto.pauta);
  }
}

function getAbaVotacao_(contexto) {
  const ss = getSS_();
  const sufixo = FASE1_CONFIG.conselhos[contexto.conselho];
  const prefixo = FASE1_CONFIG.pautas[contexto.pauta];
  const nome = prefixo + '_' + sufixo;
  const sheet = ss.getSheetByName(nome);
  if (!sheet) throw new Error('Aba não encontrada: ' + nome);
  return sheet;
}

function getColunas_(sheet) {
  const headers = sheet
    .getRange(FASE1_CONFIG.headerRow, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  const map = mapHeaders_(headers);

  function find(labels, required) {
    for (let i = 0; i < labels.length; i++) {
      const key = norm_(labels[i]);
      if (map[key]) return map[key];
    }
    if (required) {
      throw new Error('Coluna não encontrada na aba ' + sheet.getName() + ': ' + labels[0]);
    }
    return null;
  }

  return {
    segmento: find(['Segmento'], false),
    ente: find(['Ente/Órgão ou Município', 'Ente/Órgão', 'Município', 'Ente'], true),
    representante: find(['Representante votante', 'Representante', 'Nome', 'Membro'], true),
    peso: find(['Peso do voto', 'Peso'], true),
    presenca: find(['Situação de presença', 'Comparecimento'], true),
    voto: find(['Voto'], true),
    observacoes: find(['Observações', 'Observação'], false)
  };
}

function mapHeaders_(headers) {
  const map = {};
  (headers || []).forEach(function(h, i) {
    const key = norm_(h);
    if (key) map[key] = i + 1;
  });
  return map;
}

function getUltimaLinhaDados_(sheet, colEnte) {
  const maxRows = sheet.getLastRow();
  if (maxRows < FASE1_CONFIG.firstDataRow) return FASE1_CONFIG.firstDataRow - 1;

  const values = sheet
    .getRange(
      FASE1_CONFIG.firstDataRow,
      colEnte,
      maxRows - FASE1_CONFIG.firstDataRow + 1,
      1
    )
    .getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim()) return FASE1_CONFIG.firstDataRow + i;
  }

  return FASE1_CONFIG.firstDataRow - 1;
}

function ensureSheetWithHeaders_(sheetName, headers, hidden) {
  const ss = getSS_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  if (hidden) {
    try { sheet.hideSheet(); } catch (e) {}
  }

  return sheet;
}

function appendLog_(sheetName, headers, rows) {
  if (!rows || !rows.length) return;

  const hidden = sheetName.charAt(0) === '_' &&
    sheetName !== FASE2_CONFIG.debugSheet &&
    sheetName !== FASE2_CONFIG.debugMatchSheet;

  const sheet = ensureSheetWithHeaders_(sheetName, headers, hidden);
  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);

  SpreadsheetApp.flush();
}

/** =========================
 *  UTILITÁRIOS GERAIS E WEBHOOK
 *  ========================= */

function garantirTokenPublicoWebhook_() {
  let token = getProps_().getProperty(FASE2_CONFIG.zoomUrlTokenProperty);
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    getProps_().setProperty(FASE2_CONFIG.zoomUrlTokenProperty, token);
  }
  return token;
}

function obterUrlWebhookZoom_() {
  let base = '';
  try {
    base = ScriptApp.getService().getUrl() || '';
  } catch (e) {}

  if (!base) return '';
  const token = garantirTokenPublicoWebhook_();
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'token=' + encodeURIComponent(token);
}

function validarTokenPublicoWebhook_(e) {
  const esperado = getProps_().getProperty(FASE2_CONFIG.zoomUrlTokenProperty) || '';
  if (!esperado) return true;

  const recebido = e && e.parameter ? String(e.parameter.token || '') : '';
  return recebido === esperado;
}

function responderValidacaoZoom_(body) {
  const plainToken = String(
    body && body.payload && body.payload.plainToken || ''
  );
  const secret = getProps_().getProperty(FASE2_CONFIG.zoomSecretProperty) || '';

  if (!plainToken || !secret) {
    return { ok: false, erro: 'Plain token ou Secret Token ausente.' };
  }

  const signature = Utilities.computeHmacSha256Signature(plainToken, secret);
  const encryptedToken = signature.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');

  return { plainToken: plainToken, encryptedToken: encryptedToken };
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function extrairMeetingId_(value) {
  const s = String(value || '').trim();
  if (!s) return '';

  const match = s.match(/(?:\/j\/|confno=|meetingId=)?(\d{9,12})/);
  return match ? match[1] : normalizarMeetingId_(s);
}

function normalizarMeetingId_(value) {
  return String(value || '').replace(/\D/g, '');
}

function dataEventoZoom_(eventTs) {
  const n = Number(eventTs || 0);
  if (!n) return new Date();
  return new Date(n < 100000000000 ? n * 1000 : n);
}

function statusEhPresente_(status) {
  const n = norm_(status);
  return n.indexOf('presente') >= 0 && n.indexOf('ausente') < 0;
}

function statusEhSuplente_(status) {
  return norm_(status).indexOf('suplente') >= 0;
}

function numero_(value) {
  if (typeof value === 'number') return value;

  let s = String(value || '').trim();
  if (!s) return 0;

  if (s.indexOf(',') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  }

  const match = s.match(/-?\d+(\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function formatarNumero_(value) {
  const n = Number(value || 0);
  if (Math.abs(n - Math.round(n)) < 0.0001) return String(Math.round(n));
  return n.toFixed(1).replace('.', ',');
}

function valorPorCabecalho_(row, map, label) {
  const col = map[norm_(label)];
  return col ? row[col - 1] : '';
}

function norm_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function formatarStatusAoVivo_(res) {
  let texto = 'Status da presença ao vivo\n\n';
  texto += 'Configurado: ' + (res.configurado ? 'Sim' : 'Não') + '\n';
  texto += 'Secret Token configurado: ' + (res.secretConfigurado ? 'Sim' : 'Não') + '\n';
  texto += 'URL do Apps Script: ' + (res.url || '-') + '\n\n';

  if (res.reuniaoAtiva) {
    texto += 'Conselho ativo: ' + res.reuniaoAtiva.conselho + '\n';
    texto += 'Meeting ID: ' + res.reuniaoAtiva.meetingId + '\n';
    texto += 'Participantes ativos no Zoom: ' + res.totalAtivos + '\n';
    texto += 'Presenças do Plano B: ' + (res.totalManuais || 0) + '\n';
  } else {
    texto += 'Nenhuma reunião ao vivo ativa.\n';
  }

  if (res.pendencias && res.pendencias.length) {
    texto += '\nPendências:\n- ' + res.pendencias.join('\n- ');
  }

  return texto;
}
