// ============================================================
// ZYLLAH DIGITAL · GOOGLE APPS SCRIPT v3.1
// CORS: 100% JSONP — tudo via doGet, sem fetch POST
// ============================================================
//
// PLANILHA_ID — configurar EXCLUSIVAMENTE via Propriedades do Script:
//   Apps Script → ⚙ Configurações do projeto → Propriedades do script → Adicionar propriedade
//   Chave: PLANILHA_ID   |   Valor: <id da nova planilha>
//
// Enquanto a propriedade não existir, o script usa o fallback hardcoded abaixo
// (planilha atual, pré-v3.1). Para trocar de planilha no futuro, NÃO mexer no código —
// apenas editar o valor em Propriedades do Script.
// ============================================================

const PLANILHA_ID = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID')
  || '1I8QDLkKAZnO808oDU6IR90zup_DnxBryHfpyXm4W4cU';
const EMAIL_GUILHERME = Session.getActiveUser().getEmail();
const MODELO_HAIKU = 'claude-haiku-4-5-20251001';
const MODELO_SONNET = 'claude-sonnet-4-20250514';
const PASTA_APRESENTACOES = 'Apresentações Prospects';

const COL = {
  CLIENTES: { ID:1,NOME:2,ESP:3,PLANO:4,TIPO:5,INICIO:6,VENCIMENTO:7,VALOR:8,PAGAMENTO:9,ENTREGAS:10,MARCOS:11,REUNIAO:12,PROX_PAG:13,NOTAS:14,DATA_ULTIMO_ALERTA:15,STATUS_FINANCEIRO:16 },
  PROSPECTS: { ID:1,NOME:2,ESP:3,RISCO:4,INSTAGRAM:5,SEGUIDORES_IG:6,TEM_SITE:7,GOOGLE_MEU_NEGOCIO:8,STATUS:9,DATA_COMENTARIO:10,DATA_DM:11,GANCHO:12,PROXIMO_PASSO:13,NOTAS:14,APRES_GERADA:15 },
  PAUTA: { ID:1,DATA_PUBLICACAO:2,CLIENTE:3,TEMA:4,PLATAFORMA:5,FORMATO:6,STATUS:7,LEGENDA_GERADA:8,CFM_OK:9,OBSERVACAO_CFM:10,LINK_ARTE:11,NOTAS:12 },
  FINANCEIRO: { MES:1,RECEITA_BRUTA:2,CUSTOS:3,MARGEM_LIQUIDA:4,NF_CLOUDIA_EMITIDA:5,NOTAS:6 },
  DEMANDAS: { ID:1,CRIADA_EM:2,TITULO:3,DESCRICAO:4,ORIGEM:5,REF_ID:6,PRIORIDADE:7,STATUS:8,PRAZO:9,NOTAS:10,ATUALIZADA_EM:11 }
};

// ============================================================
// PONTO DE ENTRADA ÚNICO — doGet com JSONP
// Hub envia: ?action=X&data=JSON_ENCODED&callback=FUNC
// ============================================================

function doGet(e) {
  const action = e.parameter.action || '';
  const cb = e.parameter.callback || '';
  const dataRaw = e.parameter.data ? decodeURIComponent(e.parameter.data) : '{}';

  let resultado;
  try {
    const d = JSON.parse(dataRaw);
    if      (action==='listar')               resultado = listar();
    else if (action==='listarProspects')      resultado = listarProspects();
    else if (action==='listarPauta')          resultado = listarPauta();
    else if (action==='listarFinanceiro')     resultado = listarFinanceiro();
    else if (action==='getRitual')            resultado = getRitual();
    else if (action==='salvar')               resultado = salvar(d);
    else if (action==='remover')              resultado = remover(d.id);
    else if (action==='salvarProspect')       resultado = salvarProspect(d);
    else if (action==='removerProspect')      resultado = removerProspect(d.id);
    else if (action==='salvarPauta')          resultado = salvarPauta(d);
    else if (action==='removerPauta')         resultado = removerPauta(d.id);
    else if (action==='salvarFinanceiro')     resultado = salvarFinanceiro(d);
    else if (action==='salvarRitual')         resultado = salvarRitual(d.ritual||d);
    else if (action === 'gerarComentarioYT') resultado = gerarComentarioYT(d);
    else if (action === 'buscarYouTube') resultado = buscarYoutube(d);
    else if (action === 'gerarPauta') resultado = gerarPauta(d);
    else if (action==='listarDemandas')           resultado = listarDemandas();
    else if (action==='salvarDemanda')            resultado = salvarDemanda(d);
    else if (action==='removerDemanda')           resultado = removerDemanda(d.id);
    else if (action==='gerarApresentacaoProspect') resultado = gerarApresentacaoProspect(d);
    else if (action==='mapearMais10')             resultado = mapearMais10(d);
    else if (action==='salvarTodosProspects')     resultado = salvarTodosProspects(d);
    else resultado = { erro: 'Ação inválida: ' + action };
  } catch(err) {
    resultado = { erro: 'Erro: ' + err.message };
  }

  const json = JSON.stringify(resultado);
  if (cb) {
    return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// doPost mantido como fallback (não usado pelo hub)
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    return doGet({ parameter: { action: d.action, data: encodeURIComponent(JSON.stringify(d)) } });
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({erro:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// SETUP
// ============================================================

function setupPlanilha() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAba(ss,'Clientes',['id','nome','esp','plano','tipo','inicio','vencimento','valor','pagamento','entregas','marcos','reuniao','proxPag','notas','data_ultimo_alerta','status_financeiro']);
  criarAba(ss,'Prospects',['id','nome','esp','risco','instagram','seguidores_ig','tem_site','google_meu_negocio','status','data_comentario','data_dm','gancho','proximo_passo','notas','apresentacao_gerada']);
  criarAba(ss,'Pauta',['id','data_publicacao','cliente','tema','plataforma','formato','status','legenda_gerada','cfm_ok','observacao_cfm','link_arte','notas']);
  criarAba(ss,'Financeiro',['mes','receita_bruta','custos','margem_liquida','nf_cloudia_emitida','notas']);
  criarAba(ss,'Demandas',['id','criada_em','titulo','descricao','origem','ref_id','prioridade','status','prazo','notas','atualizada_em']);
  SpreadsheetApp.getUi().alert('✅ Planilha configurada! Execute configurarApiKey() e depois setupContexto(). Opcional: semearDemandas() para carregar backlog desta sessão.');
}

function setupContexto() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbaCtx(ss,'CTX_Permanente',['chave','valor'],[
    ['empresa.nome','Zyllah Digital'],['empresa.fundador','Guilherme Caetano'],
    ['empresa.cidade','Nova Friburgo, RJ'],['empresa.nicho','Especialistas de saúde'],
    ['plano.essencial.preco','1800'],['plano.presenca.preco','2800'],['plano.autoridade.preco','4500'],
    ['regra.encantarys','Empresa da Fernanda (esposa). Laboratório interno. Não é cliente externo.'],
    ['meta.financeira','R$ 5.000 líquidos/mês com 3–4 clientes'],
  ]);
  criarAbaCtx(ss,'CTX_Evolutivo',['chave','valor','atualizado_em','observacao'],[
    ['empresa.status','pré-receita','2026-04-15','40+ dias de operação'],
    ['cloudia.status','credenciado','2026-04-15','Nível Membro 10% desde 09/04'],
    ['flow.status','rascunho','2026-04-15','Não apresentar a prospects'],
    ['stack.apps_script','v3.0 ativo','2026-04-15','CORS resolvido via JSONP'],
    ['stack.gateway_pagamento','pendente','2026-04-15','Decidir antes do 1º contrato'],
    ['mei.status','pendente','2026-04-15','Necessário para NF Cloudia'],
  ]);
  criarAbaCtx(ss,'CTX_Transitorio',['data','categoria','item','status','nota'],[
    ['2026-04-15','cloudia','Primeiro lead','urgente','Prazo 09/05/2026'],
    ['2026-04-15','financeiro','MEI','urgente','Antes do primeiro contrato'],
    ['2026-04-15','prospect','Dra. Kelly Purger','aquecendo','Comentários feitos em 2 posts'],
    ['2026-04-15','prospect','Dr. Cleonicio Cordeiro','pendente','Comentários não iniciados'],
    ['2026-04-15','prospect','Dra. Fabricia Corbett','pendente','Comentários não iniciados'],
  ]);
  criarAba(ss,'CTX_Arquivo',['data_original','data_arquivamento','categoria','item','status','nota']);
  SpreadsheetApp.getUi().alert('✅ Contexto criado! Execute testarExportacao() para gerar os 3 arquivos no Drive.');
}

function criarAba(ss, nome, cabecalhos) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.getRange(1,1,1,cabecalhos.length).setValues([cabecalhos]).setFontWeight('bold').setBackground('#1A1714').setFontColor('#B8976A');
  }
  return aba;
}

function criarAbaCtx(ss, nome, cabecalhos, dados) {
  const aba = criarAba(ss, nome, cabecalhos);
  if (dados && dados.length && aba.getLastRow()<=1) aba.getRange(2,1,dados.length,dados[0].length).setValues(dados);
  return aba;
}

function configurarApiKey() {
  // DESATIVADA POR SEGURANÇA — nunca hardcode chaves no código.
  // Configure manualmente no editor do Apps Script:
  //   Ícone ⚙ (Configurações do projeto) → Propriedades do script → Adicionar propriedade
  // Chaves necessárias:
  //   - ANTHROPIC_API_KEY  (começa com sk-ant-)
  //   - YT_API_KEY         (começa com AIza)
  Logger.log('Configure ANTHROPIC_API_KEY e YT_API_KEY via Propriedades do Script (Configurações do projeto).');
}

function instalarTriggerOnEdit() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='onEditInstalavel') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditInstalavel').forSpreadsheet(ss).onEdit().create();
  Logger.log('Trigger onEdit instalado.');
}

// ============================================================
// CLIENTES
// ============================================================

function listar() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Clientes');
  if (!s) return {clientes:[]};
  return {clientes: s.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:r[0],nome:r[1],esp:r[2],plano:r[3],tipo:r[4],inicio:fmtIso(r[5]),
    vencimento:fmtIso(r[6]),valor:r[7],pagamento:r[8],entregas:r[9],marcos:r[10],
    reuniao:fmtIso(r[11]),proxPag:fmtIso(r[12]),notas:r[13]
  }))};
}

function salvar(c) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Clientes');
  if (!s) return {erro:'Aba Clientes não encontrada'};
  if (!c.id) c.id = 'c'+Date.now();
  const ent = typeof c.entregas==='object' ? JSON.stringify(c.entregas) : (c.entregas||'[]');
  const mar = typeof c.marcos==='object' ? JSON.stringify(c.marcos) : (c.marcos||'{}');
  const row = [c.id,c.nome,c.esp,c.plano,c.tipo,c.inicio,c.vencimento,c.valor,c.pagamento,ent,mar,c.reuniao||'',c.proxPag||'',c.notas||''];
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===c.id){s.getRange(i+1,1,1,row.length).setValues([row]);return {ok:true,acao:'atualizado'};} }
  s.appendRow(row); return {ok:true,acao:'criado',id:c.id};
}

function remover(id) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Clientes');
  if (!s) return {erro:'Aba não encontrada'};
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===id){s.deleteRow(i+1);return {ok:true};} }
  return {erro:'Não encontrado'};
}

// ============================================================
// PROSPECTS — todos os campos opcionais exceto nome
// ============================================================

function listarProspects() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Prospects');
  if (!s) return {prospects:[]};
  return {prospects: s.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:r[0],nome:r[1],esp:r[2]||'',risco:r[3]||'',instagram:r[4]||'',
    seguidores_ig:r[5]||'',tem_site:r[6]||'',google_meu_negocio:r[7]||'',
    status:r[8]||'novo',data_comentario:fmtIso(r[9]),data_dm:fmtIso(r[10]),
    gancho:r[11]||'',proximo_passo:r[12]||'',notas:r[13]||'',apresentacao_gerada:r[14]||''
  }))};
}

function salvarProspect(p) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Prospects');
  if (!s) return {erro:'Aba Prospects não encontrada'};
  if (!p.id) p.id = 'p'+Date.now();
  const row = [p.id,p.nome||'',p.esp||'',p.risco||'',p.instagram||'',p.seguidores_ig||'',p.tem_site||'',p.google_meu_negocio||'',p.status||'novo',p.data_comentario||'',p.data_dm||'',p.gancho||'',p.proximo_passo||'',p.notas||'',p.apresentacao_gerada||''];
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===p.id){s.getRange(i+1,1,1,row.length).setValues([row]);return {ok:true,acao:'atualizado'};} }
  s.appendRow(row); return {ok:true,acao:'criado',id:p.id};
}

function removerProspect(id) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Prospects');
  if (!s) return {erro:'Aba não encontrada'};
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===id){s.deleteRow(i+1);return {ok:true};} }
  return {erro:'Não encontrado'};
}

// Seed em lote — usado pelo auto-seed do hub na primeira abertura
function salvarTodosProspects(d) {
  const lista = typeof d.prospects==='string' ? JSON.parse(d.prospects) : (d.prospects||[]);
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Prospects');
  if (!s) return {erro:'Aba Prospects não encontrada'};
  let criados=0;
  lista.forEach(p=>{ try{salvarProspect(p);criados++;}catch(e){Logger.log('Erro seed: '+e);} });
  return {ok:true,criados};
}

// ============================================================
// PAUTA
// ============================================================

function listarPauta() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Pauta');
  if (!s) return {pauta:[]};
  return {pauta: s.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:r[0],data_publicacao:fmtIso(r[1]),cliente:r[2],tema:r[3],plataforma:r[4],
    formato:r[5],status:r[6],legenda_gerada:r[7],cfm_ok:r[8],observacao_cfm:r[9],link_arte:r[10],notas:r[11]
  }))};
}

function salvarPauta(p) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Pauta');
  if (!s) return {erro:'Aba Pauta não encontrada'};
  if (!p.id) p.id = 'pt'+Date.now();
  const row = [p.id,p.data_publicacao||'',p.cliente||'',p.tema||'',p.plataforma||'',p.formato||'',p.status||'rascunho',p.legenda_gerada||'',p.cfm_ok||'',p.observacao_cfm||'',p.link_arte||'',p.notas||''];
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===p.id){s.getRange(i+1,1,1,row.length).setValues([row]);return {ok:true,acao:'atualizado'};} }
  s.appendRow(row); return {ok:true,acao:'criado',id:p.id};
}

function removerPauta(id) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Pauta');
  if (!s) return {erro:'Aba não encontrada'};
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===id){s.deleteRow(i+1);return {ok:true};} }
  return {erro:'Não encontrado'};
}

// ============================================================
// FINANCEIRO
// ============================================================

function listarFinanceiro() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Financeiro');
  if (!s) return {financeiro:[]};
  return {financeiro: s.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    mes:r[0],receita_bruta:r[1],custos:r[2],margem_liquida:r[3],nf_cloudia_emitida:r[4],notas:r[5]
  }))};
}

function salvarFinanceiro(f) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Financeiro');
  if (!s) return {erro:'Aba Financeiro não encontrada'};
  const margem = (parseFloat(f.receita_bruta||0)-parseFloat(f.custos||0)).toFixed(2);
  const row = [f.mes,f.receita_bruta||0,f.custos||0,margem,f.nf_cloudia_emitida||'nao',f.notas||''];
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===f.mes){s.getRange(i+1,1,1,row.length).setValues([row]);return {ok:true,acao:'atualizado'};} }
  s.appendRow(row); return {ok:true,acao:'criado'};
}

// ============================================================
// DEMANDAS — backlog de itens não resolvidos (meta-tracking)
// origem: conversa | hub | cliente | prospect | zyllah | externo
// prioridade: urgente | alta | media | baixa
// status: aberta | em_andamento | bloqueada | concluida | descartada
// ============================================================

function listarDemandas() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Demandas');
  if (!s) return {demandas:[]};
  return {demandas: s.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:r[0],criada_em:fmtIso(r[1]),titulo:r[2]||'',descricao:r[3]||'',
    origem:r[4]||'zyllah',ref_id:r[5]||'',prioridade:r[6]||'media',
    status:r[7]||'aberta',prazo:fmtIso(r[8]),notas:r[9]||'',
    atualizada_em:fmtIso(r[10])
  }))};
}

function salvarDemanda(d) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Demandas');
  if (!s) return {erro:'Aba Demandas não encontrada. Rode setupPlanilha().'};
  const agora = new Date();
  if (!d.id) { d.id = 'd'+Date.now(); d.criada_em = d.criada_em || agora; }
  const row = [
    d.id, d.criada_em || agora, d.titulo||'', d.descricao||'',
    d.origem||'zyllah', d.ref_id||'', d.prioridade||'media',
    d.status||'aberta', d.prazo||'', d.notas||'', agora
  ];
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) {
    if (dados[i][0]===d.id) { s.getRange(i+1,1,1,row.length).setValues([row]); return {ok:true,acao:'atualizado',id:d.id}; }
  }
  s.appendRow(row); return {ok:true,acao:'criado',id:d.id};
}

function removerDemanda(id) {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Demandas');
  if (!s) return {erro:'Aba não encontrada'};
  const dados = s.getDataRange().getValues();
  for (let i=1;i<dados.length;i++) { if(dados[i][0]===id){s.deleteRow(i+1);return {ok:true};} }
  return {erro:'Não encontrado'};
}

// Pré-carrega backlog desta sessão (18/04/2026). Só cria se aba está vazia.
// Executar manualmente pelo editor do Apps Script uma vez após setupPlanilha().
function semearDemandas() {
  const s = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName('Demandas');
  if (!s) { SpreadsheetApp.getUi().alert('❌ Aba Demandas não existe. Rode setupPlanilha() primeiro.'); return; }
  if (s.getLastRow() > 1) { SpreadsheetApp.getUi().alert('ℹ️ Aba Demandas já tem dados. Semeação abortada para não duplicar.'); return; }
  const hoje = new Date();
  const backlog = [
    ['Check-up b-roll 5 vídeos','Revisar estado das 10 composições Remotion (TextoAnimado + FeedInstagram). Fotos Loremflickr baixa qualidade — decidir: aceitar, trocar por Unsplash, ou abandonar feed-mock.','zyllah','','alta','aberta','',''],
    ['Localizar simulacros Gemini','Usuário mencionou simulacros realistas mas pasta 05_PRODUCAO_AV/RAW está vazia em A/B/C/D. Pedir caminho correto.','conversa','','media','bloqueada','','Aguarda Guilherme apontar pasta'],
    ['Prospector — prompt-cerco Manus','Prompt anti-alucinação para varredura de presença digital (médicos Nova Friburgo + Petrópolis + Niterói + RJ). Avaliar alternativa gratuita.','zyllah','','alta','aberta','',''],
    ['Vendedor — roleplay objeções','Sessão de treino: Guilherme interpreta médico resistente, Claude rebate até esgotar inseguranças.','zyllah','','media','aberta','',''],
    ['Produtor plano B — posts programados','Copy + design HTML (div para F12 print) para postagens regulares. Claude Design (17/04/2026) resolve design lado.','zyllah','','alta','aberta','',''],
    ['Gestor — módulo relacionamento cliente','Coleta estruturada de info para evolução de cliente (repos./pos./continuidade). Novo stack_gestor.html.','zyllah','','media','aberta','',''],
    ['Integrar stack_followup/juridico/produtor ao Apps Script','Os 3 rascunhos rodam 100% em localStorage — se cache limpar, perde histórico. Migrar para abas da planilha.','zyllah','','media','aberta','',''],
    ['Set de gravação — iluminação + guarda-roupa','Teste 17/04 ficou "YouTuber casual" (luz escura, camisa verde off-brand, moldura colorida). Ajustar luz frontal neutra + figurino sóbrio.','zyllah','','baixa','aberta','',''],
    ['MEI antes do primeiro contrato','Abrir MEI para emitir NF Cloudia (prazo 09/05/2026).','zyllah','','urgente','aberta','2026-05-09',''],
    ['Gateway de pagamento','Decidir antes do primeiro contrato assinado.','zyllah','','alta','aberta','','']
  ];
  const linhas = backlog.map(b => {
    const id = 'd' + Date.now() + Math.floor(Math.random()*1000);
    return [id, hoje, b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], hoje];
  });
  s.getRange(2,1,linhas.length,11).setValues(linhas);
  SpreadsheetApp.getUi().alert('✅ ' + linhas.length + ' demandas pré-carregadas na aba Demandas.');
}

// ============================================================
// RITUAL
// ============================================================

function getRitual() {
  const p = PropertiesService.getScriptProperties().getProperty('RITUAL_DATA');
  return {ritual: p?JSON.parse(p):{}};
}

function salvarRitual(ritual) {
  PropertiesService.getScriptProperties().setProperty('RITUAL_DATA',JSON.stringify(ritual||{}));
  return {ok:true};
}

// ============================================================
// TRIGGERS AUTOMÁTICOS
// ============================================================

function onEdit(e) {}

function onEditInstalavel(e) {
  if (!e||!e.range) return;
  const sheet=e.range.getSheet(), nom=sheet.getName(), col=e.range.getColumn(), row=e.range.getRow(), val=(e.value||'').toString().trim().toLowerCase();
  if (row===1) return;
  try {
    if (nom==='Clientes'&&col===COL.CLIENTES.PAGAMENTO) {
      if(val==='atrasado'){sheet.getRange(row,COL.CLIENTES.DATA_ULTIMO_ALERTA).setValue(new Date());sheet.getRange(row,COL.CLIENTES.STATUS_FINANCEIRO).setValue('d3');}
      else if(val==='em-dia') sheet.getRange(row,COL.CLIENTES.STATUS_FINANCEIRO).setValue('ok');
    }
    if (nom==='Pauta'&&col===COL.PAUTA.STATUS) {
      if(val==='pronto_legenda') gerarLegenda(row);
      if(val==='aprovado') criarEventoCalendar(row,sheet);
    }
    if (nom==='Prospects'&&col===COL.PROSPECTS.STATUS) {
      if(val==='dm_enviada') sheet.getRange(row,COL.PROSPECTS.DATA_DM).setValue(new Date());
    }
    if (nom==='CTX_Evolutivo'&&col===2) {
      sheet.getRange(row,3).setValue(new Date().toISOString().split('T')[0]);
    }
  } catch(err){Logger.log('Erro onEdit: '+err);}
}

function verificacaoDiaria() {
  const ss=SpreadsheetApp.openById(PLANILHA_ID);
  const hoje=new Date(); hoje.setHours(0,0,0,0);
  verificarPagamentos(ss,hoje);
  verificarReunioes(ss,hoje);
  verificarEntregas(ss,hoje);
  arquivarProspectsSemResposta(ss,hoje);
  gerarApresentacoesNovosProspects(ss);
}

function resumoSemanal() {
  const ss=SpreadsheetApp.openById(PLANILHA_ID);
  const hoje=new Date(); hoje.setHours(0,0,0,0);
  const dow=hoje.getDay();
  const diasAteSegunda=dow===0?1:8-dow;
  const proxSeg=new Date(hoje); proxSeg.setDate(hoje.getDate()+diasAteSegunda);
  const proxDom=new Date(proxSeg); proxDom.setDate(proxSeg.getDate()+6);

  let c=`BRIEFING SEMANAL ZYLLAH\n${formatarData(proxSeg)} – ${formatarData(proxDom)}\n${'─'.repeat(50)}\n\n`;
  c+=`★ FOCO: ${sugerirFoco(ss,hoje)}\n\n`;
  c+=`📅 AGENDA\n`;
  try {
    CalendarApp.getDefaultCalendar().getEvents(proxSeg,proxDom).forEach(ev=>{
      if(['[REUNIAO]','[ENTREGA]','[PAGAMENTO]','[URGENTE]','[FOLLOWUP]'].some(t=>ev.getTitle().startsWith(t))) c+=`  ${formatarData(ev.getStartTime())} — ${ev.getTitle()}\n`;
    });
  } catch(e){c+='  (Erro Calendar)\n';}
  c+='\n';

  c+=`⚠️ ALERTAS\n`;
  const abaCl=ss.getSheetByName('Clientes');
  if(abaCl){let ok=true;abaCl.getDataRange().getValues().slice(1).forEach(r=>{const sf=(r[15]||'').toString();if(r[0]&&sf&&sf!=='ok'){c+=`  ${r[1]} — ${sf.toUpperCase()} · R$ ${r[7]}\n`;ok=false;}});if(ok)c+='  Todos em dia.\n';}
  c+='\n';

  c+=`🎯 PROSPECTS\n`;
  const abaPr=ss.getSheetByName('Prospects');
  if(abaPr){const at=['novo','pesquisado','aquecendo','dm_enviada','reuniao'];let tp=false;abaPr.getDataRange().getValues().slice(1).forEach(r=>{const st=(r[8]||'').toLowerCase();if(r[0]&&at.includes(st)){c+=`  ${r[1]} (${r[2]||'—'}) · ${st}\n`;tp=true;}});if(!tp)c+='  Nenhum ativo.\n';}

  c+=`\n${'─'.repeat(50)}\nhttps://docs.google.com/spreadsheets/d/${PLANILHA_ID}\n`;
  enviarEmail('📋 Briefing Semanal Zyllah — '+formatarData(proxSeg),c);
}

function sugerirFoco(ss, hoje) {
  const abaCl=ss.getSheetByName('Clientes');
  if(abaCl){
    for(const r of abaCl.getDataRange().getValues().slice(1)){
      if(!r[0])continue;
      const pag=(r[8]||'').toLowerCase(),sf=(r[15]||'').toString();
      if(pag==='atrasado'||sf.startsWith('d')) return `FINANCEIRO — ${r[1]} com pagamento em atraso.`;
    }
    const am=new Date(hoje); am.setDate(hoje.getDate()+1);
    for(const r of abaCl.getDataRange().getValues().slice(1)){
      if(!r[0]||!r[9])continue;
      try{for(const e of JSON.parse(r[9].toString())){if(e.done)continue;const d=parsarData(e.data);if(d&&(datasIguais(d,hoje)||datasIguais(d,am)))return `PRODUCAO — "${e.desc}" para ${r[1]}.`;}}catch(e){}
    }
  }
  const abaPr=ss.getSheetByName('Prospects');
  if(abaPr){
    for(const r of abaPr.getDataRange().getValues().slice(1)){
      if(!r[0]||(r[8]||'').toLowerCase()!=='dm_enviada')continue;
      const d=parsarData(r[10]);if(!d)continue;
      const dias=Math.floor((hoje-d)/86400000);
      if(dias>=5)return `PROSPECCAO — ${r[1]} sem resposta há ${dias} dias.`;
    }
  }
  return `CONTEUDO — Sem urgências. Produzir e aquecer prospects.`;
}

function lembreteNF() {
  const mes=new Date().getMonth()+1;
  if(![3,6,9,12].includes(mes))return;
  const prazo=new Date(new Date().getFullYear(),new Date().getMonth()+1,15);
  enviarEmail('🧾 Emitir NF Cloudia',`Emitir para financeiro@cloudia.com.br até ${formatarData(prazo)}.\nZyllah Digital`);
}

function verificarPagamentos(ss,hoje) {
  const aba=ss.getSheetByName('Clientes');if(!aba)return;
  const em3=new Date(hoje);em3.setDate(hoje.getDate()+3);
  aba.getDataRange().getValues().slice(1).forEach((r,idx)=>{
    if(!r[0])return;
    const pp=parsarData(r[12]),sf=(r[15]||'').toString().trim(),ult=parsarData(r[14]),pag=(r[8]||'').toLowerCase();
    if(pp&&datasIguais(pp,em3)) enviarEmail('Lembrete pagamento — '+r[1],`${r[1]} vence em 3 dias.\nValor: R$ ${r[7]}\nZyllah Digital`);
    if(pag==='atrasado'||sf.startsWith('d')){
      const dias=ult?Math.floor((hoje-ult)/86400000):0;
      const rowNum=idx+2;
      if(sf==='d3'&&dias>=4){aba.getRange(rowNum,16).setValue('d7');aba.getRange(rowNum,15).setValue(new Date());enviarEmail('⚠️ D+7 — '+r[1],`Pagamento pendente há ${dias} dias.\nZyllah Digital`);}
      else if(sf==='d7'&&dias>=8){aba.getRange(rowNum,16).setValue('d15');aba.getRange(rowNum,15).setValue(new Date());enviarEmail('🛑 D+15 — '+r[1],`Considerar pausa de entregas.\nZyllah Digital`);}
      else if(sf==='d15'&&dias>=16){aba.getRange(rowNum,16).setValue('d30');aba.getRange(rowNum,15).setValue(new Date());enviarEmail('🚨 D+30 — '+r[1],`URGENTE: 30 dias de inadimplência.\nZyllah Digital`);}
    }
  });
}

function verificarReunioes(ss,hoje) {
  const aba=ss.getSheetByName('Clientes');if(!aba)return;
  const am=new Date(hoje);am.setDate(hoje.getDate()+1);
  aba.getDataRange().getValues().slice(1).forEach(r=>{
    if(!r[0])return;
    const re=parsarData(r[11]);
    if(re&&datasIguais(re,am))enviarEmail('📅 Reunião amanhã — '+r[1],`${r[1]} amanhã.\nPlano: ${r[3]}\nZyllah Digital`);
  });
}

function verificarEntregas(ss,hoje) {
  const aba=ss.getSheetByName('Clientes');if(!aba)return;
  const am=new Date(hoje);am.setDate(hoje.getDate()+1);
  aba.getDataRange().getValues().slice(1).forEach(r=>{
    if(!r[0]||!r[9])return;
    try{JSON.parse(r[9].toString()).forEach(e=>{if(e.done)return;const d=parsarData(e.data);if(d&&datasIguais(d,am))enviarEmail('📦 Entrega amanhã — '+r[1]+': '+e.desc,`${e.desc} para ${r[1]} amanhã.\nPrioridade: ${e.prio||'normal'}\nZyllah Digital`);});}catch(e){}
  });
}

function arquivarProspectsSemResposta(ss,hoje) {
  const aba=ss.getSheetByName('Prospects');if(!aba)return;
  aba.getDataRange().getValues().forEach((r,i)=>{
    if(i===0||!r[0])return;
    if((r[8]||'').toLowerCase()!=='dm_enviada')return;
    const d=parsarData(r[10]);if(!d)return;
    if(Math.floor((hoje-d)/86400000)>=7){
      aba.getRange(i+1,9).setValue('arquivado');
      enviarEmail('[FOLLOWUP] Arquivado — '+r[1],`DM sem resposta há 7+ dias.\nInstagram: ${r[4]}\nZyllah Digital`);
    }
  });
}

// ============================================================
// APRESENTAÇÕES (Sonnet) — máx 3 por rodada
// ============================================================

function gerarApresentacoesNovosProspects(ss) {
  const aba=ss.getSheetByName('Prospects');if(!aba)return;
  const dados=aba.getDataRange().getValues();
  let processados=0;
  for(let i=1;i<dados.length&&processados<3;i++){
    const r=dados[i];
    const nome=(r[COL.PROSPECTS.NOME-1]||'').toString().trim();if(!nome)continue;
    const apres=(r[COL.PROSPECTS.APRES_GERADA-1]||'').toString().trim().toLowerCase();
    if(apres==='ok'||apres==='gerando')continue;
    const status=(r[COL.PROSPECTS.STATUS-1]||'novo').toLowerCase();
    if(['novo','arquivado','perdido'].includes(status))continue;
    const prospect={
      nome,esp:(r[2]||'Especialista de saúde').toString().trim(),
      risco:(r[3]||'').toString().trim(),instagram:(r[4]||'').toString().trim(),
      seguidores_ig:r[5]||'',tem_site:(r[6]||'').toString().trim().toLowerCase(),
      google_meu_negocio:(r[7]||'').toString().trim().toLowerCase(),
      status,gancho:(r[11]||'').toString().trim(),notas:(r[13]||'').toString().trim()
    };
    try{
      aba.getRange(i+1,COL.PROSPECTS.APRES_GERADA).setValue('gerando');
      const perfis=determinarPerfis(prospect);
      const textos=gerarTextosApresentacao(prospect,perfis);
      if(!textos){aba.getRange(i+1,COL.PROSPECTS.APRES_GERADA).setValue('erro_api');continue;}
      const pasta=obterOuCriarSubpasta(PASTA_APRESENTACOES,prospect.nome);
      const links=[];
      perfis.forEach((perfil,idx)=>{
        if(!textos[perfil])return;
        const html=montarHTMLSimples(prospect,perfil,textos[perfil]);
        const nomeArq=(idx+1)+'_'+sanitizarNome(prospect.nome)+'_'+perfil+'.html';
        const arq=salvarHTMLDrive(pasta,nomeArq,html);
        links.push({perfil,url:arq.getUrl(),nome:nomeArq});
      });
      aba.getRange(i+1,COL.PROSPECTS.APRES_GERADA).setValue('ok');
      enviarEmailApresentacoes(prospect,links);
      processados++;Utilities.sleep(2000);
    }catch(err){Logger.log('Erro apres '+nome+': '+err);aba.getRange(i+1,COL.PROSPECTS.APRES_GERADA).setValue('erro: '+err.message.substring(0,40));}
  }
}

function determinarPerfis(p) {
  const seg=parseInt(p.seguidores_ig)||0,temSite=p.tem_site==='sim'||p.tem_site==='true';
  const notas=p.notas.toLowerCase();
  const desc=p.risco==='alto'||['ocupad','resist','cetica','desconfi','difícil','nao acredita'].some(s=>notas.includes(s));
  const init=seg<500&&!temSite,cons=seg>2000||(temSite&&seg>500)||p.risco==='baixo';
  const perfis=new Set();
  if(desc)perfis.add('desconfiada');if(init)perfis.add('iniciante');if(cons)perfis.add('consolidada');
  if(perfis.size===0)perfis.add('consolidada');
  if(perfis.size===1){if(perfis.has('consolidada'))perfis.add('desconfiada');else perfis.add('consolidada');}
  return Array.from(perfis);
}

function gerarTextosApresentacao(prospect,perfis) {
  const apiKey=PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if(!apiKey)return null;
  const prompt=`Redator da Zyllah Digital (agência presença digital para saúde).
Prospect: ${prospect.nome} | Esp: ${prospect.esp} | IG: ${prospect.instagram||'—'} | Seguidores: ${prospect.seguidores_ig||'—'}
Site: ${prospect.tem_site||'—'} | GMN: ${prospect.google_meu_negocio||'—'} | Status: ${prospect.status}
Gancho: ${prospect.gancho||'—'} | Notas: ${prospect.notas||'—'}
Perfis: ${perfis.join(', ')}
Gere para cada perfil: capa{pretag,titulo,subtitulo}, problema{tag,titulo,corpo}, solucao{tag,titulo,corpo}, aspiracao{frase}, cta{tag,titulo,corpo}, include_piloto(bool).
Tom: consolidada=par a par; iniciante=encorajador; desconfiada=direto sem promessa (include_piloto:false).
JSON válido apenas, sem markdown.`;
  try{
    const resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{
      method:'post',contentType:'application/json',
      headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({model:MODELO_SONNET,max_tokens:4096,messages:[{role:'user',content:prompt}]}),
      muteHttpExceptions:true
    });
    if(resp.getResponseCode()!==200){Logger.log('Erro API: '+resp.getResponseCode());return null;}
    const txt=JSON.parse(resp.getContentText()).content[0].text.trim().replace(/```json|```/g,'').trim();
    return JSON.parse(txt);
  }catch(err){Logger.log('Erro Claude: '+err);return null;}
}

function montarHTMLSimples(prospect,perfil,t) {
  const lb={consolidada:'Consolidada',iniciante:'Iniciante',desconfiada:'Desconfiada'};
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<title>Zyllah — ${prospect.nome} · ${lb[perfil]||perfil}</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;1,300&family=Jost:wght@300;400&display=swap" rel="stylesheet">
<style>:root{--ink:#1A1714;--gold:#B8976A;--cream:#FAF7F2;}*{box-sizing:border-box;margin:0;padding:0;}body{background:#111;font-family:Jost,sans-serif;font-weight:300;color:var(--cream);padding:60px;max-width:1200px;margin:0 auto;}.slide{margin-bottom:80px;padding:60px;border:1px solid rgba(184,151,106,.15);}h1{font-family:'Cormorant Garamond',serif;font-style:italic;font-size:48px;color:var(--cream);margin-bottom:12px;}.tag{font-size:10px;letter-spacing:.4em;text-transform:uppercase;color:var(--gold);display:block;margin-bottom:16px;}h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-size:36px;color:var(--cream);margin-bottom:16px;}.linha{width:56px;height:1px;background:var(--gold);margin-bottom:24px;}.corpo{font-size:18px;line-height:1.8;color:rgba(255,255,255,.6);}.sub{font-size:16px;color:rgba(255,255,255,.4);margin-top:12px;}.pretag{font-size:10px;letter-spacing:.4em;text-transform:uppercase;color:var(--gold);margin-bottom:20px;display:block;}header{margin-bottom:60px;border-bottom:1px solid rgba(184,151,106,.1);padding-bottom:32px;}.perfil-tag{display:inline-block;font-size:10px;letter-spacing:.3em;padding:4px 12px;border:1px solid rgba(184,151,106,.3);color:var(--gold);margin-top:8px;}</style></head><body>
<header><div class="tag">Zyllah Digital · Apresentação personalizada</div>
<h1>${prospect.nome}</h1><p class="sub">${prospect.esp||'Especialista de saúde'}</p>
<div class="perfil-tag">${lb[perfil]||perfil}</div></header>
<div class="slide"><span class="pretag">${t.capa?.pretag||''}</span><h1>${t.capa?.titulo||''}</h1><p class="sub">${t.capa?.subtitulo||''}</p></div>
<div class="slide"><span class="tag">${t.problema?.tag||''}</span><h2>${t.problema?.titulo||''}</h2><div class="linha"></div><div class="corpo">${t.problema?.corpo||''}</div></div>
<div class="slide"><span class="tag">${t.solucao?.tag||''}</span><h2>${t.solucao?.titulo||''}</h2><div class="linha"></div><div class="corpo">${t.solucao?.corpo||''}</div></div>
<div class="slide" style="text-align:center;"><div class="corpo" style="font-family:'Cormorant Garamond',serif;font-style:italic;font-size:36px;">${t.aspiracao?.frase||''}</div></div>
<div class="slide"><span class="tag">${t.cta?.tag||''}</span><h2>${t.cta?.titulo||''}</h2><div class="linha"></div><div class="corpo">${t.cta?.corpo||''}</div></div>
</body></html>`;
}

// ============================================================
// EXPORTAÇÃO DE CONTEXTO
// ============================================================

function exportarContexto() {
  const ss=SpreadsheetApp.openById(PLANILHA_ID);
  const pasta=obterOuCriarPasta('Zyllah_Contexto');
  [['CTX_Permanente','01_permanente.md'],['CTX_Evolutivo','02_evolutivo.md'],['CTX_Transitorio','03_transitorio.md']].forEach(([abaName,fileName])=>{
    const aba=ss.getSheetByName(abaName);if(!aba)return;
    const dados=aba.getDataRange().getValues().slice(1);
    let md=`# Zyllah — ${abaName.replace('CTX_','')}\n*Gerado: ${new Date().toLocaleString('pt-BR')}*\n\n`;
    dados.forEach(r=>{if(r[0]&&(r[3]||r[2]||'').toString().toLowerCase()!=='arquivado')md+=`- **${r[0]}:** ${r[1]} ${r[2]?'*('+r[2]+')*':''} ${r[3]?'— '+r[3]:''}\n`;});
    salvarArquivoDrive(pasta,fileName,md);
  });
  Logger.log('Contexto exportado. Pasta: Zyllah_Contexto no Drive.');
}

function faxinaMensal() {
  const ss=SpreadsheetApp.openById(PLANILHA_ID);
  const orig=ss.getSheetByName('CTX_Transitorio'),dest=ss.getSheetByName('CTX_Arquivo');
  if(!orig||!dest)return;
  const dados=orig.getDataRange().getValues(),hoje=new Date().toLocaleDateString('pt-BR');
  const manter=[dados[0]],arquivar=[];
  for(let i=1;i<dados.length;i++){(dados[i][3]||'').toLowerCase()==='arquivado'?arquivar.push([dados[i][0],hoje,...dados[i].slice(1)]):manter.push(dados[i]);}
  orig.clearContents();
  if(manter.length)orig.getRange(1,1,manter.length,manter[0].length).setValues(manter);
  if(arquivar.length)dest.getRange(dest.getLastRow()+1,1,arquivar.length,arquivar[0].length).setValues(arquivar);
  Logger.log('Faxina: '+arquivar.length+' itens arquivados.');
}

// ============================================================
// CALENDAR
// ============================================================

function criarEventoCalendar(row,aba) {
  const d=aba.getRange(row,1,1,12).getValues()[0];
  const dp=parsarData(d[1]);if(!dp)return;
  const titulo=`[ENTREGA] ${d[2]||'Cliente'} — ${d[3]||'Entrega'}`;
  try{const ini=new Date(dp);ini.setHours(9,0,0,0);const fim=new Date(dp);fim.setHours(10,0,0,0);CalendarApp.getDefaultCalendar().createEvent(titulo,ini,fim,{description:'Formato: '+(d[5]||'—')});}catch(e){Logger.log('Erro Calendar: '+e);}
}

// ============================================================
// LEGENDA (Haiku)
// ============================================================

function gerarLegenda(row) {
  const ss=SpreadsheetApp.openById(PLANILHA_ID);
  const aba=ss.getSheetByName('Pauta');if(!aba)return;
  const d=aba.getRange(row,1,1,12).getValues()[0];
  const tema=d[3]||'',plat=d[4]||'',fmt=d[5]||'',cli=d[2]||'';
  if(!tema||!plat)return;
  const apiKey=PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if(!apiKey){aba.getRange(row,10).setValue('ERRO: Configure a chave API.');return;}
  const prompt=`Assistente de conteúdo da Zyllah Digital.\nTema: ${tema} | Plataforma: ${plat} | Formato: ${fmt} | Cliente: ${cli}\nGere legenda (autoridade, sem urgência, máx 150 palavras feed / 80 reel, 5 hashtags) e verifique CFM.\nJSON: {"legenda":"...","cfm_ok":"sim|nao|nao_aplicavel","observacao_cfm":"..."}`;
  try{
    const resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{method:'post',contentType:'application/json',headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},payload:JSON.stringify({model:MODELO_HAIKU,max_tokens:1024,messages:[{role:'user',content:prompt}]}),muteHttpExceptions:true});
    if(resp.getResponseCode()!==200){aba.getRange(row,10).setValue('ERRO API: '+resp.getResponseCode());return;}
    const res=JSON.parse(JSON.parse(resp.getContentText()).content[0].text.trim().replace(/```json|```/g,'').trim());
    aba.getRange(row,8).setValue(res.legenda||'');aba.getRange(row,9).setValue(res.cfm_ok||'');aba.getRange(row,10).setValue(res.observacao_cfm||'');aba.getRange(row,7).setValue('legenda_gerada');
    if(res.cfm_ok==='nao')enviarEmail('⚠️ CFM — '+tema,`Problema: ${res.observacao_cfm}\nAcesse a aba Pauta.\nZyllah Digital`);
  }catch(err){aba.getRange(row,10).setValue('ERRO: '+err.message);}
}

// ============================================================
// UTILITÁRIOS
// ============================================================

function enviarEmail(assunto,corpo){try{GmailApp.sendEmail(EMAIL_GUILHERME,assunto,corpo);}catch(e){Logger.log('Erro e-mail: '+e);}}
function parsarData(v){if(!v)return null;if(v instanceof Date)return isNaN(v.getTime())?null:v;const d=new Date(v.toString().trim());return isNaN(d.getTime())?null:d;}
function fmtIso(v){if(!v)return '';if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().split('T')[0];return v.toString();}
function datasIguais(d1,d2){if(!d1||!d2)return false;return d1.getFullYear()===d2.getFullYear()&&d1.getMonth()===d2.getMonth()&&d1.getDate()===d2.getDate();}
function formatarData(d){if(!d)return '—';const dt=new Date(d);return `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;}
function obterOuCriarPasta(nome){const p=DriveApp.getFoldersByName(nome);return p.hasNext()?p.next():DriveApp.createFolder(nome);}
function obterOuCriarSubpasta(raiz,sub){const r=obterOuCriarPasta(raiz);const s=r.getFoldersByName(sub);return s.hasNext()?s.next():r.createFolder(sub);}
function salvarArquivoDrive(pasta,nome,conteudo){const f=pasta.getFilesByName(nome);if(f.hasNext())f.next().setContent(conteudo);else pasta.createFile(nome,conteudo,MimeType.PLAIN_TEXT);}
function salvarHTMLDrive(pasta,nome,conteudo){const f=pasta.getFilesByName(nome);if(f.hasNext()){const a=f.next();a.setContent(conteudo);return a;}return pasta.createFile(nome,conteudo,MimeType.HTML);}
function enviarEmailApresentacoes(prospect,links){const txt=links.map(l=>`[${l.perfil.toUpperCase()}] ${l.nome}\n${l.url}`).join('\n\n');enviarEmail(`📊 Apresentações — ${prospect.nome}`,`${prospect.nome} | ${prospect.esp} | ${prospect.status}\n\n${txt}\n\nZyllah Digital`);}
function sanitizarNome(n){return n.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').substring(0,30);}

// ============================================================
// APRESENTAÇÃO ON-DEMAND — gerada por solicitação do hub
// Fluxo: hub envia prospect+plano → Sonnet gera textos → montarHTML16x9 → Drive → URL
// ============================================================

function gerarApresentacaoProspect(d) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return {erro:'ANTHROPIC_API_KEY não configurada'};

  const p = typeof d.prospect==='string' ? JSON.parse(d.prospect) : d.prospect;
  if (!p || !p.nome) return {erro:'Prospect inválido'};

  const plano    = (d.plano || 'essencial').toLowerCase();
  const piloto   = (d.piloto || 'false') === 'true';
  const obs      = d.obs || '';
  const perfil   = determinarPerfil(p);

  // ── Sonnet: textos personalizados ──
  const textos = gerarTextosApresPersonalizado(p, plano, perfil, obs, apiKey);
  if (!textos) return {erro:'Falha ao gerar textos com Sonnet'};

  // ── Montar HTML 16x9 ──
  const html = montarHTML16x9(p, textos, plano, piloto, perfil);

  // ── Salvar no Drive ──
  const nomeArq = sanitizarNome(p.nome) + '_' + plano + '_' + Date.now() + '.html';
  const pasta   = obterOuCriarSubpasta(PASTA_APRESENTACOES, sanitizarNome(p.nome));
  const arq     = salvarHTMLDrive(pasta, nomeArq, html);

  // Garante que o arquivo seja acessível via link público
  try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}

  return {ok:true, url: arq.getDownloadUrl().replace('/export?', '/view?')
    .replace('export=download&', '')
    || 'https://drive.google.com/file/d/'+arq.getId()+'/view'};
}

// Perfil único (o mais relevante) para geração de textos
function determinarPerfil(p) {
  const seg = parseInt(p.seguidores_ig)||0;
  const notas = (p.notas||'').toLowerCase();
  if (p.risco==='alto'||['ocupad','resist','cetica','desconfi','difícil'].some(s=>notas.includes(s))) return 'desconfiada';
  if (seg >= 3000 || p.tem_site==='sim') return 'consolidada';
  return 'iniciante';
}

function gerarTextosApresPersonalizado(p, plano, perfil, obs, apiKey) {
  const precos = {essencial:'1.800',presenca:'2.800',autoridade:'4.500'};
  const nomePlano = {essencial:'Essencial',presenca:'Presença Completa',autoridade:'Autoridade'};
  const primeiro = p.nome.split(' ')[0];
  const ig = p.instagram ? p.instagram+' · '+p.seguidores_ig+' seguidores · '+p.posts_ig+' posts' : 'Sem Instagram identificado';
  const presencaResumo = [
    p.tem_site==='sim'?'site próprio':'sem site',
    p.google_meu_negocio==='sim'?'Google Meu Negócio':'sem GMN',
    p.doctoralia==='sim'?'Doctoralia':'sem Doctoralia',
    p.linkedin==='sim'?'LinkedIn':'sem LinkedIn'
  ].join(', ');

  const tonMap = {
    desconfiada:'direto e sem exageros — foque em fatos e autonomia; zero promessas vagas',
    iniciante:'encorajador e gentil — mostre que a barreira de entrada é baixa',
    consolidada:'par a par, estratégico — reconheça o que já funciona antes de apontar o gap'
  };

  const prompt = `Você é redator sênior da Zyllah Digital, agência de presença digital para especialistas de saúde em Nova Friburgo RJ.

PROSPECT: ${p.nome} | Especialidade: ${p.esp}
Instagram: ${ig}
Presença: ${presencaResumo} | Nota de presença (0-10): ${p.nota_presenca||'—'}
Gap principal: ${p.gap_principal||'—'}
Gancho identificado: ${p.gancho||'—'}

PLANO SELECIONADO: ${nomePlano[plano]} (R$ ${precos[plano]}/mês)
PERFIL DO PROSPECT: ${perfil}
TOM: ${tonMap[perfil]||tonMap.consolidada}
${obs?'OBSERVAÇÕES ADICIONAIS: '+obs:''}

Gere os textos para os slides da apresentação personalizada. REGRA PRINCIPAL: slide02b é o diagnóstico gentil mas impactante — ${primeiro} deve se reconhecer na situação e sentir curiosidade, não constrangimento.

Retorne APENAS JSON válido (sem markdown):
{
  "slide02": {
    "tag": "string curta (3-5 palavras em caps)",
    "titulo": "título em 1-2 linhas que captura o problema central de presença digital para especialistas de ${p.esp}",
    "corpo": "2-4 linhas descrevendo a realidade que ${primeiro} vive — sem mencionar a Zyllah ainda"
  },
  "slide02b": {
    "titulo_principal": "frase de 2-3 linhas no formato: '${primeiro}, encontramos [o que está impedindo/a razão pela qual]...' — use os dados reais",
    "cenario_atual": "3-4 linhas: descrição factual dos dados reais acima (mencione números concretos quando disponíveis)",
    "lacunas": "3-4 linhas: o que está ausente ou abaixo do potencial, baseado no gap_principal",
    "potencial": "3-4 linhas: o que seria possível alcançar com uma estratégia estruturada — concreto e realista"
  },
  "slide03": {
    "tag": "string curta (3-5 palavras)",
    "titulo": "título de como a Zyllah resolve o gap identificado",
    "corpo": "3 pontos no formato '<strong>Label:</strong> descrição' separados por <br>"
  },
  "aspiracao": "frase poética de 2 linhas personalizada para ${p.esp} — use <em> na parte mais inspiracional",
  "cta": {
    "proximos": "4 próximos passos no formato '<strong>01 · Label</strong> — descrição' separados por <br>"
  }
}`;

  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:'post', contentType:'application/json',
      headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({
        model:'claude-sonnet-4-5',
        max_tokens:2048,
        messages:[{role:'user',content:prompt}]
      }),
      muteHttpExceptions:true
    });
    if (resp.getResponseCode()!==200) { Logger.log('Erro API '+resp.getResponseCode()+': '+resp.getContentText()); return null; }
    const txt = JSON.parse(resp.getContentText()).content[0].text.trim().replace(/```json\n?|```/g,'').trim();
    return JSON.parse(txt);
  } catch(err) { Logger.log('Erro Sonnet apres: '+err); return null; }
}

function montarHTML16x9(p, t, plano, includePiloto, perfil) {
  const primeiro = p.nome.split(' ')[0];
  const ig = p.instagram || '';
  const cidade = 'Nova Friburgo';
  const nomePlano = {essencial:'Essencial',presenca:'Presença Completa',autoridade:'Autoridade'};
  const precos = {essencial:'R$ 1.800',presenca:'R$ 2.800',autoridade:'R$ 4.500'};
  const impl   = {essencial:'+ R$ 1.200 de implantação',presenca:'+ R$ 1.800 de implantação',autoridade:'+ R$ 2.500 de implantação'};
  const planoTag= {essencial:'Plano 01',presenca:'Plano 02 · Recomendado',autoridade:'Plano 03'};
  const planoSub= {essencial:'Identidade + presença funcional<br>Mínimo 6 meses',presenca:'O sistema Zyllah inteiro<br>Mínimo 6 meses',autoridade:'Para dominar o digital<br>Mínimo 12 meses'};
  const planoItens = {
    essencial:['Identidade visual definida — como você aparece em qualquer lugar','Instagram ativo toda semana, sem você precisar pensar no que postar','Fotos profissionais do seu consultório a cada trimestre','Relatório mensal do que funcionou e o que melhorou','Canal direto com a equipe em dias úteis'],
    presenca: ['Tudo do Essencial','Site que representa quem você é — pronto para receber paciente novo','Vídeos curtos de autoridade, editados e publicados','Instagram e LinkedIn gerenciados simultaneamente','Mensagens fora do horário respondidas automaticamente','Fotos profissionais todo mês','Uma conversa por mês para ajustar o rumo juntos'],
    autoridade:['Tudo do Presença Completa','Anúncios que alcançam pacientes que ainda não sabem que você existe','Vídeos semanais que mantêm sua autoridade em evidência constante','WhatsApp que qualifica e encaminha o paciente sozinho, 24 horas','LinkedIn posicionado para colegas, hospitais e imprensa','Reuniões quinzenais e visão de resultados a qualquer momento','Atendimento prioritário — você nunca espera']
  };
  const planoClasse = plano==='autoridade' ? 's-plano-ink s-plano' : 's-plano';

  const N = includePiloto ? 7 : 6;
  // Contadores fixos — N slides total, CTA sempre é o último
  const cS1=`1 · ${N}`,cS2=`2 · ${N}`,cS2b=`3 · ${N}`,cS3=`4 · ${N}`,cAsp=`5 · ${N}`,cPln=`6 · ${N}`,cPil=`7 · ${N}`,cCta=`${N} · ${N}`;

  const itensHTML = (planoItens[plano]||planoItens.essencial).map(i=>`<div class="item">${i}</div>`).join('');

  // Logo SVG inline (símbolo — reutilizado em todos os slides)
  const logoSym = `<svg style="display:none"><symbol id="logo-svg" viewBox="0 0 1451.04 1080"><defs><style>.cl1{fill:#c3a366}.cl2{fill:#c1a260}.cl3{fill:#c7a866}.cl4{fill:#a58346}.cl5{fill:#a38342}</style></defs><path class="cl1" d="M109.72,379.36l-2.14-89.96,301.48-.47-116.04,167.8h-37.51l92.37-135.5-.26-2.05-209.52.94s-5.44,17.21-7.92,28.54c-2.66,12.16-6.81,30.6-6.81,30.6l-13.66.1Z"/><polygon class="cl2" points="462.05 288.05 358.71 458.18 334.04 458.32 462.05 288.05"/><path class="cl3" d="M136.37,525.99c5.29-7.07,16.14-18.73,28.68-29.87,21.13-18.76,43.55-27.43,65.7-29.71,38.59-3.98,78.24,13.89,119.03,14.95,40.42,1.05,61.94-13.12,91.96-35.44-17.34,21.74-46.65,46.78-75.93,51.31-7.21,1.23-14.06,1.22-21.3.58-25.23-2.19-60.93-10.88-88.75-13.25-27.82-1.18-54.75,4.57-80.47,17.31-11.84,5.86-32,18.17-38.87,24.1h-.05Z"/><path class="cl4" d="M462.54,616.14l3.52,92.6-319.47,2.26,135.21-196.13h35.34l-112.44,164.22.97,2.99,226.87-1.7s5.79-15.43,9.85-32.83c3.28-14.06,5.39-31.17,5.39-31.17l14.77-.23Z"/><polygon class="cl5" points="91.97 707.77 214.37 513.64 241.88 513.49 91.97 707.77"/><path class="cl3" d="M643.13,506.97l-5.71,33.88h-114.76l90.4-131.73h-59.08s-10.21-.38-17.44,7.61-8.38,15.61-8.38,15.61l4.19-34.26s1.32,1.74,5.33,2.28c.42.06,2.56.33,5.93.33,22.19,0,95.34.81,95.34.81l-90.61,132.49,65.73-.76s10.41.38,16.88-6.47c6.47-6.85,12.18-19.8,12.18-19.8Z"/><path class="cl3" d="M649.09,401.77h50.76s-9.9,3.55-11.17,8.12c-1.27,4.57,37.82,64.21,37.82,64.21,0,0,36.55-54.31,37.06-63.45.51-9.14-13.2-8.88-13.2-8.88h42.39s-7.61,2.03-11.93,6.35c-4.31,4.31-48.22,73.6-48.22,73.6l-.51,48.73s-.44,4.01,4.89,6.57c2.35,1.13,9.75,1.63,9.75,1.63v1.77l-46.95.18.13-1.77s5.9-.31,8.95-2.02,4.19-4.57,4.19-4.57l-.38-47.78s0-.95-.95-2.47-44.54-68.15-44.54-68.15c0,0-1.71-3.81-6.47-6.28s-11.6-4.19-11.6-4.19v-1.6Z"/><path class="cl3" d="M804.95,401.73h49.49v1.62s-6.09-.67-10.18,2.19c-4.09,2.86-4.47,5.33-4.47,5.33l-.1,117.83s-.07,1.79,1.33,3.33c1.73,1.9,3.52,1.9,3.52,1.9l49.87.19s6.47.38,13.71-8.95c7.23-9.33,9.14-17.8,9.14-17.8h1.33l-6.76,33.41h-103.93s5.67-.16,10.47-5.14c3.39-3.52,2.95-7.61,2.95-7.61l-.19-114.12s.1-4-5.33-6.95c-5.43-2.95-10.85-3.9-10.85-3.9v-1.33Z"/><path class="cl3" d="M932.51,401.73h49.49v1.62s-6.09-.67-10.18,2.19c-4.09,2.86-4.47,5.33-4.47,5.33l-.1,117.83s-.07,1.79,1.33,3.33c1.73,1.9,3.52,1.9,3.52,1.9l49.87.19s6.47.38,13.71-8.95c7.23-9.33,9.14-17.8,9.14-17.8h1.33l-6.76,33.41h-103.93s5.67-.16,10.47-5.14c3.39-3.52,2.95-7.61,2.95-7.61l-.19-114.12s.1-4-5.33-6.95c-5.43-2.95-10.85-3.9-10.85-3.9v-1.33Z"/><path class="cl3" d="M1198.18,537.91c-3.05-.76-4.82-2.03-8.12-6.85-3.3-4.82-63.07-131.73-63.07-131.73,0,0-54.31,126.98-57.22,131.31-2.9,4.33-6.38,6.52-9.23,7.76-2.86,1.24-6.77,1-6.77,1v1.38l40.36-.21.25-1.65s-6.6-.25-10.41-2.54c-3.81-2.28-3.81-1.78-3.81-6.85s15.48-40.1,15.48-40.1l54.06.25s18.15,37.06,18.15,40.36-.25,6.09-3.05,7.49c-2.79,1.4-8.63,1.02-8.63,1.02v2.03h48.6v-2.03s-3.55.13-6.6-.63ZM1098.51,484.14l24.03-59.28,25.89,59.14-49.92.15Z"/><path class="cl3" d="M1215.69,402.04h46.92v1.52s-6.19-.38-9.04,1.9-4.76,5.14-4.76,7.8-.1,52.35-.1,52.35h78.05s0-54.25,0-54.25c0,0,.76-3.33-4.28-6-5.04-2.66-11.61-2.09-11.61-2.09l.1-1.24h49.02v1.43s-9.52-.48-11.99,1.81-2.76,4-2.76,5.33,0,118.02,0,119.83.1,3.81,5.14,6.47c5.04,2.66,9.61,2.09,9.61,2.09v1.33h-47.68l.1-1.33s7.42.19,9.71-.67c2.28-.86,3.05-1.71,3.71-3.24s1.24-1.81,1.24-6.76.1-56.35.1-56.35l-79-.1v58.63c0,2.09.48,4.09,4.95,6.09s9.61,1.9,9.61,1.9v1.81h-47.49v-1.71s7.71,1.14,11.04-1.9c3.33-3.05,3.9-4.47,3.9-6.28s-.1-116.97-.1-118.69-1.33-4.09-5.14-6-9.71-2.66-9.71-2.66l.48-1.05Z"/></symbol></svg>`;

  const CORNERS = `<div class="corner tl"></div><div class="corner tr"></div><div class="corner bl"></div><div class="corner br"></div>`;
  const LOGO = `<div class="logo-rodape"><svg style="width:150px;height:auto"><use href="#logo-svg"/></svg></div>`;

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Zyllah — ${p.nome} · ${nomePlano[plano]}</title>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;1,300&family=Jost:wght@300;400&display=swap" rel="stylesheet">
<style>
  :root{--ink:#1A1714;--gold:#B8976A;--gold-light:#D4B896;--cream:#FAF7F2;--cream-mid:#F0EBE1;--brown:#4A3F35;}
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
  body{background:#111;font-family:'Jost',sans-serif;font-weight:300;display:flex;flex-direction:column;align-items:center;gap:2px;padding:40px 20px 80px;}
  h2.label{color:#555;font-family:'Jost',sans-serif;font-weight:300;font-size:11px;letter-spacing:3px;text-transform:uppercase;margin:32px 0 8px;align-self:flex-start;margin-left:calc(50% - 960px);}
  .slide{width:1920px;height:1080px;position:relative;overflow:hidden;flex-shrink:0;}
  .corner{position:absolute;width:28px;height:28px;border-color:var(--gold);border-style:solid;opacity:.5;z-index:10;}
  .corner.tl{top:56px;left:80px;border-width:1px 0 0 1px;}
  .corner.tr{top:56px;right:80px;border-width:1px 1px 0 0;}
  .corner.bl{bottom:56px;left:80px;border-width:0 0 1px 1px;}
  .corner.br{bottom:56px;right:80px;border-width:0 1px 1px 0;}
  .logo-rodape{position:absolute;bottom:80px;right:110px;z-index:10;}
  .counter{position:absolute;bottom:84px;left:110px;font-family:'Jost',sans-serif;font-weight:300;font-size:13px;letter-spacing:2px;opacity:.4;z-index:10;}
  .grade{position:absolute;inset:0;background-image:linear-gradient(rgba(184,151,106,.04)1px,transparent 1px),linear-gradient(90deg,rgba(184,151,106,.04)1px,transparent 1px);background-size:54px 54px;}
  /* S1 — Capa */
  .s1{background:var(--ink);}
  .s1 .counter{color:var(--gold);}
  .s1 .content{position:absolute;bottom:140px;left:160px;right:800px;}
  .s1 .pretag{font-size:11px;letter-spacing:.5em;text-transform:uppercase;color:var(--gold);margin-bottom:28px;display:flex;align-items:center;gap:16px;}
  .s1 .pretag::before{content:'';display:block;width:32px;height:1px;background:var(--gold);}
  .s1 h1{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:96px;line-height:1.0;color:white;margin-bottom:36px;}
  .s1 .sub{font-size:22px;color:rgba(255,255,255,.45);line-height:1.7;max-width:680px;}
  /* S2 — Problema */
  .s2{background:var(--cream);}
  .s2 .counter{color:var(--ink);opacity:.3;}
  .s2 .corner{border-color:var(--ink);opacity:.15;}
  .s2 .content{position:absolute;top:50%;left:160px;right:160px;transform:translateY(-50%);}
  .s2 .tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:20px;display:block;}
  .s2 h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;line-height:1.05;color:var(--ink);margin-bottom:36px;}
  .s2 .linha{width:56px;height:1px;background:var(--gold);margin-bottom:32px;}
  .s2 .corpo{font-size:26px;color:var(--brown);line-height:1.8;max-width:960px;}
  .num-fantasma{position:absolute;font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:700px;line-height:1;color:var(--gold);opacity:.05;right:80px;top:50%;transform:translateY(-50%);z-index:0;user-select:none;}
  /* S2b — Diagnóstico */
  .s2b{background:var(--ink);}
  .s2b .counter{color:rgba(184,151,106,.4);}
  .s2b .content{position:absolute;top:50%;left:160px;right:160px;transform:translateY(-50%);}
  .s2b .tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:28px;display:block;}
  .s2b h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:68px;line-height:1.1;color:white;margin-bottom:60px;}
  .s2b .grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:60px;}
  .s2b .col-label{font-size:11px;letter-spacing:.35em;text-transform:uppercase;color:var(--gold);margin-bottom:16px;display:block;}
  .s2b .col-body{font-size:20px;color:rgba(255,255,255,.55);line-height:1.65;}
  /* S3 — Solução */
  .s3{background:var(--ink);}
  .s3 .counter{color:var(--gold);}
  .s3 .content{position:absolute;top:50%;left:160px;right:160px;transform:translateY(-50%);}
  .s3 .tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:20px;display:block;}
  .s3 h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;line-height:1.05;color:white;margin-bottom:36px;}
  .s3 .linha{width:56px;height:1px;background:var(--gold);margin-bottom:32px;opacity:.6;}
  .s3 .corpo{font-size:26px;color:rgba(255,255,255,.55);line-height:1.8;max-width:1100px;}
  .s3 .corpo strong{color:var(--gold-light);font-weight:400;}
  /* S-Aspir */
  .s-aspir{position:relative;overflow:hidden;background:#1a1714;}
  .s-aspir .bg-img{position:absolute;inset:0;background:linear-gradient(135deg,#2a2420 0%,#1a1714 100%);}
  .s-aspir .bg-overlay{position:absolute;inset:0;background:linear-gradient(135deg,rgba(26,23,20,.85)0%,rgba(26,23,20,.4)100%);z-index:1;}
  .aspir-content{position:absolute;inset:0;z-index:2;display:flex;flex-direction:column;align-items:center;justify-content:center;}
  .aspir-frase{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;line-height:1.2;color:white;text-align:center;}
  .aspir-frase em{color:var(--gold);font-style:italic;}
  .aspir-linha-topo{width:1px;height:80px;background:var(--gold);opacity:.4;margin:0 auto 48px;}
  .aspir-linha-base{width:1px;height:80px;background:var(--gold);opacity:.4;margin:48px auto 0;}
  /* S-Plano */
  .s-plano{background:var(--cream);}
  .s-plano .counter{color:var(--ink);opacity:.3;}
  .s-plano .corner{border-color:var(--ink);opacity:.15;}
  .s-plano .content{position:absolute;top:110px;left:160px;right:160px;bottom:110px;display:flex;gap:80px;align-items:center;}
  .s-plano .col-esq{flex:0 0 420px;display:flex;flex-direction:column;justify-content:center;height:100%;}
  .s-plano .col-dir{flex:1;display:flex;flex-direction:column;justify-content:center;}
  .s-plano .plano-tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:12px;display:block;}
  .s-plano .plano-nome{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:72px;line-height:1.0;color:var(--ink);margin-bottom:8px;}
  .s-plano .plano-sub{font-size:15px;color:rgba(74,63,53,.6);margin-bottom:32px;letter-spacing:.05em;line-height:1.5;}
  .s-plano .linha{width:56px;height:1px;background:var(--gold);margin-bottom:28px;}
  .s-plano .preco-row{display:flex;align-items:baseline;gap:12px;margin-bottom:8px;}
  .s-plano .preco{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;color:var(--gold);line-height:1;}
  .s-plano .preco-label{font-size:16px;color:rgba(74,63,53,.55);}
  .s-plano .implantacao{font-size:13px;color:rgba(74,63,53,.5);letter-spacing:.05em;}
  .s-plano .itens{display:flex;flex-direction:column;gap:18px;}
  .s-plano .item{display:flex;align-items:flex-start;gap:20px;font-size:30px;color:var(--brown);line-height:1.45;}
  .s-plano .item::before{content:'✦';color:var(--gold);opacity:.7;font-size:14px;margin-top:9px;flex-shrink:0;}
  .s-plano-ink{background:var(--ink);}
  .s-plano-ink .counter{color:var(--gold);}
  .s-plano-ink .corner{border-color:var(--gold);opacity:.4;}
  .s-plano-ink .plano-nome{color:white;}
  .s-plano-ink .plano-sub{color:rgba(255,255,255,.4);}
  .s-plano-ink .item{color:rgba(255,255,255,.65);}
  .s-plano-ink .preco-label{color:rgba(255,255,255,.4);}
  .s-plano-ink .implantacao{color:rgba(255,255,255,.35);}
  /* S-Piloto */
  .s-piloto{background:var(--ink);}
  .s-piloto .counter{color:var(--gold);}
  .s-piloto .content{position:absolute;top:50%;left:160px;right:160px;transform:translateY(-50%);text-align:center;}
  .s-piloto .tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:24px;display:block;}
  .s-piloto h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;line-height:1.05;color:white;margin-bottom:40px;}
  .s-piloto .linha{width:56px;height:1px;background:var(--gold);margin:0 auto 40px;opacity:.5;}
  .s-piloto .condicoes{font-size:22px;color:rgba(255,255,255,.45);line-height:1.8;}
  .s-piloto .condicoes strong{color:var(--gold-light);font-weight:400;}
  /* S-CTA */
  .s-cta{background:var(--cream);}
  .s-cta .counter{color:var(--ink);opacity:.3;}
  .s-cta .corner{border-color:var(--ink);opacity:.15;}
  .s-cta .content{position:absolute;top:50%;left:160px;right:160px;transform:translateY(-50%);}
  .s-cta .tag{font-size:11px;letter-spacing:.45em;text-transform:uppercase;color:var(--gold);margin-bottom:24px;display:block;}
  .s-cta h2{font-family:'Cormorant Garamond',serif;font-style:italic;font-weight:300;font-size:80px;line-height:1.05;color:var(--ink);margin-bottom:40px;}
  .s-cta .linha{width:56px;height:1px;background:var(--gold);margin:0 auto 40px;}
  .s-cta .proximos{font-size:22px;color:var(--brown);line-height:2.0;}
  .s-cta .proximos strong{color:var(--ink);font-weight:400;}
  @media print{body{background:none!important;padding:0!important;gap:0!important;}h2.label{display:none!important;}.slide{page-break-after:always;margin:0!important;}}
</style>
</head>
<body>
${logoSym}

<!-- SLIDE 1 — CAPA -->
<h2 class="label">01 · Capa</h2>
<div class="slide s1">
  <div class="grade"></div>
  ${CORNERS}
  <div class="counter">${cS1}</div>
  ${LOGO}
  <div class="content">
    <div class="pretag">${p.esp||'Especialista de saúde'}</div>
    <h1>Para <em style="color:var(--gold)">${primeiro}</em></h1>
    <div class="sub">Diagnóstico preparado especialmente para você<br>pela equipe Zyllah Digital · ${cidade}</div>
  </div>
</div>

<!-- SLIDE 2 — PROBLEMA -->
<h2 class="label">02 · O problema</h2>
<div class="slide s2">
  ${CORNERS}
  <div class="counter">${cS2}</div>
  ${LOGO}
  <div class="num-fantasma">?</div>
  <div class="content">
    <span class="tag">${t.slide02?.tag||'A realidade'}</span>
    <h2>${t.slide02?.titulo||'O desafio da presença digital'}</h2>
    <div class="linha"></div>
    <div class="corpo">${t.slide02?.corpo||''}</div>
  </div>
</div>

<!-- SLIDE 2b — DIAGNÓSTICO PERSONALIZADO -->
<h2 class="label">02b · Diagnóstico personalizado</h2>
<div class="slide s2b">
  <div class="grade"></div>
  ${CORNERS}
  <div class="counter">${cS2b}</div>
  ${LOGO}
  <div class="content">
    <span class="tag">Diagnóstico · ${ig||p.nome} · ${cidade}</span>
    <h2>${t.slide02b?.titulo_principal||primeiro+', encontramos o que está limitando seu crescimento digital.'}</h2>
    <div class="grid">
      <div>
        <span class="col-label">Cenário atual</span>
        <div class="col-body">${t.slide02b?.cenario_atual||'—'}</div>
      </div>
      <div>
        <span class="col-label">O que está faltando</span>
        <div class="col-body">${t.slide02b?.lacunas||'—'}</div>
      </div>
      <div>
        <span class="col-label">Potencial identificado</span>
        <div class="col-body">${t.slide02b?.potencial||'—'}</div>
      </div>
    </div>
  </div>
</div>

<!-- SLIDE 3 — SOLUÇÃO -->
<h2 class="label">03 · A solução</h2>
<div class="slide s3">
  <div class="grade"></div>
  ${CORNERS}
  <div class="counter">${cS3}</div>
  ${LOGO}
  <div class="content">
    <span class="tag">${t.slide03?.tag||'Sistema Zyllah'}</span>
    <h2>${t.slide03?.titulo||'Comunicação digital que opera por você'}</h2>
    <div class="linha"></div>
    <div class="corpo">${t.slide03?.corpo||''}</div>
  </div>
</div>

<!-- SLIDE 4 — ASPIRAÇÃO -->
<h2 class="label">04 · Aspiração</h2>
<div class="slide s-aspir">
  ${CORNERS}
  <div class="counter" style="color:rgba(184,151,106,.35);z-index:3">${cAsp}</div>
  <div class="bg-img"></div>
  <div class="bg-overlay"></div>
  <div class="aspir-content">
    <div class="aspir-linha-topo"></div>
    <div class="aspir-frase">${t.aspiracao||'Sua presença digital funcionando perfeitamente<br><em>enquanto você se dedica ao que realmente importa</em>'}</div>
    <div class="aspir-linha-base"></div>
  </div>
</div>

<!-- SLIDE 5 — PLANO -->
<h2 class="label">05 · Plano ${nomePlano[plano]}</h2>
<div class="slide ${planoClasse}">
  ${plano==='autoridade'?'<div class="grade"></div>':''}
  ${CORNERS}
  <div class="counter">${cPln}</div>
  ${LOGO}
  <div class="num-fantasma" style="font-size:900px;opacity:0.03">${nomePlano[plano][0]}</div>
  <div class="content">
    <div class="col-esq">
      <span class="plano-tag">${planoTag[plano]}</span>
      <div class="plano-nome">${nomePlano[plano]}</div>
      <div class="plano-sub">${planoSub[plano]}</div>
      <div class="linha"></div>
      <div class="preco-row">
        <div class="preco">${precos[plano]}</div>
        <div class="preco-label">/mês</div>
      </div>
      <div class="implantacao">${impl[plano]} na contratação</div>
    </div>
    <div class="col-dir">
      <div class="itens">${itensHTML}</div>
    </div>
  </div>
</div>

${includePiloto ? `
<!-- SLIDE 6 — PILOTO -->
<h2 class="label">06 · Piloto</h2>
<div class="slide s-piloto">
  <div class="grade"></div>
  ${CORNERS}
  <div class="counter">${cPil}</div>
  ${LOGO}
  <div class="content">
    <span class="tag">Sem compromisso</span>
    <h2>Duas semanas para<br>você ver na prática</h2>
    <div class="linha"></div>
    <div class="condicoes"><strong>100% gratuito</strong> — configuramos e operamos por 2 semanas.<br>Sem contrato, sem cartão, sem compromisso.<br>Se não fizer sentido para você, encerramos sem custo e sem pressão.<br>Se fizer, conversamos sobre continuidade quando você estiver pronto.</div>
  </div>
</div>
` : ''}

<!-- SLIDE CTA -->
<h2 class="label">${includePiloto ? '07' : '06'} · Próximos passos</h2>
<div class="slide s-cta">
  ${CORNERS}
  <div class="counter">${cCta}</div>
  ${LOGO}
  <div class="content">
    <span class="tag">Como avançamos</span>
    <h2>Simples e direto</h2>
    <div class="linha"></div>
    <div class="proximos">${t.cta?.proximos||'<strong>01 · Conversa</strong> — 30 minutos para entender sua rotina e confirmar o diagnóstico<br><strong>02 · Proposta</strong> — Contrato simples, sem pegadinhas<br><strong>03 · Setup</strong> — Configuração da estrutura de comunicação<br><strong>04 · Operação</strong> — Sistema funcionando, você focada no que realmente importa'}</div>
  </div>
</div>

</body>
</html>`;
}

// ============================================================
// MAPEAR +10 — Sonnet pesquisa novos especialistas em NF
// ============================================================

function mapearMais10(d) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return {erro:'ANTHROPIC_API_KEY não configurada'};

  const cidade    = d.cidade || 'Nova Friburgo RJ';
  const espExist  = d.especialidades_existentes || '';

  const prompt = `Você é um pesquisador de prospecção para Zyllah Digital, agência de presença digital para especialistas de saúde em ${cidade}.

Especialidades JÁ MAPEADAS (não repetir): ${espExist}

Sua tarefa: identifique 10 novos especialistas médicos ou odontológicos que provavelmente atuam em ${cidade} e que ainda NÃO foram mapeados. Use seu conhecimento sobre o perfil típico de profissionais de saúde em cidades médias do interior do RJ.

Para cada especialista, preencha os campos com base no que é PROVÁVEL para o perfil (não invente dados específicos como CRM ou Instagram real — use apenas estimativas realistas):

Responda APENAS JSON válido:
{
  "novos": [
    {
      "nome": "Dr(a). [Nome Plausível]",
      "esp": "especialidade",
      "crm": "N/D",
      "instagram": "",
      "seguidores_ig": "",
      "posts_ig": "",
      "nota_presenca": "2",
      "tem_site": "nao",
      "google_meu_negocio": "nao",
      "doctoralia": "nao",
      "linkedin": "nao",
      "gap_principal": "descrição do gap típico desta especialidade",
      "gancho": "ângulo de abordagem sugerido",
      "risco": "baixo",
      "notas": "Gerado por Sonnet — verificar manualmente antes de abordar",
      "proximo_passo": "Verificar presença digital manualmente"
    }
  ],
  "msg": "resumo da pesquisa"
}`;

  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:'post', contentType:'application/json',
      headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({
        model:'claude-sonnet-4-5',
        max_tokens:3000,
        messages:[{role:'user',content:prompt}]
      }),
      muteHttpExceptions:true
    });
    if (resp.getResponseCode()!==200) return {erro:'API '+resp.getResponseCode()};
    const txt = JSON.parse(resp.getContentText()).content[0].text.trim().replace(/```json\n?|```/g,'').trim();
    return JSON.parse(txt);
  } catch(err) { return {erro:err.message}; }
}

// ── YOUTUBE — BUSCA VIA APPS SCRIPT ──────────────────────────

function buscarYoutube(d) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('YT_API_KEY');
  if (!apiKey) return { erro: 'YT_API_KEY não configurada no Apps Script' };

  const termos = d.termos || ['marketing médico','autoridade médica'];
  const maxPorTermo = 3;
  const seteDias = new Date(Date.now() - 7*24*60*60*1000).toISOString();

  const videos = [];
  const vistos = new Set();

  termos.forEach(termo => {
    try {
      const url = `https://www.googleapis.com/youtube/v3/search?part=snippet&q=${encodeURIComponent(termo)}&type=video&publishedAfter=${seteDias}&maxResults=${maxPorTermo}&relevanceLanguage=pt&key=${apiKey}`;
      const resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      const data = JSON.parse(resp.getContentText());
      if (data.error) { Logger.log('YT error: '+JSON.stringify(data.error)); return; }
      (data.items||[]).forEach(item => {
        const id = item.id?.videoId;
        if (!id || vistos.has(id)) return;
        vistos.add(id);
        videos.push({
          id, termo,
          titulo: item.snippet?.title,
          canal: item.snippet?.channelTitle,
          data: item.snippet?.publishedAt,
          thumb: item.snippet?.thumbnails?.medium?.url
        });
      });
    } catch(e) { Logger.log('Erro termo "'+termo+'": '+e.message); }
  });

  return { videos };
}

// ── PAUTA IA (Haiku) ──────────────────────────────────────────

function gerarPauta(d) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { erro: 'Chave não configurada' };

  const diagnostico = d.diagnostico || '';
  const clientes = d.clientes || '[]';
  const prospects = d.prospects || '[]';
  const alvo = d.alvo || 'zyllah';

  let prompt;
  if (alvo === 'cliente') {
    const cn = d.cliente_nome || '(cliente)';
    const ce = d.cliente_esp || '(especialidade não informada)';
    prompt = `Você é estrategista de conteúdo da Zyllah Digital para o cliente abaixo.

Cliente: ${cn}
Especialidade: ${ce}
Diagnóstico da semana anterior: ${diagnostico || '(não informado)'}

Regras CFM obrigatórias:
- PROIBIDO: antes/depois, promessa de resultado, preço de procedimento, depoimento identificado, sensacionalismo
- PERMITIDO: conteúdo educativo, informativo, humanização do especialista, rotina da clínica, mitos e verdades

Gere exatamente 5 sugestões de pauta para ESTE cliente publicar nas redes dele (Instagram/LinkedIn/Reels/Stories) na próxima semana — adequadas à especialidade e ao CFM.

Responda APENAS com JSON válido, sem markdown, sem explicações fora do JSON:
{
  "sugestoes": [
    {
      "tema": "Título do conteúdo (máx 60 chars)",
      "plataformas": ["instagram"],
      "formato": "carrossel | reels | post | stories",
      "justificativa": "Por que este tema agora (1 frase)"
    }
  ]
}`;
  } else {
    prompt = `Você é estrategista de conteúdo da ZYLLAH DIGITAL — agência de presença digital para especialistas de saúde em Nova Friburgo, RJ.

A Zyllah é uma MARCA em construção, solo, pré-receita. A Zyllah NÃO é médica — é agência. O público das postagens da Zyllah são MÉDICOS ESPECIALISTAS que ainda não contrataram a Zyllah (prospects).

Gere 5 sugestões de pauta para a CONTA DA PRÓPRIA ZYLLAH publicar (Instagram e LinkedIn), com o objetivo de:
- Demonstrar autoridade em presença digital estratégica para saúde
- Atrair médicos como prospects
- Diferenciar a Zyllah de agências genéricas (que entregam o mesmo feed pra todo mundo)
- Fazer leitura crítica e ética do mercado (CFM, promessas milagrosas, feeds clichê de clínica)
- Trazer bastidor da operação, cases, dados, estudos

NÃO aplicar regras CFM aqui — a Zyllah não é profissional de saúde. O tom é sóbrio, crítico, preciso. Evitar clichês de agência ("boost", "transformar", "6 dígitos", "seu negócio nas alturas").

Contexto atual:
- Clientes ativos da Zyllah: ${clientes}
- Prospects em aquecimento: ${prospects}
- Diagnóstico da semana anterior: ${diagnostico || '(não informado)'}

Responda APENAS com JSON válido, sem markdown, sem explicações fora do JSON:
{
  "sugestoes": [
    {
      "tema": "Título do conteúdo (máx 60 chars)",
      "plataformas": ["instagram","linkedin"],
      "formato": "carrossel | reels | post | stories",
      "justificativa": "Por que este tema agora (1 frase)"
    }
  ]
}`;
  }

  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: MODELO_HAIKU,
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) return { erro: 'API Anthropic retornou ' + code + ': ' + resp.getContentText().substring(0,120), sugestoes: [] };
    const raw = JSON.parse(resp.getContentText());
    const txt = raw.content?.[0]?.text || '{}';
    const parsed = JSON.parse(txt.replace(/```json|```/g,'').trim());
    return parsed;
  } catch(e) {
    return { erro: e.message, sugestoes: [] };
  }
}

// ── YOUTUBE — GERAÇÃO DE COMENTÁRIO (Haiku) ──────────────────

function gerarComentarioYT(d) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { erro: 'Chave não configurada' };
  const prompt = `Você é Guilherme Caetano, fundador da Zyllah Digital — agência de presença digital para especialistas de saúde em Nova Friburgo, RJ.\nEscreva um comentário para o vídeo abaixo. Seja genuíno, agregue valor, SEM vender. 2-3 frases. Termine com pergunta aberta. Português pt-BR.\n\nVídeo: "${d.titulo}"\nCanal: ${d.canal}\nTermo: ${d.termo}\n\nResponda SOMENTE com o texto do comentário.`;
  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: MODELO_HAIKU, max_tokens: 300, messages: [{ role: 'user', content: prompt }] }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return { erro: 'API erro ' + resp.getResponseCode() };
    return { comentario: JSON.parse(resp.getContentText()).content[0].text.trim() };
  } catch(err) { return { erro: err.message }; }
}

instalarTriggerOnEdit()

// ============================================================
// TESTES
// ============================================================

function testarEmail(){enviarEmail('[TESTE] Zyllah v3','Script v3 ok. CORS resolvido via JSONP puro.');}
function testarLegenda(){gerarLegenda(2);}
function testarVerificacao(){verificacaoDiaria();}
function testarResumo(){resumoSemanal();}
function testarExportacao(){exportarContexto();Logger.log('Verifique a pasta Zyllah_Contexto no Drive.');}
function testarApresentacao(){gerarApresentacoesNovosProspects(SpreadsheetApp.openById(PLANILHA_ID));}