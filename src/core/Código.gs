/**
 * SINALIZAÇÃO DO MALL
 * Runtime atual: MVP-3.32.0-SINALIZACAO-S26.10 | Fase S26.10
 *
 * GOVERNANÇA DE NOMENCLATURA — NÃO RENOMEAR APIs LEGADAS SEM MIGRAÇÃO:
 * - S5A/S5B em nomes de funções identificam a trilha histórica do offline e permanecem por compatibilidade.
 * - S23.6 identifica a origem histórica da exclusão operacional/soft delete.
 * - Referências posteriores a S25.6 em histórico/relatórios documentam a mesma regra de preservação do ativo;
 *   não devem ser confundidas com a antiga UI experimental de gestão de versões cartográficas, removida na S26.6.2-A.
 * - S25.5 continua sendo a publicação cartográfica controlada RASCUNHO -> VALIDAÇÃO -> PUBLICADA.
 * - S231 é o identificador legado e único do serviço de Google Slides; preservar uma única declaração global.
 */
const APP = Object.freeze({
  ID: 'SINALIZACAO_MALL', NOME: 'CEOP — Central de Operações',
  VERSAO: 'MVP-3.32.0-SINALIZACAO-S26.10', FASE: 'S26.10',
  TIMEZONE: 'America/Fortaleza', MODO_DADOS: 'LOCAL_INDEPENDENTE'
});

function doGet(e) {
  if (e && e.parameter && e.parameter.admin === 'sync_boxes_roxo') {
    try {
      const res = typeof inserirBoxesIlhaCentralSetorRoxo === 'function' ? inserirBoxesIlhaCentralSetorRoxo() : { ok: false, error: 'Função não encontrada' };
      return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err?.message || err) }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if (e && e.parameter && e.parameter.admin === 'sanear_toponimia') {
    try {
      const resRoxo = typeof inserirBoxesIlhaCentralSetorRoxo === 'function' ? inserirBoxesIlhaCentralSetorRoxo() : { inseridos: 0 };
      const res = typeof sanearToponimiaLojasMapa === 'function' ? sanearToponimiaLojasMapa() : { ok: false };
      const resLojistas = typeof executarSincronizacaoRuasPlanilhaLojistas_ === 'function' ? executarSincronizacaoRuasPlanilhaLojistas_() : { ok: false };
      return ContentService.createTextOutput(JSON.stringify({ ok: true, sanear: res, roxo: resRoxo, lojistas: resLojistas }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err?.message || err) }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if (e && e.parameter && e.parameter.admin === 'sanear_ruas_lojistas') {
    try {
      const resLojistas = typeof executarSincronizacaoRuasPlanilhaLojistas_ === 'function' ? executarSincronizacaoRuasPlanilhaLojistas_() : { ok: false };
      return ContentService.createTextOutput(JSON.stringify({ ok: true, lojistas: resLojistas }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err?.message || err) }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if (e && e.parameter && e.parameter.debug === 'raw') {
    const raw = HtmlService.createTemplateFromFile('index').evaluate().getContent();
    return ContentService.createTextOutput(raw).setMimeType(ContentService.MimeType.TEXT);
  }
  if (e && e.parameter && e.parameter.debug === 'line979') {
    const raw = HtmlService.createTemplateFromFile('index').evaluate().getContent();
    const lines = raw.split('\n');
    const result = {
      totalLines: lines.length,
      line979: lines[978] || '',
      around: lines.slice(Math.max(0, 965), Math.min(lines.length, 990))
    };
    return ContentService.createTextOutput(JSON.stringify(result, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }
  // Auditoria S26.10 — remover ALLOWALL: impede embedding/clickjacking.
  // Comportamento padrão do Apps Script (X-Frame-Options) passa a valer.
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(APP.NOME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover, maximum-scale=1');
}
function include(nome) { return HtmlService.createHtmlOutputFromFile(nome).getContent(); }

function onOpen() {
  // Auditoria S26.10 — menu atualizado: itens S26.7 encerrados removidos;
  // 'setupAtual' (setupS267) não representa mais a fase corrente.
  SpreadsheetApp.getUi().createMenu('CEOP — Operações')
    .addItem('Diagnóstico da release S26.10', 'diagnosticoS2610F2')
    .addSeparator()
    .addItem('Garantir colunas de Autorização de Serviço (AS)', 'menuGarantirColunasCeopAs')
    .addItem('Inicializar estruturas de Patrimônio e Reclamações', 'menuInicializarPatrimonioEReclamacoes')
    .addItem('Sanear toponímia de boxes e galerias', 'menuSanearToponimiaLojasMapa')
    .addItem('Sanear caracteres corrompidos (acentuação / )', 'menuSanearCaracteresLojistas')
    .addItem('Inserir boxes da Ilha Central do Setor Roxo (16 boxes)', 'menuInserirBoxesIlhaCentralSetorRoxo')
    .addItem('Provisionar novo setor (Clonar base limpa)', 'menuProvisionarNovoSetor')
    .addItem('Abrir pasta raiz no Drive', 'abrirPastaRaizS0')
    .addToUi();
}

/**
 * Menu interativo para inicializar as abas e estruturas de Patrimônio e Reclamações.
 */
function menuInicializarPatrimonioEReclamacoes() {
  const ui = SpreadsheetApp.getUi();
  try {
    garantirAbaPatrimonioOrdensServico_();
    garantirAbasReclamacoes_();
    ui.alert(
      'Estruturas Inicializadas',
      'As abas PATRIMONIO_ORDENS_SERVICO, CHAMADOS_RECLAMACOES e CHAMADOS_ANEXOS foram verificadas e inicializadas com sucesso na planilha.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Erro na Inicialização', 'Ocorreu um erro: ' + String(e?.message || e), ui.ButtonSet.OK);
  }
}


/**
 * Menu interativo para executar o saneamento de caracteres corrompidos na Base Mestre de Lojistas.
 */
function menuSanearCaracteresLojistas() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Saneamento de Caracteres (Encoding / )',
    'Deseja analisar e corrigir os caracteres corrompidos () na Planilha Mestre de Lojistas?\n\n' +
    'Esta operação irá restaurar automaticamente as acentuações corrompidas na importação:\n' +
    '• Disponível (em mais de 1.400 registros);\n' +
    '• Categorias de Mix (Confecções, Vestuário, Calçados, Serviços, etc.);\n' +
    '• Subtipos de Lojas (Satélite, Alimentação, Serviços);\n' +
    '• Nomes próprios e fantasias com acentuação.\n\n' +
    'Deseja prosseguir?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    const res = typeof executarSaneamentoCaracteresLojistas_ === 'function'
      ? executarSaneamentoCaracteresLojistas_()
      : { celulasCorrigidas: 0, linhasComCorrecao: 0 };
    ui.alert(
      'Saneamento Concluído',
      'Processo finalizado com sucesso!\n\n' +
      '• Total de linhas verificadas: ' + (res.totalLinhas || 0) + '\n' +
      '• Células corrigidas: ' + res.celulasCorrigidas + '\n' +
      '• Linhas restauradas: ' + res.linhasComCorrecao + '\n\n' +
      (res.celulasCorrigidas > 0
        ? 'Todos os caracteres foram restaurados com suas respectivas acentuações na aba LOJISTAS.'
        : 'Nenhum caractere corrompido foi encontrado na base.'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Aviso', 'Falha ao sanear caracteres: ' + (err?.message || err), ui.ButtonSet.OK);
  }
}

/**
 * Menu interativo para executar o saneamento toponímico de boxes e galerias na planilha.
 */
function menuSanearToponimiaLojasMapa() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Saneamento Toponímico de Boxes e Galerias',
    'Deseja analisar e sanear a toponímia da aba LOJAS_MAPA em todos os setores?\n\n' +
    'Esta operação irá:\n' +
    '• Reassociar boxes de ponta de galeria das avenidas perimetrais (AVALN, AVCRP, RJSA) para suas respectivas ruas verticais;\n' +
    '• Corrigir o eixo de Dom Manuel no Setor Roxo;\n' +
    '• Garantir a presença dos 16 boxes da Ilha Central do Setor Roxo;\n' +
    '• Preencher os códigos canônicos LUC diretamente na planilha.\n\n' +
    'Deseja prosseguir?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    const resRoxo = typeof inserirBoxesIlhaCentralSetorRoxo === 'function' ? inserirBoxesIlhaCentralSetorRoxo() : { inseridos: 0 };
    const res = sanearToponimiaLojasMapa();
    const totalInseridosRoxo = resRoxo.inseridos || 0;
    ui.alert(
      'Saneamento Concluído',
      'Processo finalizado com sucesso!\n\n' +
      '• Total de boxes analisados: ' + res.totalAnalisados + '\n' +
      '• Boxes corrigidos/padronizados: ' + res.totalCorrigidos + '\n' +
      (totalInseridosRoxo > 0 ? '• Boxes da Ilha Central do Roxo inseridos: ' + totalInseridosRoxo + '\n\n' : '\n') +
      (res.totalCorrigidos > 0 || totalInseridosRoxo > 0
        ? 'Os boxes foram atualizados com sucesso na aba LOJAS_MAPA.'
        : 'Todos os boxes já estavam devidamente sincronizados com a Base Mestre.'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Erro', 'Falha ao executar o saneamento: ' + (err?.message || err), ui.ButtonSet.OK);
  }
}

/**
 * Menu interativo para inserir os 16 boxes da Ilha Central do Setor Roxo.
 */
function menuInserirBoxesIlhaCentralSetorRoxo() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = inserirBoxesIlhaCentralSetorRoxo();
    ui.alert(
      'Ilha Central do Setor Roxo',
      res.mensagem || ('Boxes inseridos: ' + res.inseridos),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Erro', 'Falha ao inserir boxes da Ilha Central: ' + (err?.message || err), ui.ButtonSet.OK);
  }
}

/**
 * Menu interativo para verificar e aplicar as colunas de AS e CEOP na planilha.
 */
function menuGarantirColunasCeopAs() {
  const ui = SpreadsheetApp.getUi();
  try {
    const res = garantirColunasCeopAs_();
    if (res.adicionadas && res.adicionadas.length > 0) {
      ui.alert(
        'Estrutura Atualizada',
        'Foram adicionadas com sucesso ' + res.adicionadas.length + ' colunas para Autorização de Serviço (AS) na aba REGISTROS:\n\n• ' +
        res.adicionadas.join('\n• '),
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Estrutura Íntegra',
        'A aba REGISTROS já contém todas as colunas necessárias para Autorização de Serviço (AS). Nenhuma alteração foi necessária.',
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    ui.alert('Aviso', 'Falha ao verificar colunas: ' + (err?.message || err), ui.ButtonSet.OK);
  }
}

function lerConfigComoObjeto_(ss) {
  // R02 (auditoria S26.10) — cache de 60s; invalidado em setConfigValue_().
  try {
    const hit = CacheService.getScriptCache().get('MALL_CFG_V1');
    if (hit) return JSON.parse(hit);
  } catch (_) {}
  const sh = ss.getSheetByName('CONFIG'); if (!sh || sh.getLastRow() < 2) return {};
  const cfg = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().reduce((a, r) => { const k = String(r[0] || '').trim(); if (k) a[k] = String(r[1] ?? '').trim(); return a; }, {});
  try { CacheService.getScriptCache().put('MALL_CFG_V1', JSON.stringify(cfg), 60); } catch (_) {}
  return cfg;
}
function check_(arr, nome, ok, detalhe) { arr.push({ nome, ok: !!ok, detalhe: String(detalhe || '') }); }
function abrirPastaRaizS0() { const cfg = lerConfigComoObjeto_(SpreadsheetApp.getActive()); if (!cfg.DRIVE_ROOT_FOLDER_ID) return SpreadsheetApp.getUi().alert('Pasta raiz não configurada.'); const u = `https://drive.google.com/drive/folders/${cfg.DRIVE_ROOT_FOLDER_ID}`; SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<a href="${u}" target="_blank">Abrir pasta</a>`).setWidth(360).setHeight(90), 'Pasta raiz'); }
