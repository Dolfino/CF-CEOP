/**
 * SINALIZAÇÃO DO MALL / CEOP — MÓDULO DE GESTÃO DE PATRIMÔNIO E SERVIÇOS INTERNOS
 * Baseline: MVP-3.32.0-SINALIZACAO-S26.10
 *
 * Responsável pelo fluxo completo de:
 * 1. Registro e triagem de Ordens de Serviço (OS) de Patrimônio;
 * 2. Controle de intervenções físicas no box (divisórias, testeiras, MDF, portas, teto, etc.);
 * 3. Geração do documento oficial "CONTROLE DE PATRIMÔNIO – SERVIÇOS INTERNO" em PDF;
 * 4. Coleta e fixação de assinaturas digitais do cliente e representante CEOP.
 */

const PATRIMONIO_TABELA_NOME = 'PATRIMONIO_ORDENS_SERVICO';

const PATRIMONIO_COLUNAS_CANONICAS = Object.freeze([
  'ID_OS',
  'DATA_CRIACAO',
  'TAG_ORIGEM',
  'NOME_SOLICITANTE',
  'CPF_SOLICITANTE',
  'TELEFONE_SOLICITANTE',
  'EMAIL_SOLICITANTE',
  'SETOR',
  'RUA',
  'BOX_LOJA',
  'SERVICOS_JSON',
  'OBSERVACOES',
  'STATUS',
  'DATA_CONCLUSAO',
  'URL_TERMO_PDF',
  'ID_TERMO_DRIVE',
  'ASSINATURA_CLIENTE_DATA',
  'ASSINATURA_RESPONSAVEL_DATA',
  'OPERADOR_RESPONSAVEL',
  'ATUALIZADO_EM'
]);

const PATRIMONIO_TAGS_ORIGEM = Object.freeze([
  'SOLICITAÇÃO COMERCIAL',
  'SOLICITAÇÃO FINANCEIRO/COBRANÇA',
  'SOLICITAÇÃO PERMISSIONÁRIO',
  'SOLICITAÇÃO SETOR MARKETING',
  'SOLICITAÇÃO SETOR OPERACIONAL'
]);

const PATRIMONIO_CHECKLIST_SERVICOS = Object.freeze([
  'RETIRADA DE LATERAL / DIVISÓRIA DO BOX',
  'INSTALAÇÃO DE LATERAL / DIVISÓRIA DO BOX',
  'RETIRADA DE MDF (FUNDO DO BOX)',
  'INSTALAÇÃO DE MDF (FUNDO DO BOX)',
  'RETIRADA DE RABICHO ELÉTRICO DO BOX',
  'INSTALAÇÃO DE RABICHO ELÉTRICO DO BOX',
  'RETIRADA DE TESTEIRA DO BOX',
  'INSTALAÇÃO DE TESTEIRA DO BOX',
  'DESOCUPAÇÃO',
  'MANUTENÇÃO DA PORTA',
  'RETIRADA DE PORTA',
  'INSTALAÇÃO DE PORTA',
  'RETIRADA DE ESTRUTURA',
  'INSTALAÇÃO DE ESTRUTURA',
  'INSTALAÇÃO DE TAPUME',
  'RETIRADA DE TAPUME',
  'DESLOCAMENTO',
  'MONTAGEM DE MEGA BOX',
  'MONTAGEM DE BOX',
  'UNIFICAÇÃO DE BOX',
  'RETIRADA/INSTALAÇÃO/FACHADA',
  'RETIRADA/INSTALAÇÃO/TETO',
  'DESMONTAGEM DE MÓVEL',
  'DESMONTAGEM DE BOX',
  'RETIRADA DE PERTENCES'
]);

/**
 * Garante a existência e estrutura da aba PATRIMONIO_ORDENS_SERVICO na planilha.
 */
function garantirAbaPatrimonioOrdensServico_(ss) {
  const planilha = ss || SpreadsheetApp.getActive();
  let aba = planilha.getSheetByName(PATRIMONIO_TABELA_NOME);

  if (!aba) {
    aba = planilha.insertSheet(PATRIMONIO_TABELA_NOME);
    aba.setTabColor('#eab308'); // Tom amarelo/ouro para patrimônio
    aba.getRange(1, 1, 1, PATRIMONIO_COLUNAS_CANONICAS.length)
      .setValues([PATRIMONIO_COLUNAS_CANONICAS])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    aba.setFrozenRows(1);

    // Ajusta larguras iniciais
    aba.setColumnWidth(1, 140); // ID_OS
    aba.setColumnWidth(2, 160); // DATA_CRIACAO
    aba.setColumnWidth(3, 190); // TAG_ORIGEM
    aba.setColumnWidth(4, 220); // NOME_SOLICITANTE
    aba.setColumnWidth(5, 130); // CPF_SOLICITANTE
    aba.setColumnWidth(6, 140); // TELEFONE_SOLICITANTE
    aba.setColumnWidth(7, 200); // EMAIL_SOLICITANTE
    aba.setColumnWidth(8, 110); // SETOR
    aba.setColumnWidth(9, 180); // RUA
    aba.setColumnWidth(10, 110); // BOX_LOJA
    aba.setColumnWidth(11, 260); // SERVICOS_JSON
    aba.setColumnWidth(12, 260); // OBSERVACOES
    aba.setColumnWidth(13, 120); // STATUS
    aba.setColumnWidth(14, 130); // DATA_CONCLUSAO
    aba.setColumnWidth(15, 220); // URL_TERMO_PDF
    return aba;
  }

  // Verifica se faltam cabeçalhos
  const lastCol = aba.getLastColumn();
  if (lastCol < 1) {
    aba.getRange(1, 1, 1, PATRIMONIO_COLUNAS_CANONICAS.length)
      .setValues([PATRIMONIO_COLUNAS_CANONICAS])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    aba.setFrozenRows(1);
    return aba;
  }

  const cabecalhosExistentes = aba.getRange(1, 1, 1, lastCol).getValues()[0].map(c => String(c || '').trim());
  const faltantes = PATRIMONIO_COLUNAS_CANONICAS.filter(c => !cabecalhosExistentes.includes(c));

  if (faltantes.length > 0) {
    aba.getRange(1, lastCol + 1, 1, faltantes.length)
      .setValues([faltantes])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
  }

  return aba;
}

/**
 * Localiza ou cria pasta no Google Drive para termos assinados de Patrimônio.
 */
function obterPastaPatrimonioTermosNoDrive_(cfg) {
  let pastaRaiz = null;
  const rootId = String((cfg && cfg.DRIVE_ROOT_FOLDER_ID) || '').trim();

  if (rootId) {
    try { pastaRaiz = DriveApp.getFolderById(rootId); } catch (_) {}
  }

  if (!pastaRaiz) {
    try {
      const ssId = SpreadsheetApp.getActive().getId();
      const ssFile = DriveApp.getFileById(ssId);
      const parents = ssFile.getParents();
      if (parents.hasNext()) pastaRaiz = parents.next();
    } catch (_) {}
  }

  if (!pastaRaiz) {
    pastaRaiz = DriveApp.getRootFolder();
  }

  // Subpasta CEOP_PATRIMONIO
  let pastaPatrimonio = null;
  const itP = pastaRaiz.getFoldersByName('CEOP_PATRIMONIO');
  if (itP.hasNext()) {
    pastaPatrimonio = itP.next();
  } else {
    pastaPatrimonio = pastaRaiz.createFolder('CEOP_PATRIMONIO');
  }

  // Subpasta TERMOS_ASSINADOS
  let pastaTermos = null;
  const itT = pastaPatrimonio.getFoldersByName('TERMOS_ASSINADOS');
  if (itT.hasNext()) {
    pastaTermos = itT.next();
  } else {
    pastaTermos = pastaPatrimonio.createFolder('TERMOS_ASSINADOS');
  }

  return pastaTermos;
}

/**
 * Gera o próximo ID de Ordem de Serviço no formato OS-YYYY-XXXXX
 */
function gerarIdProximaOSPatrimonio_(aba) {
  const anoAtual = new Date().getFullYear();
  const lastRow = aba.getLastRow();
  if (lastRow <= 1) return `OS-${anoAtual}-00001`;

  const dados = aba.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxSeq = 0;
  const regex = new RegExp(`^OS-${anoAtual}-(\\d+)$`, 'i');

  for (let i = 0; i < dados.length; i++) {
    const val = String(dados[i][0] || '').trim();
    const match = val.match(regex);
    if (match) {
      const seq = parseInt(match[1], 10);
      if (seq > maxSeq) maxSeq = seq;
    }
  }

  const proximo = maxSeq + 1;
  return `OS-${anoAtual}-${String(proximo).padStart(5, '0')}`;
}

/**
 * Cria uma nova Ordem de Serviço de Patrimônio.
 * @param {Object} payload Dados recebidos do formulário de solicitação.
 */
function criarOrdemServicoPatrimonio(payload) {
  try {
    const ss = SpreadsheetApp.getActive();
    const aba = garantirAbaPatrimonioOrdensServico_(ss);
    const idOS = gerarIdProximaOSPatrimonio_(aba);
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const tagOrigem = String(payload.tagOrigem || 'SOLICITAÇÃO COMERCIAL').trim();
    const nomeSolicitante = String(payload.nomeSolicitante || '').trim().toUpperCase();
    const cpfSolicitante = String(payload.cpfSolicitante || '').replace(/\D/g, '');
    const telefoneSolicitante = String(payload.telefoneSolicitante || '').trim();
    const emailSolicitante = String(payload.emailSolicitante || '').trim().toLowerCase();
    const setor = String(payload.setor || '').trim().toUpperCase();
    const rua = String(payload.rua || '').trim();
    const boxLoja = String(payload.boxLoja || '').trim().toUpperCase();
    const servicos = Array.isArray(payload.servicos) ? payload.servicos : [];
    const observacoes = String(payload.observacoes || '').trim();
    const operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'operador@centrofashion.com.br';

    if (!nomeSolicitante) throw new Error('O nome do solicitante é obrigatório.');
    if (!setor) throw new Error('O setor é obrigatório.');
    if (!boxLoja) throw new Error('O número do box/loja é obrigatório.');
    if (servicos.length === 0 && !observacoes) throw new Error('Selecione ao menos um serviço ou adicione observações.');

    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const novaLinha = new Array(cabecalhos.length).fill('');

    function setVal(col, val) {
      const idx = cabecalhos.indexOf(col);
      if (idx >= 0) novaLinha[idx] = val;
    }

    setVal('ID_OS', idOS);
    setVal('DATA_CRIACAO', agora);
    setVal('TAG_ORIGEM', tagOrigem);
    setVal('NOME_SOLICITANTE', nomeSolicitante);
    setVal('CPF_SOLICITANTE', cpfSolicitante);
    setVal('TELEFONE_SOLICITANTE', telefoneSolicitante);
    setVal('EMAIL_SOLICITANTE', emailSolicitante);
    setVal('SETOR', setor);
    setVal('RUA', rua);
    setVal('BOX_LOJA', boxLoja);
    setVal('SERVICOS_JSON', JSON.stringify(servicos));
    setVal('OBSERVACOES', observacoes);
    setVal('STATUS', 'PENDENTE');
    setVal('DATA_CONCLUSAO', '');
    setVal('URL_TERMO_PDF', '');
    setVal('ID_TERMO_DRIVE', '');
    setVal('ASSINATURA_CLIENTE_DATA', '');
    setVal('ASSINATURA_RESPONSAVEL_DATA', '');
    setVal('OPERADOR_RESPONSAVEL', operador);
    setVal('ATUALIZADO_EM', agora);

    aba.appendRow(novaLinha);

    return {
      ok: true,
      idOS: idOS,
      mensagem: `Ordem de Serviço ${idOS} registrada com sucesso.`,
      dados: {
        idOS,
        dataCriacao: agora,
        tagOrigem,
        nomeSolicitante,
        cpfSolicitante,
        telefoneSolicitante,
        emailSolicitante,
        setor,
        rua,
        boxLoja,
        servicos,
        observacoes,
        status: 'PENDENTE'
      }
    };
  } catch (err) {
    console.error('[PATRIMONIO] Erro ao criar OS:', err);
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Lista as Ordens de Serviço de Patrimônio com suporte a filtros.
 */
function listarOrdensServicoPatrimonio(filtro) {
  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(PATRIMONIO_TABELA_NOME);
    if (!aba || aba.getLastRow() <= 1) return { ok: true, total: 0, itens: [] };

    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    const cabecalhos = aba.getRange(1, 1, 1, lastCol).getValues()[0].map(c => String(c || '').trim());
    const dados = aba.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const f = filtro || {};
    const statusFiltro = String(f.status || '').trim().toUpperCase();
    const setorFiltro = String(f.setor || '').trim().toUpperCase();
    const buscaFiltro = String(f.busca || '').trim().toUpperCase();

    const itens = [];
    for (let i = 0; i < dados.length; i++) {
      const row = dados[i];
      const item = {};
      cabecalhos.forEach((col, idx) => {
        item[col] = row[idx];
      });

      // Filtro por status
      if (statusFiltro && String(item.STATUS || '').toUpperCase() !== statusFiltro) continue;
      // Filtro por setor
      if (setorFiltro && String(item.SETOR || '').toUpperCase() !== setorFiltro) continue;
      // Busca livre (ID, Solicitante, Box, Rua)
      if (buscaFiltro) {
        const strBusca = `${item.ID_OS} ${item.NOME_SOLICITANTE} ${item.CPF_SOLICITANTE} ${item.BOX_LOJA} ${item.RUA}`.toUpperCase();
        if (!strBusca.includes(buscaFiltro)) continue;
      }

      // Converte serviços JSON
      let servicosParsed = [];
      try {
        servicosParsed = JSON.parse(item.SERVICOS_JSON || '[]');
      } catch (_) {
        servicosParsed = item.SERVICOS_JSON ? [item.SERVICOS_JSON] : [];
      }
      item.SERVICOS_ARRAY = servicosParsed;

      itens.push(item);
    }

    // Ordena decrescente por ID_OS / Data
    itens.reverse();

    return {
      ok: true,
      total: itens.length,
      itens: itens
    };
  } catch (err) {
    console.error('[PATRIMONIO] Erro ao listar OS:', err);
    return { ok: false, error: err.message || String(err), itens: [] };
  }
}

/**
 * Obtém os detalhes de uma OS específica por ID.
 */
function obterOrdemServicoPatrimonioPorId(idOS) {
  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(PATRIMONIO_TABELA_NOME);
    if (!aba || aba.getLastRow() <= 1) return { ok: false, error: 'Tabela de patrimônio vazia ou não encontrada.' };

    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    const cabecalhos = aba.getRange(1, 1, 1, lastCol).getValues()[0].map(c => String(c || '').trim());
    const dados = aba.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const idxId = cabecalhos.indexOf('ID_OS');
    if (idxId < 0) return { ok: false, error: 'Coluna ID_OS ausente.' };

    for (let i = 0; i < dados.length; i++) {
      if (String(dados[i][idxId] || '').trim().toUpperCase() === String(idOS || '').trim().toUpperCase()) {
        const item = {};
        cabecalhos.forEach((col, idx) => {
          item[col] = dados[i][idx];
        });
        try {
          item.SERVICOS_ARRAY = JSON.parse(item.SERVICOS_JSON || '[]');
        } catch (_) {
          item.SERVICOS_ARRAY = [];
        }
        return { ok: true, item: item, rowIndex: i + 2 };
      }
    }

    return { ok: false, error: `Ordem de Serviço ${idOS} não encontrada.` };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Atualiza o status e notas de uma OS de Patrimônio.
 */
function atualizarStatusOrdemPatrimonio(idOS, novoStatus, observacoesAdicionais, dataConclusao) {
  try {
    const res = obterOrdemServicoPatrimonioPorId(idOS);
    if (!res.ok) return res;

    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(PATRIMONIO_TABELA_NOME);
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const row = res.rowIndex;
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const idxStatus = cabecalhos.indexOf('STATUS');
    const idxObs = cabecalhos.indexOf('OBSERVACOES');
    const idxDataConc = cabecalhos.indexOf('DATA_CONCLUSAO');
    const idxAtualizado = cabecalhos.indexOf('ATUALIZADO_EM');

    if (idxStatus >= 0 && novoStatus) aba.getRange(row, idxStatus + 1).setValue(novoStatus);
    if (idxDataConc >= 0 && dataConclusao) aba.getRange(row, idxDataConc + 1).setValue(dataConclusao);
    if (idxAtualizado >= 0) aba.getRange(row, idxAtualizado + 1).setValue(agora);

    if (idxObs >= 0 && observacoesAdicionais) {
      const obsAtual = String(aba.getRange(row, idxObs + 1).getValue() || '');
      const novaObs = obsAtual ? `${obsAtual} | [${agora}] ${observacoesAdicionais}` : `[${agora}] ${observacoesAdicionais}`;
      aba.getRange(row, idxObs + 1).setValue(novaObs);
    }

    return { ok: true, mensagem: `Status da OS ${idOS} atualizado para ${novoStatus}.` };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Gera o documento oficial "CONTROLE DE PATRIMÔNIO – SERVIÇOS INTERNO" em PDF com assinaturas digitais.
 * @param {string} idOS Identificador da OS.
 * @param {string} assinaturaClienteBase64 DataURL (PNG) da assinatura do cliente.
 * @param {string} assinaturaRepresentanteBase64 DataURL (PNG) da assinatura do fiscal/CEOP.
 */
function gerarTermoControlePatrimonioPdf(idOS, assinaturaClienteBase64, assinaturaRepresentanteBase64) {
  try {
    const res = obterOrdemServicoPatrimonioPorId(idOS);
    if (!res.ok) return res;

    const d = res.item;
    const ss = SpreadsheetApp.getActive();
    const cfg = lerConfigComoObjeto_(ss);
    const pastaTermos = obterPastaPatrimonioTermosNoDrive_(cfg);

    const dataServicoFormatada = d.DATA_CONCLUSAO || Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', 'dd/MM/yyyy');
    const hoje = new Date();
    const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    const dataExtenso = `${hoje.getDate()} de ${meses[hoje.getMonth()]} de ${hoje.getFullYear()}`;

    const servicosTexto = (d.SERVICOS_ARRAY && d.SERVICOS_ARRAY.length > 0)
      ? d.SERVICOS_ARRAY.join('; ')
      : 'SERVIÇOS DE PATRIMÔNIO CONFORME SOLICITAÇÃO';

    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @page { size: A4; margin: 25mm 20mm; }
        body {
          font-family: Arial, sans-serif;
          color: #000;
          font-size: 13px;
          line-height: 1.5;
          margin: 0;
          padding: 20px 40px;
        }
        .header {
          text-align: left;
          margin-bottom: 25px;
        }
        .logo-text {
          font-size: 18px;
          font-weight: bold;
          color: #171b68;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .logo-sub {
          font-size: 10px;
          font-weight: bold;
          color: #e11d48;
          letter-spacing: 1px;
        }
        h2 {
          text-align: center;
          font-size: 15px;
          font-weight: bold;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-top: 30px;
          margin-bottom: 35px;
          text-decoration: underline;
        }
        .field-group {
          margin-bottom: 14px;
          font-size: 13px;
        }
        .field-label {
          font-weight: bold;
          color: #111;
        }
        .field-value {
          font-weight: bold;
          text-decoration: underline;
          color: #111;
        }
        .declaration-box {
          margin-top: 30px;
          margin-bottom: 30px;
          font-size: 13px;
          text-align: justify;
          line-height: 1.6;
        }
        .date-line {
          margin-top: 25px;
          margin-bottom: 25px;
          font-size: 13px;
        }
        .signer-info {
          margin-bottom: 12px;
          font-size: 13px;
        }
        .signatures-container {
          margin-top: 40px;
        }
        .sig-block {
          margin-bottom: 30px;
        }
        .sig-line {
          border-top: 1px solid #000;
          width: 380px;
          margin-top: 35px;
          margin-bottom: 4px;
        }
        .sig-title {
          font-size: 12px;
          font-weight: bold;
        }
        .sig-img {
          max-height: 55px;
          max-width: 260px;
          display: block;
          margin-bottom: -15px;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="logo-text">Centro Fashion</div>
        <div class="logo-sub">FORTALEZA</div>
      </div>

      <h2>CONTROLE DE PATRIMÔNIO – SERVIÇOS INTERNO</h2>

      <div class="field-group">
        <span class="field-label">Ref. Retirada/Colocação de:</span><br>
        <span class="field-value">${servicosTexto}</span>
      </div>

      <div class="field-group">
        <span class="field-label">Setor:</span> <span class="field-value">${d.SETOR || '-'}</span>
      </div>

      <div class="field-group">
        <span class="field-label">Rua:</span> <span class="field-value">${d.RUA || '-'}</span>
      </div>

      <div class="field-group">
        <span class="field-label">Número:</span> <span class="field-value">${d.BOX_LOJA || '-'}</span>
      </div>

      <div class="field-group">
        <span class="field-label">Data do serviço:</span> <span class="field-value">${dataServicoFormatada}</span>
      </div>

      <div class="field-group">
        <span class="field-label">OBS.:</span> <span class="field-value">${d.OBSERVACOES || 'CONFORME REGISTRO DA ORDEM DE SERVIÇO ' + d.ID_OS}</span>
      </div>

      <div class="declaration-box">
        Pelo presente, declaro que não tenho interesse no(s) item(s) marcado(s) acima e que o(s) devolvo ao Centro Fashion Fortaleza. Tenho ciência de que o(s) material(is) pertence(m) ao empreendimento e só podem ser usados nas dependências do mesmo.
      </div>

      <div class="date-line">
        Fortaleza, ${dataExtenso}
      </div>

      <div class="signer-info">
        <strong>Nome:</strong> <span class="field-value">${d.NOME_SOLICITANTE || '-'}</span>
      </div>

      <div class="signer-info">
        <strong>RG/CPF:</strong> <span class="field-value">${d.CPF_SOLICITANTE || '-'}</span>
      </div>

      <div class="signer-info">
        <strong>E-mail:</strong> <span class="field-value">${d.EMAIL_SOLICITANTE || '-'}</span>
      </div>

      <div class="signer-info">
        <strong>Telefone:</strong> <span class="field-value">${d.TELEFONE_SOLICITANTE || '-'}</span>
      </div>

      <div class="signatures-container">
        <div class="sig-block">
          ${assinaturaClienteBase64 ? `<img src="${assinaturaClienteBase64}" class="sig-img" alt="Assinatura do Cliente" />` : ''}
          <div class="sig-line"></div>
          <div class="sig-title">Assinatura do Cliente:</div>
        </div>

        <div class="sig-block">
          ${assinaturaRepresentanteBase64 ? `<img src="${assinaturaRepresentanteBase64}" class="sig-img" alt="Assinatura Representante" />` : ''}
          <div class="sig-line"></div>
          <div class="sig-title">Assinatura do Representante Centro Fashion:</div>
        </div>
      </div>
    </body>
    </html>
    `;

    const blobHtml = Utilities.newBlob(html, 'text/html', `Termo_Patrimonio_${idOS}.html`);
    const pdfBlob = blobHtml.getAs('application/pdf').setName(`Termo_Patrimonio_${idOS}.pdf`);
    const arquivoPdf = pastaTermos.createFile(pdfBlob);
    arquivoPdf.setDescription(`Termo Oficial de Controle de Patrimônio da OS ${idOS} - Centro Fashion Fortaleza`);

    // Atualiza a planilha com os links
    const aba = ss.getSheetByName(PATRIMONIO_TABELA_NOME);
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const row = res.rowIndex;
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const setVal = (col, val) => {
      const idx = cabecalhos.indexOf(col);
      if (idx >= 0) aba.getRange(row, idx + 1).setValue(val);
    };

    setVal('URL_TERMO_PDF', arquivoPdf.getUrl());
    setVal('ID_TERMO_DRIVE', arquivoPdf.getId());
    setVal('STATUS', 'CONCLUIDO');
    setVal('DATA_CONCLUSAO', dataServicoFormatada);
    setVal('ATUALIZADO_EM', agora);
    if (assinaturaClienteBase64) setVal('ASSINATURA_CLIENTE_DATA', 'COLETADA_DIGITALMENTE');
    if (assinaturaRepresentanteBase64) setVal('ASSINATURA_RESPONSAVEL_DATA', 'COLETADA_DIGITALMENTE');

    return {
      ok: true,
      idOS: idOS,
      urlPdf: arquivoPdf.getUrl(),
      downloadUrl: arquivoPdf.getDownloadUrl(),
      driveId: arquivoPdf.getId(),
      mensagem: `Termo de Controle de Patrimônio da OS ${idOS} gerado com sucesso em PDF.`
    };
  } catch (err) {
    console.error('[PATRIMONIO] Erro ao gerar Termo PDF:', err);
    return { ok: false, error: err.message || String(err) };
  }
}
