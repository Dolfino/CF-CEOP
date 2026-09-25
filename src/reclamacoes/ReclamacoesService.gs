/**
 * SINALIZAÇÃO DO MALL / CEOP — MÓDULO DE GESTÃO DE RECLAMAÇÕES (SAC / OUVIDORIA)
 * Baseline: MVP-3.32.0-SINALIZACAO-S26.10
 *
 * Responsável pelo fluxo completo de:
 * 1. Abertura de chamados por Lojistas/Permissionários ou Clientes Visitantes;
 * 2. Upload e custódia segura de mídias probatórias no Google Drive (Fotos, Vídeos e Áudios);
 * 3. Triagem, investigação e registro de medidas tomadas de resolução;
 * 4. Emissão do "Termo de Encerramento de Chamado de Reclamação" com Código Pipe e quitação mútua;
 * 5. Coleta de assinaturas eletrônicas do reclamante e da diretoria/CEOP.
 */

const RECLAMACOES_TABELA_NOME = 'CHAMADOS_RECLAMACOES';
const RECLAMACOES_ANEXOS_TABELA_NOME = 'CHAMADOS_ANEXOS';

const RECLAMACOES_COLUNAS_CANONICAS = Object.freeze([
  'CODIGO_PIPE',
  'DATA_ABERTURA',
  'TIPO_RECLAMANTE',
  'NOME_RECLAMANTE',
  'CPF_RECLAMANTE',
  'TELEFONE_RECLAMANTE',
  'EMAIL_RECLAMANTE',
  'SETOR',
  'RUA',
  'BOX_LOJA',
  'CATEGORIA',
  'DATA_OCORRENCIA',
  'DESCRICAO_RECLAMACAO',
  'ACEITE_LGPD',
  'STATUS',
  'DATA_RESOLUCAO',
  'MEDIDAS_TOMADAS',
  'URL_TERMO_ENCERRAMENTO',
  'ID_TERMO_DRIVE',
  'ASSINATURA_SOLICITANTE_DATA',
  'ASSINATURA_REPRESENTANTE_DATA',
  'OPERADOR_RESOLUCAO',
  'ATUALIZADO_EM'
]);

const RECLAMACOES_ANEXOS_COLUNAS_CANONICAS = Object.freeze([
  'ID_ANEXO',
  'CODIGO_PIPE',
  'TIPO_MIDIA',
  'NOME_ARQUIVO',
  'MIME_TYPE',
  'DRIVE_FILE_ID',
  'DRIVE_VIEW_URL',
  'CRIADO_EM'
]);

const RECLAMACOES_CATEGORIAS_OFICIAIS = Object.freeze([
  'USO INDEVIDO DO ESPAÇO',
  'FALTA DE ENERGIA',
  'RELACIONADOS A CHUVAS',
  'CAUSADOS POR PETS',
  'LIMPEZA E CONSERVAÇÃO',
  'OCORRÊNCIA DA BRIGADA',
  'INCIDENTE COM PRESTADORES DE SERVIÇOS',
  'ESTRUTURAL',
  'ATENDIMENTO',
  'QUEDA DE ENERGIA',
  'ACHADOS E PERDIDOS',
  'VAZAMENTO CAUSADO PELO AR CONDICIONADO'
]);

/**
 * Garante as abas CHAMADOS_RECLAMACOES e CHAMADOS_ANEXOS na planilha ativa.
 */
function garantirAbasReclamacoes_(ss) {
  const planilha = ss || SpreadsheetApp.getActive();

  // 1. Aba Principal de Reclamações
  let abaRec = planilha.getSheetByName(RECLAMACOES_TABELA_NOME);
  if (!abaRec) {
    abaRec = planilha.insertSheet(RECLAMACOES_TABELA_NOME);
    abaRec.setTabColor('#e11d48'); // Tom rubi/rosa corporativo
    abaRec.getRange(1, 1, 1, RECLAMACOES_COLUNAS_CANONICAS.length)
      .setValues([RECLAMACOES_COLUNAS_CANONICAS])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    abaRec.setFrozenRows(1);

    abaRec.setColumnWidth(1, 140); // CODIGO_PIPE
    abaRec.setColumnWidth(2, 160); // DATA_ABERTURA
    abaRec.setColumnWidth(3, 170); // TIPO_RECLAMANTE
    abaRec.setColumnWidth(4, 220); // NOME_RECLAMANTE
    abaRec.setColumnWidth(5, 130); // CPF_RECLAMANTE
    abaRec.setColumnWidth(6, 140); // TELEFONE_RECLAMANTE
    abaRec.setColumnWidth(7, 200); // EMAIL_RECLAMANTE
    abaRec.setColumnWidth(8, 110); // SETOR
    abaRec.setColumnWidth(9, 180); // RUA
    abaRec.setColumnWidth(10, 110); // BOX_LOJA
    abaRec.setColumnWidth(11, 190); // CATEGORIA
    abaRec.setColumnWidth(12, 130); // DATA_OCORRENCIA
    abaRec.setColumnWidth(13, 280); // DESCRICAO_RECLAMACAO
    abaRec.setColumnWidth(14, 110); // ACEITE_LGPD
    abaRec.setColumnWidth(15, 120); // STATUS
    abaRec.setColumnWidth(16, 130); // DATA_RESOLUCAO
    abaRec.setColumnWidth(17, 280); // MEDIDAS_TOMADAS
    abaRec.setColumnWidth(18, 220); // URL_TERMO_ENCERRAMENTO
  }

  // 2. Aba de Anexos
  let abaAnx = planilha.getSheetByName(RECLAMACOES_ANEXOS_TABELA_NOME);
  if (!abaAnx) {
    abaAnx = planilha.insertSheet(RECLAMACOES_ANEXOS_TABELA_NOME);
    abaAnx.setTabColor('#f43f5e');
    abaAnx.getRange(1, 1, 1, RECLAMACOES_ANEXOS_COLUNAS_CANONICAS.length)
      .setValues([RECLAMACOES_ANEXOS_COLUNAS_CANONICAS])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    abaAnx.setFrozenRows(1);

    abaAnx.setColumnWidth(1, 130); // ID_ANEXO
    abaAnx.setColumnWidth(2, 140); // CODIGO_PIPE
    abaAnx.setColumnWidth(3, 120); // TIPO_MIDIA
    abaAnx.setColumnWidth(4, 220); // NOME_ARQUIVO
    abaAnx.setColumnWidth(5, 140); // MIME_TYPE
    abaAnx.setColumnWidth(6, 190); // DRIVE_FILE_ID
    abaAnx.setColumnWidth(7, 240); // DRIVE_VIEW_URL
    abaAnx.setColumnWidth(8, 160); // CRIADO_EM
  }

  return { abaRec, abaAnx };
}

/**
 * Localiza ou cria as pastas no Google Drive para Reclamações.
 */
function obterPastasReclamacoesNoDrive_(cfg) {
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

  // Pasta CEOP_RECLAMACOES
  let pastaPrincipal = null;
  const itP = pastaRaiz.getFoldersByName('CEOP_RECLAMACOES');
  if (itP.hasNext()) {
    pastaPrincipal = itP.next();
  } else {
    pastaPrincipal = pastaRaiz.createFolder('CEOP_RECLAMACOES');
  }

  // Subpasta ANEXOS_MIDIA
  let pastaAnexos = null;
  const itA = pastaPrincipal.getFoldersByName('ANEXOS_MIDIA');
  if (itA.hasNext()) {
    pastaAnexos = itA.next();
  } else {
    pastaAnexos = pastaPrincipal.createFolder('ANEXOS_MIDIA');
  }

  // Subpasta TERMOS_ENCERRAMENTO
  let pastaTermos = null;
  const itT = pastaPrincipal.getFoldersByName('TERMOS_ENCERRAMENTO');
  if (itT.hasNext()) {
    pastaTermos = itT.next();
  } else {
    pastaTermos = pastaPrincipal.createFolder('TERMOS_ENCERRAMENTO');
  }

  return { pastaPrincipal, pastaAnexos, pastaTermos };
}

/**
 * Gera um Código Pipe / Protocolo numérico exclusivo de 10 dígitos (ex: 1414432738).
 */
function gerarCodigoPipeReclamacao_() {
  const timestampPart = String(Date.now()).slice(-8);
  const randomPart = String(Math.floor(Math.random() * 90 + 10));
  return `${randomPart}${timestampPart}`;
}

/**
 * Registra um novo chamado de reclamação com upload opcional de fotos/áudio/vídeo.
 * @param {Object} payload Dados do chamado (nome, cpf, contato, setor, rua, box, relato, aceiteLGPD, etc.)
 * @param {Array<Object>} anexos Lista de arquivos { nome, mimeType, base64, tipoMidia: 'IMAGEM'|'AUDIO'|'VIDEO' }
 */
function criarChamadoReclamacao(payload, anexos) {
  try {
    const ss = SpreadsheetApp.getActive();
    const { abaRec, abaAnx } = garantirAbasReclamacoes_(ss);
    const codigoPipe = gerarCodigoPipeReclamacao_();
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const tipoReclamante = String(payload.tipoReclamante || 'PERMISSIONARIO/LOJISTA').trim().toUpperCase();
    const nomeReclamante = String(payload.nomeReclamante || '').trim().toUpperCase();
    const cpfReclamante = String(payload.cpfReclamante || '').replace(/\D/g, '');
    const telefoneReclamante = String(payload.telefoneReclamante || '').trim();
    const emailReclamante = String(payload.emailReclamante || '').trim().toLowerCase();
    const setor = String(payload.setor || '').trim().toUpperCase();
    const rua = String(payload.rua || '').trim();
    const boxLoja = String(payload.boxLoja || '').trim().toUpperCase();
    const categoria = String(payload.categoria || 'OUTROS').trim().toUpperCase();
    const dataOcorrencia = String(payload.dataOcorrencia || Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', 'dd/MM/yyyy')).trim();
    const descricao = String(payload.descricaoReclamacao || '').trim();
    const aceiteLGPD = payload.aceiteLGPD ? 'SIM' : 'NÃO';

    if (!nomeReclamante) throw new Error('O nome do reclamante é obrigatório.');
    if (!descricao) throw new Error('A descrição da reclamação é obrigatória.');
    if (payload.aceiteLGPD !== true && payload.aceiteLGPD !== 'SIM') {
      throw new Error('É necessário concordar com os termos da LGPD (Lei 13.709/2018).');
    }

    // Grava o chamado na aba CHAMADOS_RECLAMACOES
    const cabecalhos = abaRec.getRange(1, 1, 1, abaRec.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const novaLinha = new Array(cabecalhos.length).fill('');

    function setVal(col, val) {
      const idx = cabecalhos.indexOf(col);
      if (idx >= 0) novaLinha[idx] = val;
    }

    setVal('CODIGO_PIPE', codigoPipe);
    setVal('DATA_ABERTURA', agora);
    setVal('TIPO_RECLAMANTE', tipoReclamante);
    setVal('NOME_RECLAMANTE', nomeReclamante);
    setVal('CPF_RECLAMANTE', cpfReclamante);
    setVal('TELEFONE_RECLAMANTE', telefoneReclamante);
    setVal('EMAIL_RECLAMANTE', emailReclamante);
    setVal('SETOR', setor);
    setVal('RUA', rua);
    setVal('BOX_LOJA', boxLoja);
    setVal('CATEGORIA', categoria);
    setVal('DATA_OCORRENCIA', dataOcorrencia);
    setVal('DESCRICAO_RECLAMACAO', descricao);
    setVal('ACEITE_LGPD', aceiteLGPD);
    setVal('STATUS', 'NOVO');
    setVal('DATA_RESOLUCAO', '');
    setVal('MEDIDAS_TOMADAS', '');
    setVal('URL_TERMO_ENCERRAMENTO', '');
    setVal('ID_TERMO_DRIVE', '');
    setVal('ASSINATURA_SOLICITANTE_DATA', '');
    setVal('ASSINATURA_REPRESENTANTE_DATA', '');
    setVal('OPERADOR_RESOLUCAO', '');
    setVal('ATUALIZADO_EM', agora);

    abaRec.appendRow(novaLinha);

    // Processa upload de anexos de mídia se fornecidos
    const anexosSalvos = [];
    if (Array.isArray(anexos) && anexos.length > 0) {
      const cfg = lerConfigComoObjeto_(ss);
      const { pastaAnexos } = obterPastasReclamacoesNoDrive_(cfg);

      const cabAnx = abaAnx.getRange(1, 1, 1, abaAnx.getLastColumn()).getValues()[0].map(c => String(c || '').trim());

      for (let i = 0; i < anexos.length; i++) {
        const item = anexos[i];
        if (!item || !item.base64) continue;

        try {
          const nomeArq = item.nome || `anexo_${codigoPipe}_${i + 1}`;
          const mime = item.mimeType || 'application/octet-stream';
          const bytes = Utilities.base64Decode(item.base64.replace(/^data:[^;]+;base64,/, ''));
          const blob = Utilities.newBlob(bytes, mime, nomeArq);
          const arquivo = pastaAnexos.createFile(blob);
          arquivo.setDescription(`Anexo da Reclamação Código Pipe ${codigoPipe}`);

          let tipoMidia = item.tipoMidia || 'IMAGEM';
          if (mime.startsWith('audio/')) tipoMidia = 'AUDIO';
          else if (mime.startsWith('video/')) tipoMidia = 'VIDEO';
          else if (mime.startsWith('image/')) tipoMidia = 'IMAGEM';

          const idAnexo = `ANX-${codigoPipe}-${i + 1}`;
          const linhaAnexo = new Array(cabAnx.length).fill('');

          const setAnx = (col, val) => {
            const idx = cabAnx.indexOf(col);
            if (idx >= 0) linhaAnexo[idx] = val;
          };

          setAnx('ID_ANEXO', idAnexo);
          setAnx('CODIGO_PIPE', codigoPipe);
          setAnx('TIPO_MIDIA', tipoMidia);
          setAnx('NOME_ARQUIVO', nomeArq);
          setAnx('MIME_TYPE', mime);
          setAnx('DRIVE_FILE_ID', arquivo.getId());
          setAnx('DRIVE_VIEW_URL', arquivo.getUrl());
          setAnx('CRIADO_EM', agora);

          abaAnx.appendRow(linhaAnexo);

          anexosSalvos.push({
            idAnexo,
            tipoMidia,
            nomeArquivo: nomeArq,
            url: arquivo.getUrl()
          });
        } catch (eAnx) {
          console.error('[RECLAMACOES] Falha ao salvar anexo:', eAnx);
        }
      }
    }

    return {
      ok: true,
      codigoPipe: codigoPipe,
      mensagem: `Reclamação registrada com sucesso! Código Pipe: ${codigoPipe}`,
      anexosSalvos: anexosSalvos
    };
  } catch (err) {
    console.error('[RECLAMACOES] Erro ao criar chamado:', err);
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Lista todos os chamados de reclamações com suporte a filtros e busca.
 */
function listarChamadosReclamacoes(filtro) {
  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(RECLAMACOES_TABELA_NOME);
    if (!aba || aba.getLastRow() <= 1) return { ok: true, total: 0, itens: [] };

    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    const cabecalhos = aba.getRange(1, 1, 1, lastCol).getValues()[0].map(c => String(c || '').trim());
    const dados = aba.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const f = filtro || {};
    const statusFiltro = String(f.status || '').trim().toUpperCase();
    const categoriaFiltro = String(f.categoria || '').trim().toUpperCase();
    const setorFiltro = String(f.setor || '').trim().toUpperCase();
    const buscaFiltro = String(f.busca || '').trim().toUpperCase();

    const itens = [];
    for (let i = 0; i < dados.length; i++) {
      const row = dados[i];
      const item = {};
      cabecalhos.forEach((col, idx) => {
        item[col] = row[idx];
      });

      if (statusFiltro && String(item.STATUS || '').toUpperCase() !== statusFiltro) continue;
      if (categoriaFiltro && String(item.CATEGORIA || '').toUpperCase() !== categoriaFiltro) continue;
      if (setorFiltro && String(item.SETOR || '').toUpperCase() !== setorFiltro) continue;

      if (buscaFiltro) {
        const str = `${item.CODIGO_PIPE} ${item.NOME_RECLAMANTE} ${item.CPF_RECLAMANTE} ${item.BOX_LOJA} ${item.CATEGORIA}`.toUpperCase();
        if (!str.includes(buscaFiltro)) continue;
      }

      itens.push(item);
    }

    itens.reverse();

    return {
      ok: true,
      total: itens.length,
      itens: itens
    };
  } catch (err) {
    console.error('[RECLAMACOES] Erro ao listar chamados:', err);
    return { ok: false, error: err.message || String(err), itens: [] };
  }
}

/**
 * Obtém detalhes completos do chamado pelo Código Pipe, incluindo anexos.
 */
function obterDetalhesChamadoReclamacao(codigoPipe) {
  try {
    const ss = SpreadsheetApp.getActive();
    const abaRec = ss.getSheetByName(RECLAMACOES_TABELA_NOME);
    if (!abaRec || abaRec.getLastRow() <= 1) return { ok: false, error: 'Tabela de reclamações vazia.' };

    const lastRow = abaRec.getLastRow();
    const lastCol = abaRec.getLastColumn();
    const cabRec = abaRec.getRange(1, 1, 1, lastCol).getValues()[0].map(c => String(c || '').trim());
    const dadosRec = abaRec.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const idxPipe = cabRec.indexOf('CODIGO_PIPE');
    if (idxPipe < 0) return { ok: false, error: 'Coluna CODIGO_PIPE não encontrada.' };

    let chamado = null;
    let rowIndex = -1;

    for (let i = 0; i < dadosRec.length; i++) {
      if (String(dadosRec[i][idxPipe] || '').trim() === String(codigoPipe || '').trim()) {
        chamado = {};
        cabRec.forEach((col, idx) => {
          chamado[col] = dadosRec[i][idx];
        });
        rowIndex = i + 2;
        break;
      }
    }

    if (!chamado) return { ok: false, error: `Chamado com Código Pipe ${codigoPipe} não encontrado.` };

    // Busca os anexos vinculados
    const abaAnx = ss.getSheetByName(RECLAMACOES_ANEXOS_TABELA_NOME);
    const anexos = [];
    if (abaAnx && abaAnx.getLastRow() > 1) {
      const colAnx = abaAnx.getLastColumn();
      const cabAnx = abaAnx.getRange(1, 1, 1, colAnx).getValues()[0].map(c => String(c || '').trim());
      const dadosAnx = abaAnx.getRange(2, 1, abaAnx.getLastRow() - 1, colAnx).getValues();
      const idxPipeAnx = cabAnx.indexOf('CODIGO_PIPE');

      for (let j = 0; j < dadosAnx.length; j++) {
        if (String(dadosAnx[j][idxPipeAnx] || '').trim() === String(codigoPipe || '').trim()) {
          const itemAnx = {};
          cabAnx.forEach((c, idx) => {
            itemAnx[c] = dadosAnx[j][idx];
          });
          anexos.push(itemAnx);
        }
      }
    }

    chamado.ANEXOS = anexos;

    return {
      ok: true,
      chamado: chamado,
      rowIndex: rowIndex
    };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Atualiza o status e as medidas tomadas de resolução de um chamado.
 */
function atualizarResolucaoChamado(codigoPipe, dadosResolucao) {
  try {
    const res = obterDetalhesChamadoReclamacao(codigoPipe);
    if (!res.ok) return res;

    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(RECLAMACOES_TABELA_NOME);
    const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const row = res.rowIndex;
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const idxStatus = cab.indexOf('STATUS');
    const idxMedidas = cab.indexOf('MEDIDAS_TOMADAS');
    const idxDataRes = cab.indexOf('DATA_RESOLUCAO');
    const idxOp = cab.indexOf('OPERADOR_RESOLUCAO');
    const idxAtu = cab.indexOf('ATUALIZADO_EM');

    const novoStatus = dadosResolucao.status || 'RESOLVIDO';
    const medidas = String(dadosResolucao.medidasTomadas || '').trim();
    const dataRes = dadosResolucao.dataResolucao || Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', 'dd/MM/yyyy');
    const operador = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'operador@centrofashion.com.br';

    if (idxStatus >= 0) aba.getRange(row, idxStatus + 1).setValue(novoStatus);
    if (idxMedidas >= 0 && medidas) aba.getRange(row, idxMedidas + 1).setValue(medidas);
    if (idxDataRes >= 0) aba.getRange(row, idxDataRes + 1).setValue(dataRes);
    if (idxOp >= 0) aba.getRange(row, idxOp + 1).setValue(operador);
    if (idxAtu >= 0) aba.getRange(row, idxAtu + 1).setValue(agora);

    return { ok: true, mensagem: `Chamado ${codigoPipe} atualizado com sucesso.` };
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

/**
 * Gera o documento oficial "Termo de Encerramento de Chamado de Reclamação" em PDF com assinaturas digitais.
 * @param {string} codigoPipe Código Pipe do chamado.
 * @param {string} assinaturaSolicitanteBase64 Assinatura do cliente/lojista (PNG Base64).
 * @param {string} assinaturaRepresentanteBase64 Assinatura do gestor CEOP (PNG Base64).
 */
function gerarTermoEncerramentoReclamacaoPdf(codigoPipe, assinaturaSolicitanteBase64, assinaturaRepresentanteBase64) {
  try {
    const res = obterDetalhesChamadoReclamacao(codigoPipe);
    if (!res.ok) return res;

    const d = res.chamado;
    const ss = SpreadsheetApp.getActive();
    const cfg = lerConfigComoObjeto_(ss);
    const { pastaTermos } = obterPastasReclamacoesNoDrive_(cfg);

    const hojeFormatada = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', 'dd/MM/yyyy');
    const dataResolucaoFinal = d.DATA_RESOLUCAO || hojeFormatada;

    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @page { size: A4; margin: 20mm 20mm; }
        body {
          font-family: Arial, Helvetica, sans-serif;
          color: #111827;
          font-size: 12.5px;
          line-height: 1.5;
          margin: 0;
          padding: 10px 30px;
        }
        .header {
          margin-bottom: 20px;
        }
        .logo-title {
          font-size: 16px;
          font-weight: 800;
          color: #171b68;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .logo-sub {
          font-size: 10px;
          font-weight: 700;
          color: #e11d48;
          letter-spacing: 1px;
        }
        h2 {
          text-align: center;
          font-size: 15px;
          font-weight: bold;
          margin-top: 15px;
          margin-bottom: 25px;
          color: #0f172a;
        }
        .legal-paragraph {
          text-align: justify;
          margin-bottom: 16px;
          line-height: 1.6;
        }
        .section-header {
          font-weight: bold;
          font-size: 13px;
          margin-top: 18px;
          margin-bottom: 10px;
          color: #000;
        }
        .bullet-list {
          margin: 0 0 14px 0;
          padding-left: 18px;
        }
        .bullet-list li {
          margin-bottom: 4px;
        }
        .desc-box {
          background: #f8fafc;
          border-left: 3px solid #171b68;
          padding: 8px 12px;
          margin-top: 6px;
          margin-bottom: 14px;
          font-size: 12px;
          text-transform: uppercase;
        }
        .city-date {
          margin-top: 25px;
          margin-bottom: 30px;
        }
        .signatures-table {
          width: 100%;
          margin-top: 20px;
          border-collapse: collapse;
        }
        .sig-td {
          width: 50%;
          vertical-align: bottom;
          text-align: center;
          padding: 0 15px;
        }
        .sig-img {
          max-height: 50px;
          max-width: 220px;
          display: block;
          margin: 0 auto -10px auto;
        }
        .sig-line {
          border-top: 1px solid #334155;
          margin-top: 25px;
          margin-bottom: 6px;
        }
        .sig-name {
          font-size: 11.5px;
          font-weight: bold;
          color: #1e293b;
        }
        .footer-banner {
          margin-top: 45px;
          padding-top: 15px;
          border-top: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          font-size: 10px;
          color: #64748b;
        }
        .footer-address {
          line-height: 1.4;
        }
        .footer-brand {
          font-weight: bold;
          color: #171b68;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="logo-title">Centro Fashion</div>
        <div class="logo-sub">FORTALEZA</div>
      </div>

      <h2>Termo de Encerramento de Chamado de Reclamação</h2>

      <p class="legal-paragraph">
        <strong>CENTRO FASHION EMPREENDIMENTOS LTDA</strong>, pessoa jurídica de direito privado, inscrita no CNPJ sob o nº. 21.529.503/0001-59, com sede à Avenida Tenente Lisboa, nº. 737, Bairro Jacarecanga, Fortaleza/CE, CEP: 60.010-340., venho formalizar o encerramento do processo de resolução da reclamação registrada sob o número <strong>${d.CODIGO_PIPE}</strong>.
      </p>

      <p class="legal-paragraph">
        Considerando que a reclamação foi devidamente recebida, analisada e após diligente investigação, informamos que todas as questões apresentadas por <strong>${d.NOME_RECLAMANTE}</strong>, inscrito no CPF/MF sob o nº <strong>${d.CPF_RECLAMANTE || 'NÃO INFORMADO'}</strong>, foram devidamente resolvidas conforme descrito abaixo.
      </p>

      <div class="section-header">Resumo do Chamado de Reclamação:</div>
      <ul class="bullet-list">
        <li><strong>Data da Ocorrência:</strong> ${d.DATA_OCORRENCIA || '-'}</li>
        <li><strong>Nome do Solicitante:</strong> ${d.NOME_RECLAMANTE || '-'}</li>
        <li><strong>Email do Solicitante:</strong> ${d.EMAIL_RECLAMANTE || '-'}</li>
        <li><strong>Código Pipe:</strong> ${d.CODIGO_PIPE}</li>
        <li><strong>Localização da Loja:</strong> SETOR ${d.SETOR || '-'}, ${d.RUA || '-'}, ${d.BOX_LOJA || '-'}</li>
        <li><strong>Número de Contato:</strong> ${d.TELEFONE_RECLAMANTE || '-'}</li>
        <li><strong>Categoria da Reclamação:</strong> ${d.CATEGORIA || '-'}</li>
      </ul>

      <div class="desc-box">
        ${d.DESCRICAO_RECLAMACAO || 'DESCRIÇÃO REGISTRADA NO CHAMADO'}
      </div>

      <div class="section-header">Resumo da Resolução:</div>
      <ul class="bullet-list">
        <li><strong>Data da Resolução:</strong> ${dataResolucaoFinal}</li>
        <li><strong>Medidas tomadas para resolução do chamado:</strong> ${d.MEDIDAS_TOMADAS || 'Diligência técnica realizada e resolvida pela equipe operacional do Centro Fashion.'}</li>
      </ul>

      <p class="legal-paragraph">
        Declaramos que todas as partes envolvidas concordaram com os termos de resolução apresentados e que não existem pendências remanescentes relacionadas à reclamação mencionada.
      </p>

      <p class="legal-paragraph">
        Reconhecemos e agradecemos a cooperação do Sr(a) <strong>${d.NOME_RECLAMANTE}</strong>, durante todo o processo de resolução. Acreditamos firmemente na importância da transparência e da resolução eficaz de reclamações para manter a confiança de nossos clientes.
      </p>

      <p class="legal-paragraph">
        Este documento serve como confirmação formal do encerramento do processo de resolução da reclamação em questão. Qualquer questão adicional relacionada a este assunto deve ser direcionada a Central de Operações do Centro Fashion Fortaleza.
      </p>

      <div class="city-date">
        Fortaleza/CE, ${hojeFormatada}
      </div>

      <table class="signatures-table">
        <tr>
          <td class="sig-td">
            ${assinaturaRepresentanteBase64 ? `<img src="${assinaturaRepresentanteBase64}" class="sig-img" alt="Assinatura CFF" />` : ''}
            <div class="sig-line"></div>
            <div class="sig-name">Centro Fashion Empreendimentos Ltda</div>
          </td>
          <td class="sig-td">
            ${assinaturaSolicitanteBase64 ? `<img src="${assinaturaSolicitanteBase64}" class="sig-img" alt="Assinatura Solicitante" />` : ''}
            <div class="sig-line"></div>
            <div class="sig-name">${d.NOME_RECLAMANTE}</div>
          </td>
        </tr>
      </table>

      <div class="footer-banner">
        <div class="footer-address">
          Av. Filomeno Gomes, 430, Fortaleza, CE, Brasil<br>
          +55 85 3241-4100
        </div>
        <div class="footer-brand">
          centrofashion.com.br
        </div>
      </div>
    </body>
    </html>
    `;

    const blobHtml = Utilities.newBlob(html, 'text/html', `Termo_Reclamacao_${codigoPipe}.html`);
    const pdfBlob = blobHtml.getAs('application/pdf').setName(`Termo_Reclamacao_${codigoPipe}.pdf`);
    const arquivoPdf = pastaTermos.createFile(pdfBlob);
    arquivoPdf.setDescription(`Termo Oficial de Encerramento da Reclamação Pipe ${codigoPipe} - Centro Fashion Fortaleza`);

    // Atualiza a planilha
    const aba = ss.getSheetByName(RECLAMACOES_TABELA_NOME);
    const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(c => String(c || '').trim());
    const row = res.rowIndex;
    const agora = Utilities.formatDate(new Date(), APP.TIMEZONE || 'America/Fortaleza', "yyyy-MM-dd'T'HH:mm:ssXXX");

    const setVal = (col, val) => {
      const idx = cab.indexOf(col);
      if (idx >= 0) aba.getRange(row, idx + 1).setValue(val);
    };

    setVal('URL_TERMO_ENCERRAMENTO', arquivoPdf.getUrl());
    setVal('ID_TERMO_DRIVE', arquivoPdf.getId());
    setVal('STATUS', 'ENCERRADO');
    setVal('DATA_RESOLUCAO', dataResolucaoFinal);
    setVal('ATUALIZADO_EM', agora);
    if (assinaturaSolicitanteBase64) setVal('ASSINATURA_SOLICITANTE_DATA', 'COLETADA_DIGITALMENTE');
    if (assinaturaRepresentanteBase64) setVal('ASSINATURA_REPRESENTANTE_DATA', 'COLETADA_DIGITALMENTE');

    return {
      ok: true,
      codigoPipe: codigoPipe,
      urlPdf: arquivoPdf.getUrl(),
      downloadUrl: arquivoPdf.getDownloadUrl(),
      driveId: arquivoPdf.getId(),
      mensagem: `Termo de Encerramento do Chamado ${codigoPipe} gerado com sucesso em PDF.`
    };
  } catch (err) {
    console.error('[RECLAMACOES] Erro ao gerar Termo de Encerramento:', err);
    return { ok: false, error: err.message || String(err) };
  }
}
