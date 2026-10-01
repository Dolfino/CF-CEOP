/** Serviços locais de mapa — S2. Não consulta a planilha antiga. */
function setupS2(){
  const ss=SpreadsheetApp.getActive(); const cfg=ss.getSheetByName('CONFIG');
  if(!cfg) throw new Error('CONFIG ausente.');
  const atual=lerConfigComoObjeto_(ss);
  if(atual.S1_STATUS!=='VALIDADO' && atual.S1_STATUS!=='MIGRADO') throw new Error('S1 ainda não está validada.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE',APP.FASE,'Fase de implementação validada');
  setConfigValue_(cfg,'S2_STATUS','INSTALADO','Mapa e navegação instalados');
  setConfigValue_(cfg,'MAPA_PADRAO',atual.MAPA_PADRAO || primeiroMapaAtivoId_(),'Mapa carregado inicialmente');
  setConfigValue_(cfg,'S2_INSTALADO_EM',Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),'Data/hora de instalação S2');
  atualizarReadmeS2(); SpreadsheetApp.flush();
  return diagnosticoS2();
}

function appCarregarS2(){
  const ss=SpreadsheetApp.getActive();
  const cfg=lerConfigComoObjeto_(ss);
  // S20.1: o bootstrap normal não executa diagnosticoS2().
  // O diagnóstico completo continua disponível manualmente em diagnosticoS2().
  return {
    app:{id:APP.ID,nome:APP.NOME,versao:APP.VERSAO,fase:APP.FASE,modoDados:APP.MODO_DADOS},
    mapaPadrao:cfg.MAPA_PADRAO||'',
    mapas:listarMapasS2_(),
    diagnostico:{ok:true,modo:'LAZY',fase:APP.FASE}
  };
}

function listarMapasS2_(){
  return linhasObjetosS2_('MAPAS_SETORES').filter(r=>ativoS2_(r.ATIVO)).map(r=>({
    id:String(r.ID_MAPA_SETOR||''), idPlanta:String(r.ID_PLANTA||''), idSetor:String(r.ID_SETOR||''),
    nome:String(r.NOME||r.MAPA||r.ID_MAPA_SETOR||''), piso:String(r.PISO??''), cor:String(r.COR_SETOR||''),
    imagemId:String(r.IMAGEM_ID||''), imagemUrl:String(r.IMAGEM_URL||''),
    largura:Number(r.LARGURA_PX||r.LARGURA_ORIGINAL||0), altura:Number(r.ALTURA_PX||r.ALTURA_ORIGINAL||0),
    rotacao:Number(r.ROTACAO_GRAUS||0)
  })).filter(x=>x.id&&x.imagemId);
}

function appObterImagemMapaS2(idMapa){
  const mapa=listarMapasS2_().find(x=>x.id===String(idMapa||'')); if(!mapa) throw new Error('Mapa não encontrado.');
  const blob=DriveApp.getFileById(mapa.imagemId).getBlob();
  return {id:mapa.id,mimeType:blob.getContentType(),dataUrl:`data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`};
}

function appObterCamadasS2(idMapa){
  limparCacheLinhasObjetosS2_();
  const id=String(idMapa||'');
  return {
    referencias:pontosDaAbaS2_('PONTOS_REFERENCIA',id,['X_NORMALIZADO','X'],['Y_NORMALIZADO','Y'],['NOME','NOME_REFERENCIA','DESCRICAO'],350),
    cruzamentos:pontosDaAbaS2_('CRUZAMENTOS',id,['X_NORMALIZADO','X'],['Y_NORMALIZADO','Y'],['NOME','NOME_CRUZAMENTO','CODIGO'],350),
    lojas:pontosDaAbaS2_('LOJAS_MAPA',id,['X_NORMALIZADO','CENTRO_X','X_CENTRO','X'],['Y_NORMALIZADO','CENTRO_Y','Y_CENTRO','Y'],['NUMERO_LOJA','LUC','NOME_LOJA'],5000),
    localizacao: typeof appObterDadosLocalizacaoOfflineS6 === 'function' ? appObterDadosLocalizacaoOfflineS6(id) : null
  };
}

function extrairSiglaRuaS3_(nomeOuIdCorredor) {
  if (!nomeOuIdCorredor) return '';
  const s = String(nomeOuIdCorredor).trim();

  const MAPA_SIGLAS = {
    // SETOR AZUL / GERAL
    'DOM MANUEL': 'ADMMN', 'ADMMN': 'ADMMN',
    '25 DE MARÇO': 'R2DMR', '25 DE MARCO': 'R2DMR', 'R2DMR': 'R2DMR',
    'CORONEL FERRAZ': 'RCRFR', 'RCRFR': 'RCRFR',
    'SÃO JOSÉ': 'RSJS', 'SAO JOSE': 'RSJS', 'RSJS': 'RSJS',
    'GOVERNADOR SAMPAIO': 'RGVSM', 'RGVSM': 'RGVSM',
    'CONDE D\'EU': 'RCNDD', 'CONDE DEU': 'RCNDD', 'RCNDD': 'RCNDD',
    'GENERAL BEZERRIL': 'RGNBZ', 'RGNBZ': 'RGNBZ',
    'FLORIANO PEIXOTO': 'RFLPX', 'RFLPX': 'RFLPX',
    'MAJOR FACUNDO': 'RMJFC', 'RMJFC': 'RMJFC',
    'BARÃO DO RIO BRANCO': 'RBDRB', 'BARAO DO RIO BRANCO': 'RBDRB', 'RBDRB': 'RBDRB', 'RBRDB': 'RBDRB',
    'SENADOR POMPEU': 'RSNPM', 'RSNPM': 'RSNPM',
    'GENERAL SAMPAIO': 'RGNSM', 'RGNSM': 'RGNSM',
    '24 DE MAIO': 'R24DM', 'R24DM': 'R24DM',
    'TRISTÃO GONÇALVES': 'RTRGN', 'TRISTAO GONCALVES': 'RTRGN', 'RTRGN': 'RTRGN',
    'IMPERADOR': 'AVIMP', 'AVENIDA DO IMPERADOR': 'AVIMP', 'DO IMPERADOR': 'AVIMP', 'AVIMP': 'AVIMP',
    'PRINCESA ISABEL': 'RPRIS', 'RPRIS': 'RPRIS',
    'ALBERTO NEPOMUCENO': 'AVALN', 'NEPOMUCENO': 'AVALN', 'ALENCASTRO': 'AVALN', 'AVALN': 'AVALN',
    'CARAPINIMA': 'AVCRP', 'CELSO ROSSI': 'AVCRP', 'AVCRP': 'AVCRP',
    'JOSE AVELINO': 'RJSA', 'JOSÉ AVELINO': 'RJSA', 'RJSA': 'RJSA',

    // SETOR VERDE & BRANCO & ROXO
    'ANTONIO POMPEU': 'RANPM', 'ANTÔNIO POMPEU': 'RANPM', 'RANPM': 'RANPM',
    'METON DE ALENCAR': 'RMDAL', 'METON': 'RMDAL', 'NETON DE ALENCAR': 'RMDAL', 'NEWTON DE ALENCAR': 'RMDAL', 'RNDAL': 'RMDAL', 'RMDAL': 'RMDAL',
    'CLARINDO DE QUEIROZ': 'RCDQR', 'RCDQR': 'RCDQR',
    'DUQUE DE CAXIAS': 'ADDCX', 'ADQCX': 'ADDCX', 'ADDCX': 'ADDCX',
    'DOM PEDRO': 'RDPDI', 'DOM PEDRO I': 'RDPDI', 'RDPD1': 'RDPDI', 'RDPDI': 'RDPDI',
    'PEDRO PEREIRA': 'RPDPR', 'RPDPR': 'RPDPR',
    'LIBERATO BARROSO': 'RLBBR', 'RLBBR': 'RLBBR',
    'GUILHERME ROCHA': 'RGLRC', 'RGLRC': 'RGLRC', 'RGLC': 'RGLRC',
    'SÃO PAULO': 'RSPL', 'SAO PAULO': 'RSPL', 'SALDANHA MARINHO': 'RSPL', 'RSDMR': 'RSPL', 'RSPL': 'RSPL',
    'SENADOR ALENCAR': 'RSNAL', 'RSNAL': 'RSNAL',
    'CASTRO E SILVA': 'RCESL', 'CASTRO SILVA': 'RCESL', 'SILVA JARDIM': 'RCESL', 'GENERAL SILVA JARDIM': 'RCESL', 'RGSJR': 'RCESL', 'RCESL': 'RCESL', 'RCESJ': 'RCESL',
    'DOUTOR JOÃO MOREIRA': 'RDRJM', 'DR. JOÃO MOREIRA': 'RDRJM', 'JOÃO MOREIRA': 'RDRJM', 'JOAO MOREIRA': 'RDRJM', 'RDJMR': 'RDRJM', 'RDRJM': 'RDRJM',
    'SENADOR JAGUARIBE': 'RSNJG', 'JAGUARIBE': 'RSNJG', 'RSNGB': 'RSNJG', 'RSNJG': 'RSNJG',
    'CASTELO BRANCO': 'APCBR', 'PRES. CASTELO BRANCO': 'APCBR', 'APCRB': 'APCBR', 'APCBR': 'APCBR',

    // BRANCO / OUTROS
    'DOMINGOS OLÍMPIO': 'AVDMO', 'DOMINGOS OLIMPIO': 'AVDMO', 'MONTE ALVERNE': 'AVDMO', 'AVDMO': 'AVDMO',
    'MONSENHOR TABOSA': 'AVMTB', 'AVMTB': 'AVMTB',
    'BATURITÉ': 'TVBTR', 'BATURITE': 'TVBTR', 'BENTOS': 'TVBTR', 'TRAVESSA BENTOS': 'TVBTR', 'TVBTR': 'TVBTR',
    'PARÁ': 'TVPR', 'PARA': 'TVPR', 'PARAIBA': 'TVPR', 'TRAVESSA PARAIBA': 'TVPR', 'TVPR': 'TVPR',
    '13 DE MAIO': 'A13DM', 'A13DM': 'A13DM',

    // ROXO / VERMELHO / SERVIÇOS
    'ESTAC. VIS': 'ESTVT', 'ESTVT': 'ESTVT',
    'COSMETICOS': 'COSM', 'COSM': 'COSM',
    'SALÃO': 'SLBLZ', 'SALAO': 'SLBLZ', 'SLBLZ': 'SLBLZ',
    'MONTE': 'AVMNT', 'AVMNT': 'AVMNT',
    'POSTO POLICIAL': 'PSTPL', 'PSTPL': 'PSTPL',
    'CAIXA ELETRONICO': 'CXLNC', 'CXLNC': 'CXLNC',
    'LOJA DESIGN': 'LJDTS', 'LJDTS': 'LJDTS',
    'PRACA': 'PR', 'PRAÇA': 'PR', 'PR': 'PR',
    'PRACA ALIMENTACAO': 'PRALM', 'PRAÇA ALIMENTAÇÃO': 'PRALM', 'PRALM': 'PRALM',
    'PRACA DE ALIMENTACAO': 'PRALM', 'PRAÇA DE ALIMENTAÇÃO': 'PRALM',
    'PRARE': 'PRARE', 'PRACMM': 'PRACMM',
    'DOCAV': 'DOCAV', 'DOCAH': 'DOCAH', 'DOCAM': 'DOCAM', 'DOCAF': 'DOCAF', 'DOCAC': 'DOCAC',
    'BRMALL': 'BRMALL', 'RXMALL': 'RXMALL', 'CORS': 'CORS'
  };

  const m = s.match(/\(([A-Z0-9]+)\)/i);
  if (m) {
    const raw = m[1].toUpperCase();
    if (MAPA_SIGLAS[raw]) return MAPA_SIGLAS[raw];
    if (/^[A-Z][A-Z0-9]{1,5}$/.test(raw)) return raw;
  }

  const upper = s.toUpperCase();
  for (const [k, v] of Object.entries(MAPA_SIGLAS)) {
    if (upper.includes(k)) return v;
  }
  return '';
}

var _linhasObjetosCacheS2 = {};
var _corredoresPolylinesCache_ = {};

function limparCacheLinhasObjetosS2_() {
  _linhasObjetosCacheS2 = {};
  _corredoresPolylinesCache_ = {};
}

function pontosDaAbaS2_(aba,idMapa,xKeys,yKeys,labelKeys,limite){
  const max = (typeof limite === 'number' && limite > 0) ? limite : 5000;
  const rawRows = linhasObjetosS2_(aba);
  const rows = rawRows.filter(r=>String(r.ID_MAPA_SETOR||'')===idMapa && ativoS2_(r.ATIVO)).slice(0,max);

  let pontos = rows.map(r=>{
    const x=primeiroNumeroS2_(r,xKeys), y=primeiroNumeroS2_(r,yKeys); let label=''; for(const k of labelKeys){if(r[k]!==''&&r[k]!=null){label=String(r[k]);break;}}
    let luc = String(r.LUC || r.CODIGO_LUC || r.ID_ESPACO || '').trim();
    const numeroLoja = String(r.NUMERO_LOJA || r.NUMERO || '').trim();
    const nomeLoja = String(r.NOME_LOJA || r.NOME || '').trim();
    let corredor = String(r.ID_CORREDOR || r.CORREDOR || r.RUA || '').trim();
    let siglaRua = '';

    if (aba === 'LOJAS_MAPA') {
      const ehPerimetral = /AVALN|AVCRP|RJSA/i.test(corredor);
      const ehPontaGaleria = ehPerimetral && (
        numeroLoja === '1106' ||
        numeroLoja === '1101' ||
        numeroLoja.endsWith('06') ||
        numeroLoja.endsWith('02') ||
        numeroLoja.endsWith('08') ||
        numeroLoja.endsWith('12') ||
        numeroLoja.endsWith('10') ||
        numeroLoja.endsWith('18') ||
        numeroLoja.endsWith('20') ||
        numeroLoja.endsWith('24') ||
        numeroLoja.endsWith('29') ||
        numeroLoja.endsWith('33') ||
        numeroLoja.endsWith('37') ||
        numeroLoja.endsWith('41') ||
        numeroLoja.endsWith('43')
      );

      if ((!corredor || ehPontaGaleria) && typeof localizarCorredorVerticalS3_ === 'function' && Number.isFinite(x) && Number.isFinite(y)) {
        try {
          const cv = localizarCorredorVerticalS3_(idMapa, x, y);
          if (cv) corredor = cv.nome || cv.idCorredor || corredor;
        } catch (_) {}
      } else if (!corredor && typeof localizarCorredorS3_ === 'function' && Number.isFinite(x) && Number.isFinite(y)) {
        try {
          const c = localizarCorredorS3_(idMapa, x, y);
          if (c) corredor = c.nome || c.idCorredor || '';
        } catch (_) {}
      }
      siglaRua = extrairSiglaRuaS3_(corredor);
      if ((!luc || luc === numeroLoja || ehPontaGaleria) && siglaRua && numeroLoja) {
        luc = siglaRua + numeroLoja;
      }
    }

    return {
      x,
      y,
      label: label || numeroLoja || luc,
      luc,
      numeroLoja,
      nomeLoja,
      corredor,
      rua: corredor,
      siglaRua
    };
  }).filter(p=>Number.isFinite(p.x)&&Number.isFinite(p.y)&&p.x>=0&&p.x<=1&&p.y>=0&&p.y<=1);

  // S26.10-CEOP: Injeção resiliente dos 16 boxes da Ilha Central do Setor Roxo
  if (aba === 'LOJAS_MAPA' && idMapa === 'MAP-CFF-N3-ROXO') {
    const temIlhaCentral = pontos.some(p => p.luc === 'R24DM3241' || (p.numeroLoja === '3241' && p.siglaRua === 'R24DM'));
    if (!temIlhaCentral && typeof obterDefinicaoBoxesIlhaCentralRoxo_ === 'function') {
      const defs = obterDefinicaoBoxesIlhaCentralRoxo_();
      defs.forEach(d => {
        pontos.push({
          x: d.x,
          y: d.y,
          label: d.numero,
          luc: d.luc,
          numeroLoja: d.numero,
          nomeLoja: d.nome,
          corredor: d.corredor,
          rua: d.corredor,
          siglaRua: extrairSiglaRuaS3_(d.corredor)
        });
      });
      try {
        if (typeof inserirBoxesIlhaCentralSetorRoxo === 'function') {
          inserirBoxesIlhaCentralSetorRoxo();
        }
      } catch (_) {}
    }
  }

  return pontos;
}
function primeiroNumeroS2_(r,keys){for(const k of keys){const n=Number(r[k]);if(Number.isFinite(n))return n;}return NaN;}
function ativoS2_(v){return !['NAO','NÃO','FALSE','0','INATIVO'].includes(String(v??'SIM').trim().toUpperCase());}
function linhasObjetosS2_(aba, forcarLeitura){
  if (!_linhasObjetosCacheS2) _linhasObjetosCacheS2 = {};
  if (!forcarLeitura && _linhasObjetosCacheS2[aba]) {
    return _linhasObjetosCacheS2[aba];
  }
  const sh=SpreadsheetApp.getActive().getSheetByName(aba);
  if(!sh||sh.getLastRow()<2){
    _linhasObjetosCacheS2[aba] = [];
    return [];
  }
  const vals=sh.getDataRange().getValues(),h=vals[0].map(x=>String(x).trim());
  const res = vals.slice(1).map(row=>Object.fromEntries(h.map((k,i)=>[k,row[i]])));
  _linhasObjetosCacheS2[aba] = res;
  return res;
}
function primeiroMapaAtivoId_(){return listarMapasS2_()[0]?.id||'';}

function diagnosticoS2(){
  const ss=SpreadsheetApp.getActive(), cfg=lerConfigComoObjeto_(ss), checks=[];
  check_(checks,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID);
  check_(checks,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO);
  check_(checks,'APP_FASE',cfg.APP_FASE===APP.FASE,cfg.APP_FASE);
  check_(checks,'MODO_DADOS',cfg.MODO_DADOS===APP.MODO_DADOS,cfg.MODO_DADOS);
  check_(checks,'S1_VALIDADA',['VALIDADO','MIGRADO'].includes(cfg.S1_STATUS),cfg.S1_STATUS);
  const mapas=listarMapasS2_(); check_(checks,'MAPAS_ATIVOS',mapas.length>0,`${mapas.length} mapa(s)`);
  check_(checks,'MAPA_PADRAO',!!cfg.MAPA_PADRAO && mapas.some(m=>m.id===cfg.MAPA_PADRAO),cfg.MAPA_PADRAO||'ausente');
  mapas.forEach(m=>{try{const f=DriveApp.getFileById(m.imagemId);check_(checks,`IMAGEM_${m.id}`,true,f.getName());}catch(e){check_(checks,`IMAGEM_${m.id}`,false,e.message);}});
  ['PONTOS_REFERENCIA','CRUZAMENTOS','LOJAS_MAPA'].forEach(a=>check_(checks,`CAMADA_${a}`,!!ss.getSheetByName(a),ss.getSheetByName(a)?'OK':'ausente'));
  const falhas=checks.filter(c=>!c.ok); return {ok:!falhas.length,totalChecks:checks.length,totalFalhas:falhas.length,checks};
}
function mostrarDiagnosticoS2(){const d=diagnosticoS2();SpreadsheetApp.getUi().alert(`Diagnóstico S2`,`${d.ok?'MAPA S2 OK':'HÁ PENDÊNCIAS'}\n\n${d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n')}`,SpreadsheetApp.getUi().ButtonSet.OK);return d;}

function atualizarReadmeS2(){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName('README'); if(!sh)sh=ss.insertSheet('README'); sh.clear();
  const rows=[['SINALIZAÇÃO DO MALL — FASE S2'],[''],['VERSÃO'],[APP.VERSAO],[''],['OBJETIVO'],['Exibir a cartografia local em mapa navegável, mantendo LOCAL_INDEPENDENTE.'],[''],['ENTREGAS'],['Seleção de mapa/setor; imagem da planta; pan; zoom por roda; pinça; botões +/−; centralizar; camadas de referências, cruzamentos e lojas.'],[''],['GATE S2'],['1. Executar setupS2().'],['2. Publicar nova versão do Web App.'],['3. Abrir no desktop e no celular.'],['4. Confirmar troca de mapa, pan, zoom, pinça e Centralizar.'],['5. Confirmar que as camadas podem ser ligadas/desligadas.'],['6. Executar diagnosticoS2() sem falhas.'],[''],['NÃO FAZ PARTE DE S2'],['Cadastro de sinalização, Outbox e fotos offline.'],[''],['PRÓXIMA FASE'],['S3 — Cadastro de sinalização.']];
  sh.getRange(1,1,rows.length,1).setValues(rows);sh.getRange('A1').setFontWeight('bold').setFontSize(16).setFontColor('#171B68');sh.setColumnWidth(1,900);sh.getRange('A:A').setWrap(true);
}

/** ========================= S3 — CADASTRO ========================= */
function setupS3(){
  const ss=SpreadsheetApp.getActive(); const cfg=ss.getSheetByName('CONFIG');
  if(!cfg) throw new Error('CONFIG ausente.');
  const atual=lerConfigComoObjeto_(ss);
  if(atual.S2_STATUS!=='INSTALADO' && atual.APP_FASE!=='S2' && atual.APP_FASE!=='S3') throw new Error('S2 ainda não está instalada.');
  garantirCabecalhoS3_('REGISTROS', CABECALHOS_REGISTROS_S3_());
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE',APP.FASE,'Fase de implementação validada');
  setConfigValue_(cfg,'S3_STATUS','INSTALADO','Cadastro online de sinalização instalado');
  setConfigValue_(cfg,'S3_INSTALADO_EM',Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),'Data/hora de instalação S3');
  atualizarReadmeS3(); SpreadsheetApp.flush();
  return diagnosticoS3();
}

function appCarregarS3(){
  const base=appCarregarS2();
  return Object.assign({},base,{app:{id:APP.ID,nome:APP.NOME,versao:APP.VERSAO,fase:APP.FASE,modoDados:APP.MODO_DADOS},tipos:TIPOS_S3_,finalidades:FINALIDADES_S3_,materiais:MATERIAIS_S3_,estados:ESTADOS_S3_,condicoes:CONDICOES_S3_,responsaveis:RESPONSAVEIS_S3_});
}

const TIPOS_S3_=[
  'Alvenaria','Hidráulica','Pintura','Serralheria','Marcenaria','Cobertura / Telhado',
  'Posto de Coleta','Sanitário','Depósito de Resíduos','Jardim Interno','Jardim Externo',
  'Caixa Acústica','Projetor','Painel de LED','Rack / Switch AV','Microfone / Receptor',
  'Fancoil / Chiller','Ar Condicionado Split','Quadro Elétrico (QGBT)','Painel Elétrico','Luminária de Emergência','Sensor / Termostato',
  'Doca de Carga','Cancela','Porta Automática','Escada Rolante','Elevador de Carga / Serviço',
  'Placa direcional','Placa informativa','Placa de segurança','Totem','Adesivo de piso','Adesivo de parede','Outro'
];
const FINALIDADES_S3_=[
  'Manutenção Predial','Limpeza e Jardinagem','Som e Projeção','Iluminação e Climatização',
  'Operação Geral','Segurança Operacional','Atendimento Corretivo','Inspeção Preventiva',
  'Comunicação Visual','Outra'
];
const MATERIAIS_S3_=[
  'Alvenaria / Concreto','Metálico / Alumínio','PVC / Hidráulico','Elétrico / Eletrônico',
  'Vidro / Acrílico','Madeira / Marcenaria','Gesso / Drywall','Termoplástico / Borracha',
  'Adesivo / Vinil','Outro'
];
const ESTADOS_S3_=['Operacional / Normal','Necessita Reparo','Paralisado / Crítico','Em Manutenção','Novo / Impecável'];
const CONDICOES_S3_=[
  'Sem problema / Em funcionamento','Vazamento / Infiltração','Desgaste / Quebrado','Queimado / Sem energia',
  'Ruído anormal / Vibração','Obstruído / Sujo','Superaquecimento','Falha de comunicação / Sinal',
  'Ausente / Desativado','Outro'
];
const RESPONSAVEIS_S3_=[
  'CEOP — Manutenção Predial','CEOP — Limpeza e Jardinagem','CEOP — Som e Projeção',
  'CEOP — Iluminação e Clima','CEOP — Operações Gerais','Segurança Patrimonial',
  'TI / Automação','Lojista','Terceirizado','Outro'
];

function appPrepararLocalizacaoS3(dados){
  const d=dados||{}, idMapa=String(d.idMapaSetor||'').trim(), x=numS3_(d.x), y=numS3_(d.y);
  if(!idMapa) throw new Error('Mapa não informado.');
  if(!Number.isFinite(x)||!Number.isFinite(y)||x<0||x>1||y<0||y>1) throw new Error('Coordenadas inválidas.');
  const mapa=listarMapasS2_().find(m=>m.id===idMapa); if(!mapa) throw new Error('Mapa não encontrado.');
  const corr=localizarCorredorS3_(idMapa,x,y);
  const cruz=maisProximoS3_('CRUZAMENTOS',idMapa,x,y,['X_NORMALIZADO','X'],['Y_NORMALIZADO','Y'],['NOME','NOME_CRUZAMENTO','CODIGO'],0.05);
  const ref=maisProximoS3_('PONTOS_REFERENCIA',idMapa,x,y,['X_NORMALIZADO','X'],['Y_NORMALIZADO','Y'],['NOME','NOME_REFERENCIA','DESCRICAO'],0.06);
  const loja=maisProximoLojaS3_(idMapa,x,y,0.045);
  const seg=corr?localizarSegmentoS3_(corr.idCorredor,corr.percentual):null;
  return {
    idMapaSetor:idMapa,mapa:mapa.nome,piso:mapa.piso,x,y,
    idCorredor:corr?.idCorredor||'',rua:corr?.nome||'',distanciaCorredor:corr?.distancia??null,
    idSegmento:seg?.ID_SEGMENTO||'',trecho:seg?.NOME_SEGMENTO||seg?.DESCRICAO_PADRAO||'',
    cruzamento:cruz?.label||'',referencia:ref?.label||'',
    numeroLoja:loja?.NUMERO_LOJA||'',luc:loja?.LUC||'',nomeLoja:loja?.NOME_LOJA||'',
    resumo:montarResumoLocalizacaoS3_(mapa,corr,seg,cruz,ref,loja)
  };
}

function appCriarRegistroS3(dados){
  exigirPermissaoS14_('criarRegistro');
  const d=dados||{};
  validarRegistroS3_(d);
  const ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('REGISTROS'); if(!sh) throw new Error('Aba REGISTROS ausente.');
  garantirColunasCeopAs_(sh);
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const clientEventId=String(d.clientEventId||Utilities.getUuid()).trim();
    const existente=obterRegistroPorClientEventS3_(clientEventId);
    if(existente) {
      // R08 (auditoria S26.10) — log de idempotência (sem dados do registro).
      console.log(JSON.stringify({evento:'REGISTRO_IDEMPOTENTE',clientEventId,ts:new Date().toISOString()}));
      return {
      ok: true,
      idempotente: true,
      idRegistro: String(existente.ID_REGISTRO || ''),
      protocolo: String(existente.PROTOCOLO || ''),
      status: String(existente.STATUS || 'ATIVA'),
      clientEventId: String(clientEventId)
    };
    }
    // R08 (auditoria S26.10) — log estruturado de criação de registro.
    console.log(JSON.stringify({evento:'REGISTRO_CRIAR',clientEventId,ts:new Date().toISOString()}));
    const id='SIG-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
    const protocolo=gerarProtocoloS3_();
    const agora=new Date();
    const usuario=(usuarioRpcAtualS223_()?.email||Session.getActiveUser().getEmail())||'WEB_APP';
    const torreS2610G=(typeof s2610GNormalizarLocalizacaoRegistro_==='function')
      ? s2610GNormalizarLocalizacaoRegistro_(d)
      : {};
    const ehAS = String(d.categoriaRegistro || '').toUpperCase() === 'AUTORIZACAO_SERVICO';
    const categoria = ehAS ? 'AUTORIZACAO_SERVICO' : 'ATIVO';

    const codPatrimonio = String(d.codigoPatrimonio || d.patrimonioTag || d.asBox || '').trim();
    let txtSinalizacao = String(d.textoSinalizacao || '').trim();
    if (!ehAS && codPatrimonio && !txtSinalizacao.includes('[TAG: ' + codPatrimonio + ']')) {
      txtSinalizacao = ('[TAG: ' + codPatrimonio + '] ' + txtSinalizacao).trim();
    }
    const obj={
      ID_REGISTRO:id,CLIENT_EVENT_ID:clientEventId,PROTOCOLO:protocolo,CRIADO_EM:agora,SINCRONIZADO_EM:agora,
      STATUS:ehAS ? 'AUTORIZADO' : 'ATIVA',
      CATEGORIA_REGISTRO:categoria,
      TIPO:d.tipo,FINALIDADE:d.finalidade,TITULO:d.titulo,TEXTO_SINALIZACAO:txtSinalizacao,
      CODIGO_PATRIMONIO:codPatrimonio,DESCRICAO:d.descricao,
      MATERIAL:d.material||'',DIMENSOES:d.dimensoes||'',COR:d.cor||'',FIXACAO:d.fixacao||'',ILUMINADA:simNaoS3_(d.iluminada),DUPLA_FACE:simNaoS3_(d.duplaFace),QR_CODE:simNaoS3_(d.qrCode),BRAILLE:simNaoS3_(d.braille),PICTOGRAMA:simNaoS3_(d.pictograma),
      ESTADO_CONSERVACAO:d.estadoConservacao,CONDICAO:d.condicao,RESPONSAVEL:d.responsavel,
      ID_MAPA_SETOR:d.idMapaSetor,MAPA:d.mapa,PISO:d.piso,X_NORMALIZADO:Number(d.x),Y_NORMALIZADO:Number(d.y),ID_CORREDOR:d.idCorredor||'',RUA:d.rua||'',ID_SEGMENTO:d.idSegmento||'',TRECHO:d.trecho||'',CRUZAMENTO:d.cruzamento||'',REFERENCIA:d.referencia||'',NUMERO_LOJA:d.numeroLoja||d.asBox||'',LUC:d.luc||'',NOME_LOJA:d.nomeLoja||d.asNomeTitular||'',LOCALIZACAO_CONFIRMADA:'SIM',
      ID_PLANTA_NIVEL:torreS2610G.idPlantaNivel||'',
      X_NIVEL:Number.isFinite(torreS2610G.xNivel)?torreS2610G.xNivel:'',
      Y_NIVEL:Number.isFinite(torreS2610G.yNivel)?torreS2610G.yNivel:'',
      ID_TORRE:torreS2610G.idTorre||'',CODIGO_TORRE:torreS2610G.codigoTorre||'',NOME_TORRE:torreS2610G.nomeTorre||'',
      ID_REPRESENTACAO_TORRE:torreS2610G.idRepresentacaoTorre||'',
      VERSAO_GEOMETRIA_TORRE:Number.isFinite(torreS2610G.versaoGeometriaTorre)?torreS2610G.versaoGeometriaTorre:'',
      ORIGEM_TORRE:torreS2610G.origemTorre||'',
      DATA_INSTALACAO:d.dataInstalacao||d.asDataInicio||'',VALIDADE:d.validade||d.asDataFim||'',DATA_ULTIMA_INSPECAO:'',PROXIMA_INSPECAO:'',ORIGEM:'WEB_APP',USUARIO:usuario,DEVICE_ID:String(d.deviceId||''),VERSAO_APP:APP.VERSAO,
      SOLICITANTE_NOME:d.asNomeTitular||'',
      SOLICITANTE_CPF:d.asCpfTitular||'',
      SOLICITANTE_CONTATO:d.asContatoTitular||'',
      SOLICITANTE_EMAIL:d.asEmailSolicitante||d.asEmailTitular||'',
      TIPO_SOLICITACAO:d.asTipoSolicitacao||'',
      EMPRESA_INTERNET:d.asEmpresaInternet||'',
      DATA_INICIO_SERVICO:d.asDataInicio||'',
      DATA_FIM_SERVICO:d.asDataFim||'',
      HORARIO_SERVICO:d.asHorarioServico||'',
      HORARIO_SEGUNDA:d.asHorarioSegunda||'',
      PRESTADOR_NOME:d.asNomePrestador||'',
      PRESTADOR_CPF:d.asCpfPrestador||'',
      PRESTADOR_CONTATO:d.asContatoPrestador||'',
      PRESTADOR_EMAIL:d.asEmailPrestador||'',
      PRESTADOR_EMPRESA:d.asEmpresaPrestador||d.asEmpresa||'',
      EQUIPE_AJUDANTES:d.asEquipe||'',
      ITENS_RETIRADA:d.asItensRetirada||'',
      SERVICO_ESTRUTURA:d.asServicoEstrutura||'',
      SERVICO_REVESTIMENTO:d.asServicoRevestimento||'',
      SERVICO_INSTALACOES:d.asServicoInstalacoes||'',
      TERMOS_LGPD_ACEITOS:simNaoS3_(d.asTermosLgpd)
    };
    appendObjetoS3_(sh,obj);
    return {
      ok: true,
      idRegistro: String(id),
      protocolo: String(protocolo),
      status: String(obj.STATUS),
      clientEventId: String(clientEventId),
      sincronizadoEm: Utilities.formatDate(agora, APP.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX")
    };
  } finally { lock.releaseLock(); }
}

function validarRegistroS3_(d){
  const ehAS = String(d.categoriaRegistro || '').toUpperCase() === 'AUTORIZACAO_SERVICO';
  if (ehAS) {
    const obrigAS = [
      ['asNomeTitular','Nome do Solicitante'],
      ['asCpfTitular','CPF do Solicitante'],
      ['asContatoTitular','Contato do Solicitante'],
      ['asTipoSolicitacao','Tipo de Solicitação'],
      ['asDataInicio','Data de Início'],
      ['asDataFim','Data de Finalização'],
      ['asHorarioServico','Horário Autorizado'],
      ['idMapaSetor','Mapa']
    ];
    obrigAS.forEach(([k,n])=>{if(!String(d[k]??'').trim())throw new Error(`${n} é obrigatório.`);});
    const tipo = String(d.asTipoSolicitacao || '').toUpperCase();
    if (tipo === 'RETIRADA DE PERTENCES') {
      if (!String(d.asDescricaoServico || d.asEspecificacaoItens || '').trim()) {
        throw new Error('Especificação de itens a serem retirados é obrigatória.');
      }
    } else {
      if (!String(d.asDescricaoServico || '').trim()) {
        throw new Error('Descrição do serviço é obrigatória.');
      }
      if (!String(d.asNomePrestador || '').trim()) {
        throw new Error('Nome do prestador de serviços é obrigatório.');
      }
      if (!String(d.asCpfPrestador || '').trim()) {
        throw new Error('CPF do prestador de serviços é obrigatório.');
      }
    }
    if (String(d.asTermosLgpd || '').toUpperCase() !== 'SIM' && d.asTermosLgpd !== true) {
      throw new Error('É obrigatório o aceite das normas operacionais e LGPD.');
    }
  } else {
    const obrig=[['tipo','Tipo'],['finalidade','Finalidade'],['titulo','Título'],['descricao','Descrição'],['estadoConservacao','Estado de conservação'],['condicao','Condição'],['responsavel','Responsável'],['idMapaSetor','Mapa']];
    obrig.forEach(([k,n])=>{if(!String(d[k]??'').trim())throw new Error(`${n} é obrigatório.`);});
  }
  const x=numS3_(d.x),y=numS3_(d.y); if(!Number.isFinite(x)||!Number.isFinite(y)||x<0||x>1||y<0||y>1)throw new Error('Selecione um ponto válido no mapa.');
  if(String(d.localizacaoConfirmada||'').toUpperCase()!=='SIM')throw new Error('Confirme a localização antes de salvar.');
}

function gerarProtocoloS3_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  const data=Utilities.formatDate(new Date(),APP.TIMEZONE,'yyyyMMdd'); const prefix=`SIG-${data}-`;
  let max=0; if(sh&&sh.getLastRow()>1){const vals=sh.getRange(2,3,sh.getLastRow()-1,1).getDisplayValues().flat();vals.forEach(v=>{if(String(v).startsWith(prefix)){const n=Number(String(v).slice(prefix.length));if(Number.isFinite(n))max=Math.max(max,n);}});}
  return prefix+String(max+1).padStart(4,'0');
}
function obterRegistroPorClientEventS3_(id){const rows=linhasObjetosS2_('REGISTROS');return rows.find(r=>String(r.CLIENT_EVENT_ID||'')===id)||null;}
function appendObjetoS3_(sh,obj){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);sh.appendRow(h.map(k=>Object.prototype.hasOwnProperty.call(obj,k)?obj[k]:''));}
function simNaoS3_(v){return ['SIM','TRUE','1',true,1,'on'].includes(v)?'SIM':'NAO';}

function obterCorredoresPolylinesS3_(idMapa, somenteVerticais) {
  const cacheKey = (somenteVerticais ? 'VERT_' : 'ALL_') + String(idMapa || '').trim();
  if (_corredoresPolylinesCache_[cacheKey]) {
    return _corredoresPolylinesCache_[cacheKey];
  }
  let corredores = linhasObjetosS2_('CORREDORES').filter(r => String(r.ID_MAPA_SETOR || '') === String(idMapa || '') && ativoS2_(r.ATIVO));
  if (somenteVerticais) {
    corredores = corredores.filter(r => !/AVALN|AVCRP|RJSA|A13DM|TVBTR|AVDMO|TVPR|AVMTB/i.test(String(r.ID_CORREDOR || '')));
  }
  const pts = linhasObjetosS2_('CORREDOR_PONTOS').filter(r => ativoS2_(r.ATIVO));

  const parsed = corredores.map(c => {
    const cId = String(c.ID_CORREDOR || '');
    const p = pts.filter(q => String(q.ID_CORREDOR || '') === cId)
      .sort((a, b) => Number(a.ORDEM) - Number(b.ORDEM))
      .map(q => ({ x: numS3_(q.X_NORMALIZADO), y: numS3_(q.Y_NORMALIZADO) }))
      .filter(q => Number.isFinite(q.x) && Number.isFinite(q.y));
    return {
      idCorredor: cId,
      nome: String(c.NOME || ''),
      tolerancia: numS3_(c.TOLERANCIA_NORMALIZADA) || 0.025,
      pontos: p
    };
  }).filter(c => c.pontos.length >= 2);

  _corredoresPolylinesCache_[cacheKey] = parsed;
  return parsed;
}

function localizarCorredorS3_(idMapa,x,y){
  const corredores = obterCorredoresPolylinesS3_(idMapa, false);
  let best=null;
  for (let i = 0; i < corredores.length; i++) {
    const c = corredores[i];
    const r = distPolylineS3_(x, y, c.pontos);
    const tol = c.tolerancia || 0.025;
    if (r.dist <= Math.max(tol, 0.035) && (!best || r.dist < best.distancia)) {
      best = { idCorredor: c.idCorredor, nome: c.nome, distancia: r.dist, percentual: r.percent };
    }
  }
  return best;
}

function localizarCorredorVerticalS3_(idMapa,x,y){
  if (Number.isFinite(x) && x < 0.18 && (String(idMapa||'').includes('AMARELO') || String(idMapa||'').includes('ROXO'))) {
    return {
      idCorredor: String(idMapa||'').includes('AMARELO') ? 'COR-AM-ADMMN' : 'COR-RX-ADMMN',
      nome: 'Avenida Dom Manuel',
      distancia: 0.01,
      percentual: 0.5
    };
  }
  const corredores = obterCorredoresPolylinesS3_(idMapa, true);
  let best=null;
  for (let i = 0; i < corredores.length; i++) {
    const c = corredores[i];
    const r = distPolylineS3_(x, y, c.pontos);
    if (!best || r.dist < best.distancia) {
      best = { idCorredor: c.idCorredor, nome: c.nome, distancia: r.dist, percentual: r.percent };
    }
  }
  return (best&&best.distancia<=0.055)?best:null;
}
function distPolylineS3_(x,y,p){let total=0,lens=[];for(let i=0;i<p.length-1;i++){const l=Math.hypot(p[i+1].x-p[i].x,p[i+1].y-p[i].y);lens.push(l);total+=l;}let best={dist:Infinity,along:0},acc=0;for(let i=0;i<p.length-1;i++){const a=p[i],b=p[i+1],vx=b.x-a.x,vy=b.y-a.y,ll=vx*vx+vy*vy;let t=ll?((x-a.x)*vx+(y-a.y)*vy)/ll:0;t=Math.max(0,Math.min(1,t));const px=a.x+t*vx,py=a.y+t*vy,d=Math.hypot(x-px,y-py);if(d<best.dist)best={dist:d,along:acc+t*lens[i]};acc+=lens[i];}return {dist:best.dist,percent:total?best.along/total:0};}
function localizarSegmentoS3_(idCorredor,pct){const segs=linhasObjetosS2_('SEGMENTOS_CORREDORES').filter(r=>String(r.ID_CORREDOR||'')===idCorredor&&ativoS2_(r.ATIVO));return segs.find(s=>pct>=numS3_(s.PERCENTUAL_INICIO)&&pct<=numS3_(s.PERCENTUAL_FIM))||segs.sort((a,b)=>Math.abs(((numS3_(a.PERCENTUAL_INICIO)+numS3_(a.PERCENTUAL_FIM))/2)-pct)-Math.abs(((numS3_(b.PERCENTUAL_INICIO)+numS3_(b.PERCENTUAL_FIM))/2)-pct))[0]||null;}
function maisProximoS3_(aba,idMapa,x,y,xKeys,yKeys,labelKeys,max){let best=null;linhasObjetosS2_(aba).filter(r=>String(r.ID_MAPA_SETOR||'')===idMapa&&ativoS2_(r.ATIVO)).forEach(r=>{const px=primeiroNumeroS3_(r,xKeys),py=primeiroNumeroS3_(r,yKeys);if(!Number.isFinite(px)||!Number.isFinite(py))return;const d=Math.hypot(x-px,y-py);if(d<=max&&(!best||d<best.dist)){let label='';for(const k of labelKeys){if(String(r[k]??'').trim()){label=String(r[k]).trim();break;}}best={dist:d,label,row:r};}});return best;}
function maisProximoLojaS3_(idMapa,x,y,max){const b=maisProximoS3_('LOJAS_MAPA',idMapa,x,y,['X_NORMALIZADO','CENTRO_X','X_CENTRO','X'],['Y_NORMALIZADO','CENTRO_Y','Y_CENTRO','Y'],['NUMERO_LOJA','LUC','NOME_LOJA'],max);if(!b)return null;return {NUMERO_LOJA:String(b.row.NUMERO_LOJA||''),LUC:String(b.row.LUC||''),NOME_LOJA:String(b.row.NOME_LOJA||''),dist:b.dist};}
function primeiroNumeroS3_(r,keys){for(const k of keys){const n=numS3_(r[k]);if(Number.isFinite(n))return n;}return NaN;}
function numS3_(v){if(typeof v==='number')return v;const s=String(v??'').trim().replace(',','.');return s===''?NaN:Number(s);}
function montarResumoLocalizacaoS3_(mapa,corr,seg,cruz,ref,loja){const p=[`${mapa.nome} • Piso ${mapa.piso}`];if(corr?.nome)p.push(corr.nome);if(seg?.NOME_SEGMENTO)p.push(String(seg.NOME_SEGMENTO));if(cruz?.label)p.push(`Próx. ${cruz.label}`);if(ref?.label)p.push(`Ref. ${ref.label}`);if(loja?.NUMERO_LOJA)p.push(`Loja ${loja.NUMERO_LOJA}${loja.LUC?' • LUC '+loja.LUC:''}`);return p.join(' — ');}

const COLUNAS_CEOP_AS_REGISTROS = Object.freeze([
  'CATEGORIA_REGISTRO',
  'CODIGO_PATRIMONIO',
  'SOLICITANTE_NOME',
  'SOLICITANTE_CPF',
  'SOLICITANTE_CONTATO',
  'SOLICITANTE_EMAIL',
  'TIPO_SOLICITACAO',
  'EMPRESA_INTERNET',
  'DATA_INICIO_SERVICO',
  'DATA_FIM_SERVICO',
  'HORARIO_SERVICO',
  'HORARIO_SEGUNDA',
  'PRESTADOR_NOME',
  'PRESTADOR_CPF',
  'PRESTADOR_CONTATO',
  'PRESTADOR_EMAIL',
  'PRESTADOR_EMPRESA',
  'EQUIPE_AJUDANTES',
  'ITENS_RETIRADA',
  'SERVICO_ESTRUTURA',
  'SERVICO_REVESTIMENTO',
  'SERVICO_INSTALACOES',
  'TERMOS_LGPD_ACEITOS'
]);

/**
 * Garante as colunas necessárias para o CEOP e para o módulo de Autorização para Serviço (AS)
 * na aba REGISTROS, de forma não destrutiva, preservando os dados e formatações já existentes.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [shAlvo] Aba REGISTROS (opcional)
 * @returns {{ ok: boolean, adicionadas: string[], totalColunas: number }}
 */
function garantirColunasCeopAs_(shAlvo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = shAlvo || ss.getSheetByName('REGISTROS');
  if (!sh) {
    throw new Error('Aba REGISTROS não encontrada na planilha.');
  }

  const lastCol = Math.max(1, sh.getLastColumn());
  const atuais = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const faltantes = COLUNAS_CEOP_AS_REGISTROS.filter(c => !atuais.includes(c));

  if (!faltantes.length) {
    return { ok: true, adicionadas: [], totalColunas: atuais.length };
  }

  const maxColAtual = sh.getMaxColumns();
  const colunasRequeridas = lastCol + faltantes.length;
  if (maxColAtual < colunasRequeridas) {
    sh.insertColumnsAfter(maxColAtual, colunasRequeridas - maxColAtual);
  }

  const startCol = lastCol + 1;
  const range = sh.getRange(1, startCol, 1, faltantes.length);
  range.setValues([faltantes])
    .setFontWeight('bold')
    .setBackground('#171B68')
    .setFontColor('#FFFFFF');

  console.log(JSON.stringify({
    evento: 'CEOP_AS_COLUNAS_INSTALADAS',
    totalAdicionadas: faltantes.length,
    colunas: faltantes,
    ts: new Date().toISOString()
  }));

  return {
    ok: true,
    adicionadas: faltantes,
    totalColunas: colunasRequeridas
  };
}

/**
 * RPC para verificação e garantia das colunas de AS por administradores.
 */
function appGarantirColunasCeopAs() {
  exigirPermissaoS14_('administrar');
  return garantirColunasCeopAs_();
}

/**
 * Lista de colunas legadas do antigo projeto de sinalização física (placas).
 */
const COLUNAS_LEGADAS_SINALIZACAO = Object.freeze([
  'TEXTO_SINALIZACAO',
  'MATERIAL',
  'DIMENSOES',
  'COR',
  'FIXACAO',
  'ILUMINADA',
  'DUPLA_FACE',
  'QR_CODE',
  'BRAILLE',
  'PICTOGRAMA',
  'ESTADO_CONSERVACAO',
  'CONDICAO',
  'DATA_INSTALACAO',
  'VALIDADE',
  'DATA_ULTIMA_INSPECAO',
  'PROXIMA_INSPECAO'
]);

/**
 * Remove com segurança as 16 colunas legadas de sinalização física da aba REGISTROS,
 * preservando todas as colunas de auditoria, geolocalização, torres e CEOP/AS.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [shAlvo] Aba REGISTROS (opcional)
 * @returns {{ ok: boolean, totalRemovidas: number, colunasRemovidas: string[], totalColunasRestantes: number }}
 */
function removerColunasLegadasRegistros_(shAlvo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = shAlvo || ss.getSheetByName('REGISTROS');
  if (!sh) throw new Error('Aba REGISTROS não encontrada na planilha.');

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    return { ok: true, totalRemovidas: 0, colunasRemovidas: [], totalColunasRestantes: 0 };
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  const indicesParaDeletar = [];

  headers.forEach((h, idx) => {
    if (COLUNAS_LEGADAS_SINALIZACAO.includes(h)) {
      indicesParaDeletar.push({ nome: h, coluna: idx + 1 });
    }
  });

  if (!indicesParaDeletar.length) {
    return {
      ok: true,
      totalRemovidas: 0,
      colunasRemovidas: [],
      totalColunasRestantes: lastCol,
      mensagem: 'Nenhuma coluna legada de sinalização encontrada na aba REGISTROS.'
    };
  }

  // Deletar da direita para a esquerda (decrescente) para evitar deslocamento de índices
  indicesParaDeletar.sort((a, b) => b.coluna - a.coluna);
  const removidas = [];
  indicesParaDeletar.forEach(item => {
    sh.deleteColumn(item.coluna);
    removidas.push(item.nome);
  });

  console.log(JSON.stringify({
    evento: 'CEOP_COLUNAS_LEGADAS_REMOVIDAS',
    totalRemovidas: removidas.length,
    colunas: removidas,
    totalRestantes: sh.getLastColumn(),
    ts: new Date().toISOString()
  }));

  return {
    ok: true,
    totalRemovidas: removidas.length,
    colunasRemovidas: removidas,
    totalColunasRestantes: sh.getLastColumn()
  };
}

/**
 * RPC para remoção de colunas legadas da aba REGISTROS por administradores.
 */
function appRemoverColunasLegadasRegistros() {
  exigirPermissaoS14_('administrar');
  return removerColunasLegadasRegistros_();
}

function CABECALHOS_REGISTROS_S3_() {
  return [
    'ID_REGISTRO',
    'CLIENT_EVENT_ID',
    'PROTOCOLO',
    'CRIADO_EM',
    'SINCRONIZADO_EM',
    'STATUS',
    'TIPO',
    'FINALIDADE',
    'TITULO',
    'DESCRICAO',
    'RESPONSAVEL',
    'ID_MAPA_SETOR',
    'MAPA',
    'PISO',
    'X_NORMALIZADO',
    'Y_NORMALIZADO',
    'ID_CORREDOR',
    'RUA',
    'ID_SEGMENTO',
    'TRECHO',
    'CRUZAMENTO',
    'REFERENCIA',
    'NUMERO_LOJA',
    'LUC',
    'NOME_LOJA',
    'LOCALIZACAO_CONFIRMADA',
    'ORIGEM',
    'USUARIO',
    'DEVICE_ID',
    'VERSAO_APP',
    'ID_PLANTA_NIVEL',
    'X_NIVEL',
    'Y_NIVEL',
    'ID_TORRE',
    'CODIGO_TORRE',
    'NOME_TORRE',
    'ID_REPRESENTACAO_TORRE',
    'VERSAO_GEOMETRIA_TORRE',
    'ORIGEM_TORRE'
  ].concat(COLUNAS_CEOP_AS_REGISTROS);
}
function garantirCabecalhoS3_(aba,headers){const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(aba);if(!sh)sh=ss.insertSheet(aba);if(sh.getMaxColumns()<headers.length)sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#171B68').setFontColor('#FFFFFF');sh.setFrozenRows(1);}
function diagnosticoS3(){const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),checks=[];check_(checks,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID);check_(checks,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO);check_(checks,'APP_FASE',cfg.APP_FASE===APP.FASE,cfg.APP_FASE);check_(checks,'MODO_DADOS',cfg.MODO_DADOS===APP.MODO_DADOS,cfg.MODO_DADOS);check_(checks,'S2_INSTALADA',cfg.S2_STATUS==='INSTALADO',cfg.S2_STATUS);const sh=ss.getSheetByName('REGISTROS');check_(checks,'ABA_REGISTROS',!!sh,sh?'OK':'ausente');if(sh){const h=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0];const falt=CABECALHOS_REGISTROS_S3_().filter(x=>!h.includes(x));check_(checks,'CABECALHO_REGISTROS',falt.length===0,falt.length?'faltando: '+falt.join(', '):`${CABECALHOS_REGISTROS_S3_().length} colunas`);}const mapas=listarMapasS2_();check_(checks,'MAPAS_ATIVOS',mapas.length===5,`${mapas.length} mapa(s)`);['CORREDORES','CORREDOR_PONTOS','SEGMENTOS_CORREDORES','CRUZAMENTOS','PONTOS_REFERENCIA','LOJAS_MAPA'].forEach(a=>check_(checks,`LOCALIZACAO_${a}`,!!ss.getSheetByName(a),ss.getSheetByName(a)?'OK':'ausente'));const falhas=checks.filter(c=>!c.ok);return {ok:!falhas.length,totalChecks:checks.length,totalFalhas:falhas.length,checks};}
function mostrarDiagnosticoS3(){const d=diagnosticoS3();SpreadsheetApp.getUi().alert('Diagnóstico S3',`${d.ok?'CADASTRO S3 OK':'HÁ PENDÊNCIAS'}\n\n${d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n')}`,SpreadsheetApp.getUi().ButtonSet.OK);return d;}
function atualizarReadmeS3(){const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName('README');if(!sh)sh=ss.insertSheet('README');sh.clear();const rows=[['SINALIZAÇÃO DO MALL — FASE S3'],[''],['VERSÃO'],[APP.VERSAO],[''],['OBJETIVO'],['Cadastrar sinalizações online a partir de um ponto confirmado no mapa.'],[''],['FLUXO'],['Novo registro → marcar ponto → identificar localização → confirmar → preencher formulário → salvar.'],[''],['LOCALIZAÇÃO'],['Corredor e trecho por proximidade da geometria; cruzamento, referência e loja/LUC por proximidade cartográfica.'],[''],['GATE S3'],['1. Executar setupS3().'],['2. Publicar nova versão do Web App.'],['3. Criar um registro no Setor Azul.'],['4. Confirmar localização automática e coordenadas normalizadas.'],['5. Confirmar protocolo SIG-AAAAMMDD-0001 e linha em REGISTROS.'],['6. Repetir em outro mapa/setor.'],['7. Executar diagnosticoS3() sem falhas.'],[''],['NÃO FAZ PARTE DE S3'],['Marcadores persistentes, offline/Outbox e fotos.'],[''],['PRÓXIMA FASE'],['S4 — Marcadores de sinalização no mapa.']];sh.getRange(1,1,rows.length,1).setValues(rows);sh.getRange('A1').setFontWeight('bold').setFontSize(16).setFontColor('#171B68');sh.setColumnWidth(1,900);sh.getRange('A:A').setWrap(true);}


// =========================
// FASE S4 — MARCADORES
// =========================
function setupS4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) throw new Error('Aba CONFIG não encontrada.');

  setConfigValue_(sh, 'APP_VERSAO', 'MVP-1.4-SINALIZACAO-S4', 'Versão atualmente instalada');
  setConfigValue_(sh, 'APP_FASE', 'S4', 'Fase de implementação validada');
  setConfigValue_(sh, 'S4_STATUS', 'INSTALADO', 'Estado da Fase S4');
  setConfigValue_(sh, 'S4_INSTALADO_EM',
    Utilities.formatDate(new Date(), APP.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação da Fase S4'
  );
  return diagnosticoS4();
}

function appListarRegistrosMapaS4(idMapaSetor) {
  const mapaId = String(idMapaSetor || '').trim();
  if (!mapaId) return [];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const head = values.shift().map(String);
  const idx = {};
  head.forEach((h,i)=>idx[h]=i);

  return values
    .filter(r => String(r[idx.ID_MAPA_SETOR] || '') === mapaId)
    .filter(r => !['BAIXADA','EXCLUIDO'].includes(String(r[idx.STATUS] || '').trim().toUpperCase()))
    .map(r => ({
      idRegistro: String(r[idx.ID_REGISTRO] || ''),
      protocolo: String(r[idx.PROTOCOLO] || ''),
      status: String(r[idx.STATUS] || ''),
      tipo: String(r[idx.TIPO] || ''),
      finalidade: String(r[idx.FINALIDADE] || ''),
      titulo: String(r[idx.TITULO] || ''),
      textoSinalizacao: String(r[idx.TEXTO_SINALIZACAO] || ''),
      estadoConservacao: String(r[idx.ESTADO_CONSERVACAO] || ''),
      condicao: String(r[idx.CONDICAO] || ''),
      responsavel: String(r[idx.RESPONSAVEL] || ''),
      mapa: String(r[idx.MAPA] || ''),
      piso: String(r[idx.PISO] || ''),
      x: Number(String(r[idx.X_NORMALIZADO] || '0').replace(',','.')) || 0,
      y: Number(String(r[idx.Y_NORMALIZADO] || '0').replace(',','.')) || 0,
      rua: String(r[idx.RUA] || ''),
      trecho: String(r[idx.TRECHO] || ''),
      cruzamento: String(r[idx.CRUZAMENTO] || ''),
      referencia: String(r[idx.REFERENCIA] || ''),
      numeroLoja: String(r[idx.NUMERO_LOJA] || ''),
      luc: String(r[idx.LUC] || ''),
      nomeLoja: String(r[idx.NOME_LOJA] || ''),
      idPlantaNivel: String(idx.ID_PLANTA_NIVEL == null ? '' : (r[idx.ID_PLANTA_NIVEL] || '')),
      xNivel: idx.X_NIVEL == null || String(r[idx.X_NIVEL] == null ? '' : r[idx.X_NIVEL]).trim() === '' ? null : Number(String(r[idx.X_NIVEL]).replace(',','.')),
      yNivel: idx.Y_NIVEL == null || String(r[idx.Y_NIVEL] == null ? '' : r[idx.Y_NIVEL]).trim() === '' ? null : Number(String(r[idx.Y_NIVEL]).replace(',','.')),
      idTorre: String(idx.ID_TORRE == null ? '' : (r[idx.ID_TORRE] || '')),
      codigoTorre: String(idx.CODIGO_TORRE == null ? '' : (r[idx.CODIGO_TORRE] || '')),
      nomeTorre: String(idx.NOME_TORRE == null ? '' : (r[idx.NOME_TORRE] || '')),
      idRepresentacaoTorre: String(idx.ID_REPRESENTACAO_TORRE == null ? '' : (r[idx.ID_REPRESENTACAO_TORRE] || '')),
      versaoGeometriaTorre: idx.VERSAO_GEOMETRIA_TORRE == null ? 0 : Number(r[idx.VERSAO_GEOMETRIA_TORRE] || 0),
      origemTorre: String(idx.ORIGEM_TORRE == null ? '' : (r[idx.ORIGEM_TORRE] || '')),
      categoriaRegistro: String(idx.CATEGORIA_REGISTRO == null ? '' : (r[idx.CATEGORIA_REGISTRO] || '')),
      asNomeTitular: String(idx.SOLICITANTE_NOME == null ? '' : (r[idx.SOLICITANTE_NOME] || '')),
      asCpfTitular: String(idx.SOLICITANTE_CPF == null ? '' : (r[idx.SOLICITANTE_CPF] || '')),
      asContatoTitular: String(idx.SOLICITANTE_CONTATO == null ? '' : (r[idx.SOLICITANTE_CONTATO] || '')),
      asEmailSolicitante: String(idx.SOLICITANTE_EMAIL == null ? '' : (r[idx.SOLICITANTE_EMAIL] || '')),
      asTipoSolicitacao: String(idx.TIPO_SOLICITACAO == null ? '' : (r[idx.TIPO_SOLICITACAO] || '')),
      asEmpresaInternet: String(idx.EMPRESA_INTERNET == null ? '' : (r[idx.EMPRESA_INTERNET] || '')),
      asDataInicio: String(idx.DATA_INICIO_SERVICO == null ? '' : (r[idx.DATA_INICIO_SERVICO] || '')),
      asDataFim: String(idx.DATA_FIM_SERVICO == null ? '' : (r[idx.DATA_FIM_SERVICO] || '')),
      asHorarioServico: String(idx.HORARIO_SERVICO == null ? '' : (r[idx.HORARIO_SERVICO] || '')),
      asHorarioSegunda: String(idx.HORARIO_SEGUNDA == null ? '' : (r[idx.HORARIO_SEGUNDA] || '')),
      asNomePrestador: String(idx.PRESTADOR_NOME == null ? '' : (r[idx.PRESTADOR_NOME] || '')),
      asCpfPrestador: String(idx.PRESTADOR_CPF == null ? '' : (r[idx.PRESTADOR_CPF] || '')),
      asContatoPrestador: String(idx.PRESTADOR_CONTATO == null ? '' : (r[idx.PRESTADOR_CONTATO] || '')),
      asEmailPrestador: String(idx.PRESTADOR_EMAIL == null ? '' : (r[idx.PRESTADOR_EMAIL] || '')),
      asEmpresaPrestador: String(idx.PRESTADOR_EMPRESA == null ? '' : (r[idx.PRESTADOR_EMPRESA] || '')),
      asEquipe: String(idx.EQUIPE_AJUDANTES == null ? '' : (r[idx.EQUIPE_AJUDANTES] || '')),
      asItensRetirada: String(idx.ITENS_RETIRADA == null ? '' : (r[idx.ITENS_RETIRADA] || '')),
      asServicoEstrutura: String(idx.SERVICO_ESTRUTURA == null ? '' : (r[idx.SERVICO_ESTRUTURA] || '')),
      asServicoRevestimento: String(idx.SERVICO_REVESTIMENTO == null ? '' : (r[idx.SERVICO_REVESTIMENTO] || '')),
      asServicoInstalacoes: String(idx.SERVICO_INSTALACOES == null ? '' : (r[idx.SERVICO_INSTALACOES] || ''))
    }))
    .filter(x => x.x >= 0 && x.x <= 1 && x.y >= 0 && x.y <= 1);
}

function diagnosticoS4() {
  const checks = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('CONFIG');
  const regs = ss.getSheetByName('REGISTROS');
  const c = cfg ? cfg.getDataRange().getValues() : [];
  const obj = {};
  c.slice(1).forEach(r=>obj[String(r[0]||'')]=String(r[1]||''));

  checks.push({ok: obj.APP_FASE === 'S4', nome:'APP_FASE', detalhe:obj.APP_FASE || ''});
  checks.push({ok: obj.APP_VERSAO === 'MVP-1.4-SINALIZACAO-S4', nome:'APP_VERSAO', detalhe:obj.APP_VERSAO || ''});
  checks.push({ok: !!regs, nome:'ABA_REGISTROS', detalhe: regs ? 'OK' : 'AUSENTE'});

  let total = 0, mapas = {};
  if (regs && regs.getLastRow() > 1) {
    const vals = regs.getDataRange().getValues();
    const h = vals.shift().map(String);
    const iMapa = h.indexOf('ID_MAPA_SETOR');
    const iX = h.indexOf('X_NORMALIZADO');
    const iY = h.indexOf('Y_NORMALIZADO');
    vals.forEach(r=>{
      const id=String(r[iMapa]||'');
      const x=String(r[iX]||'');
      const y=String(r[iY]||'');
      if(id && x!=='' && y!==''){ total++; mapas[id]=(mapas[id]||0)+1; }
    });
  }
  checks.push({ok:true, nome:'REGISTROS_COM_COORDENADAS', detalhe:String(total)});
  checks.push({ok:true, nome:'MAPAS_COM_REGISTROS', detalhe:String(Object.keys(mapas).length)});
  return {ok: checks.every(x=>x.ok), totalFalhas: checks.filter(x=>!x.ok).length, checks};
}

function mostrarDiagnosticoS4() {
  const d = diagnosticoS4();
  const linhas = d.checks.map(x => `${x.ok?'✅':'❌'} ${x.nome}: ${x.detalhe}`).join('\n');
  SpreadsheetApp.getUi().alert('Diagnóstico S4', linhas, SpreadsheetApp.getUi().ButtonSet.OK);
  return d;
}
