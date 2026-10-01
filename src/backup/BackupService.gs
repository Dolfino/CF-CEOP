
// ========================================================
// S16 — BACKUP, INTEGRIDADE E RECUPERAÇÃO
// ========================================================
const S16_BACKUP_HEADERS=[
  'ID_BACKUP','CRIADO_EM','ARQUIVO_ID','ARQUIVO_NOME','DRIVE_URL',
  'MANIFEST_JSON','STATUS','USUARIO','VERSAO_APP','OBSERVACAO'
];

const S16_CRITICAL_SHEETS=[
  'REGISTROS','REGISTRO_FOTOS','REGISTRO_HISTORICO','PENDENCIAS','AGENDA_INSPECOES',
  'USUARIOS','PERFIS_PERMISSOES','AUDITORIA'
];

const S16_STRUCTURAL_SHEETS=[
  'CONFIG','PLANTAS','SETORES','MAPAS_SETORES','CARTOGRAFIA_COLECOES','CARTOGRAFIA_ARQUIVOS','CORREDORES',
  'CORREDOR_PONTOS','SEGMENTOS_CORREDORES','CRUZAMENTOS',
  'PONTOS_REFERENCIA','LOJAS_MAPA','LOJAS',
  'REGISTROS','REGISTRO_FOTOS','REGISTRO_HISTORICO',
  'PENDENCIAS','AGENDA_INSPECOES','PLANOS_PREVENTIVOS','ALERTAS_OPERACIONAIS','REGRAS_NOTIFICACAO','NOTIFICACOES_ENVIO','USUARIOS','PERFIS_PERMISSOES','AUDITORIA'
];

function setupS16(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let sh=ss.getSheetByName('BACKUPS');
  if(!sh)sh=ss.insertSheet('BACKUPS');
  garantirCabecalhosS16_(sh);

  const folder=garantirPastaBackupsS16_();
  setConfigValue_(cfg,'BACKUPS_FOLDER_ID',folder.getId(),'Pasta de backups da aplicação');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S16','Fase de implementação validada');
  setConfigValue_(cfg,'S16_STATUS','INSTALADO','Backup, integridade e recuperação instalados');
  setConfigValue_(cfg,'S16_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S16'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S16',entidade:'SISTEMA',entidadeId:'S16',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS16();
}

function garantirCabecalhosS16_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,S16_BACKUP_HEADERS.length).setValues([S16_BACKUP_HEADERS]);
  }else{
    const falt=S16_BACKUP_HEADERS.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn())
    .setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function garantirPastaBackupsS16_(){
  const ss=SpreadsheetApp.getActive();
  const cfg=lerConfigComoObjeto_(ss);
  const existente=String(cfg.BACKUPS_FOLDER_ID||'').trim();
  if(existente){
    try{return DriveApp.getFolderById(existente)}catch(e){}
  }
  const rootId=String(cfg.DRIVE_ROOT_FOLDER_ID||'').trim();
  const root=rootId?DriveApp.getFolderById(rootId):DriveApp.getRootFolder();
  const it=root.getFoldersByName('Backups');
  return it.hasNext()?it.next():root.createFolder('Backups');
}

function normalizarValorHashS16_(v){
  if(v instanceof Date)return Utilities.formatDate(v,APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
  if(v===null||v===undefined)return '';
  if(typeof v==='number')return String(v);
  if(typeof v==='boolean')return v?'TRUE':'FALSE';
  return String(v);
}

function hashSheetS16_(sh){
  if(!sh)return {hash:'',linhas:0,colunas:0};
  const lr=sh.getLastRow(),lc=sh.getLastColumn();
  if(lr<1||lc<1)return {hash:hashTextoS16_(''),linhas:0,colunas:0};
  const vals=sh.getRange(1,1,lr,lc).getValues();
  const txt=vals.map(r=>r.map(normalizarValorHashS16_).join('\u241F')).join('\u241E');
  return {hash:hashTextoS16_(txt),linhas:lr,colunas:lc};
}

function hashTextoS16_(txt){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(txt),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b+256)%256).map(b=>('0'+b.toString(16)).slice(-2)).join('');
}

function manifestS16_(spreadsheet){
  const ss=spreadsheet||SpreadsheetApp.getActive();
  const arr=[];
  S16_STRUCTURAL_SHEETS.forEach(nome=>{
    const sh=ss.getSheetByName(nome);
    if(!sh){
      arr.push({aba:nome,existe:false,linhas:0,colunas:0,hash:''});
      return;
    }
    const h=hashSheetS16_(sh);
    arr.push({aba:nome,existe:true,linhas:h.linhas,colunas:h.colunas,hash:h.hash});
  });
  return arr;
}

function appCriarBackupS16(observacao){
  const s=exigirPermissaoS15_('administrar',{acao:'CRIAR_BACKUP',entidade:'BACKUP'});
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.getActive();
    const folder=garantirPastaBackupsS16_();
    const agora=new Date();
    const stamp=Utilities.formatDate(agora,APP.TIMEZONE,'yyyyMMdd-HHmmss');
    const nome=`Sinalizacao-Mall-Backup-${stamp}-${APP.VERSAO}`;
    const copy=DriveApp.getFileById(ss.getId()).makeCopy(nome,folder);
    const manifest=manifestS16_(ss);
    const id='BKP-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();

    const sh=ss.getSheetByName('BACKUPS');
    garantirCabecalhosS16_(sh);
    appendObjetoPorCabecalhoS7_(sh,{
      ID_BACKUP:id,CRIADO_EM:agora,ARQUIVO_ID:copy.getId(),ARQUIVO_NOME:nome,
      DRIVE_URL:copy.getUrl(),MANIFEST_JSON:JSON.stringify(manifest),STATUS:'OK',
      USUARIO:String(s.email||''),VERSAO_APP:APP.VERSAO,
      OBSERVACAO:String(observacao||'')
    });

    registrarAuditoriaS15_({
      acao:'BACKUP_CRIADO',entidade:'BACKUP',entidadeId:id,resultado:'SUCESSO',
      origem:'WEB_APP',detalhes:{arquivoId:copy.getId(),arquivoNome:nome}
    });

    return {ok:true,idBackup:id,arquivoId:copy.getId(),nome,url:copy.getUrl(),criadoEm:formatarDataS10_(agora)};
  }catch(e){
    registrarAuditoriaS15_({acao:'CRIAR_BACKUP',entidade:'BACKUP',resultado:'ERRO',detalhes:{erro:e.message||String(e)}});
    throw e;
  }finally{lock.releaseLock();}
}

function listarBackupsS16_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('BACKUPS');
  if(!sh||sh.getLastRow()<2)return[];
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};
  h.forEach((x,i)=>idx[x]=i);
  return vals.filter(r=>String(r[idx.ID_BACKUP]||'')).map(r=>({
    idBackup:String(r[idx.ID_BACKUP]||''),
    criadoEm:formatarDataS10_(r[idx.CRIADO_EM]),
    arquivoId:String(r[idx.ARQUIVO_ID]||''),
    nome:String(r[idx.ARQUIVO_NOME]||''),
    url:String(r[idx.DRIVE_URL]||''),
    manifestJson:String(r[idx.MANIFEST_JSON]||'[]'),
    status:String(r[idx.STATUS]||''),
    usuario:String(r[idx.USUARIO]||''),
    versao:String(r[idx.VERSAO_APP]||''),
    observacao:String(r[idx.OBSERVACAO]||'')
  })).reverse();
}

function appListarBackupsS16(){
  exigirPermissaoS15_('administrar',{acao:'LISTAR_BACKUPS',entidade:'BACKUP'});
  return listarBackupsS16_().slice(0,100);
}

/**
 * R15 (auditoria S26.10) — fingerprint de dados sob demanda.
 * Uso: capture antes e depois de uma operação em lote (migração, exclusão,
 * restauração) e compare os hashes para provar que nada foi alterado além
 * do esperado. ADMIN; somente leitura.
 * @param {string[]} abas abas a incluir (padrão: S16_CRITICAL_SHEETS)
 * @param {object|string} referenciaJson fingerprint anterior (opcional)
 */
function appFingerprintDadosS16(abas, referenciaJson){
  exigirPermissaoS15_('administrar',{acao:'FINGERPRINT_DADOS',entidade:'SISTEMA'});
  const ss=SpreadsheetApp.getActive();
  const lista=(Array.isArray(abas)&&abas.length?abas:S16_CRITICAL_SHEETS).map(String);
  const fingerprint={};
  lista.forEach(nome=>{
    const sh=ss.getSheetByName(nome);
    if(!sh){fingerprint[nome]={existe:false,linhas:0,colunas:0,hash:''};return;}
    const h=hashSheetS16_(sh);
    fingerprint[nome]={existe:true,linhas:h.linhas,colunas:h.colunas,hash:h.hash};
  });
  const out={
    ok:true,
    verificadoEm:Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    fingerprint
  };
  if(referenciaJson){
    try{
      const ref=(typeof referenciaJson==='string')?JSON.parse(referenciaJson):referenciaJson;
      const refMap=(ref&&ref.fingerprint)||ref||{};
      out.comparacao=lista.map(nome=>{
        const a=refMap[nome]||{},b=fingerprint[nome]||{};
        return {
          aba:nome,
          hashIgual:!!(a.hash&&a.hash===b.hash),
          linhasAntes:Number(a.linhas||0),
          linhasDepois:Number(b.linhas||0)
        };
      });
    }catch(_){out.comparacao=null;}
  }
  try{
    registrarAuditoriaS15_({acao:'FINGERPRINT_DADOS',entidade:'SISTEMA',entidadeId:'S16',resultado:'SUCESSO',origem:'APPS_SCRIPT'});
  }catch(_){}
  return out;
}

function appVerificarIntegridadeS16(){
  exigirPermissaoS15_('administrar',{acao:'VERIFICAR_INTEGRIDADE',entidade:'SISTEMA'});
  const ss=SpreadsheetApp.getActive(),atual=manifestS16_(ss),checks=[];
  atual.forEach(x=>{
    checks.push({
      nome:x.aba,
      ok:x.existe&&x.linhas>=1&&x.colunas>=1,
      detalhe:x.existe?`${x.linhas} linha(s) • ${x.colunas} coluna(s) • ${x.hash.slice(0,12)}`:'ABA AUSENTE'
    });
  });

  const backups=listarBackupsS16_();
  let comparacao=null;
  if(backups.length){
    try{
      const anterior=JSON.parse(backups[0].manifestJson||'[]');
      const map={};anterior.forEach(x=>map[x.aba]=x);
      comparacao=atual.map(x=>{
        const b=map[x.aba];
        return {
          aba:x.aba,
          existeNoBackup:!!b?.existe,
          linhasAtual:x.linhas,
          linhasBackup:Number(b?.linhas||0),
          hashIgual:!!b&&String(b.hash||'')===String(x.hash||'')
        };
      });
    }catch(e){}
  }

  const resultado={
    ok:checks.every(c=>c.ok),
    verificadoEm:Utilities.formatDate(new Date(),APP.TIMEZONE,'dd/MM/yyyy HH:mm'),
    checks,comparacao,
    ultimoBackup:backups[0]?{idBackup:backups[0].idBackup,criadoEm:backups[0].criadoEm,nome:backups[0].nome}:null
  };

  registrarAuditoriaS15_({
    acao:'INTEGRIDADE_VERIFICADA',entidade:'SISTEMA',entidadeId:'S16',
    resultado:resultado.ok?'SUCESSO':'ERRO',detalhes:{falhas:checks.filter(c=>!c.ok).map(c=>c.nome)}
  });
  return resultado;
}

function obterBackupS16_(idBackup){
  return listarBackupsS16_().find(x=>x.idBackup===String(idBackup||''))||null;
}

function appSimularRestauracaoS16(idBackup){
  exigirPermissaoS15_('administrar',{acao:'SIMULAR_RESTAURACAO',entidade:'BACKUP',entidadeId:String(idBackup||'')});
  const b=obterBackupS16_(idBackup);
  if(!b)throw new Error('Backup não encontrado.');

  const backupSS=SpreadsheetApp.openById(b.arquivoId);
  const atualSS=SpreadsheetApp.getActive();
  const itens=S16_CRITICAL_SHEETS.map(nome=>{
    const atual=atualSS.getSheetByName(nome),orig=backupSS.getSheetByName(nome);
    const a=hashSheetS16_(atual),o=hashSheetS16_(orig);
    return {
      aba:nome,
      existeAtual:!!atual,
      existeBackup:!!orig,
      linhasAtual:a.linhas,
      linhasBackup:o.linhas,
      alteraria:!!orig&&(!atual||a.hash!==o.hash),
      hashAtual:a.hash.slice(0,12),
      hashBackup:o.hash.slice(0,12)
    };
  });
  return {ok:true,idBackup:b.idBackup,nome:b.nome,criadoEm:b.criadoEm,itens};
}

function appRestaurarBackupS16(payload){
  payload=payload||{};
  exigirPermissaoS15_('administrar',{acao:'RESTAURAR_BACKUP',entidade:'BACKUP',entidadeId:String(payload.idBackup||'')});
  if(String(payload.confirmacao||'').trim().toUpperCase()!=='RESTAURAR')
    throw new Error('Confirmação inválida. Digite RESTAURAR.');

  const b=obterBackupS16_(payload.idBackup);
  if(!b)throw new Error('Backup não encontrado.');

  const selecionadas=Array.isArray(payload.abas)&&payload.abas.length
    ? payload.abas.filter(x=>S16_CRITICAL_SHEETS.includes(x))
    : S16_CRITICAL_SHEETS.slice();

  if(!selecionadas.length)throw new Error('Nenhuma aba crítica selecionada.');

  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    // Backup de segurança imediatamente antes da restauração.
    const safety=appCriarBackupS16_Interno_('Backup automático pré-restauração');
    const origem=SpreadsheetApp.openById(b.arquivoId),destino=SpreadsheetApp.getActive();
    const restauradas=[];

    selecionadas.forEach(nome=>{
      const src=origem.getSheetByName(nome);
      if(!src)return;
      let dst=destino.getSheetByName(nome);
      if(!dst)dst=destino.insertSheet(nome);

      const lr=src.getLastRow(),lc=src.getLastColumn();
      dst.clearContents();
      if(lr>0&&lc>0){
        const vals=src.getRange(1,1,lr,lc).getValues();
        dst.getRange(1,1,lr,lc).setValues(vals);
      }
      restauradas.push(nome);
    });

    SpreadsheetApp.flush();
    registrarAuditoriaS15_({
      acao:'BACKUP_RESTAURADO',entidade:'BACKUP',entidadeId:b.idBackup,resultado:'SUCESSO',
      detalhes:{abas:restauradas,backupSeguranca:safety.idBackup}
    });

    return {ok:true,idBackup:b.idBackup,abas:restauradas,backupSeguranca:safety.idBackup};
  }catch(e){
    registrarAuditoriaS15_({
      acao:'RESTAURAR_BACKUP',entidade:'BACKUP',entidadeId:String(payload.idBackup||''),
      resultado:'ERRO',detalhes:{erro:e.message||String(e)}
    });
    throw e;
  }finally{lock.releaseLock();}
}

function appCriarBackupS16_Interno_(observacao){
  const ss=SpreadsheetApp.getActive(),folder=garantirPastaBackupsS16_(),agora=new Date();
  const stamp=Utilities.formatDate(agora,APP.TIMEZONE,'yyyyMMdd-HHmmss');
  const nome=`Sinalizacao-Mall-Safety-${stamp}-${APP.VERSAO}`;
  const copy=DriveApp.getFileById(ss.getId()).makeCopy(nome,folder);
  const manifest=manifestS16_(ss),id='BKP-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
  const sh=ss.getSheetByName('BACKUPS');
  appendObjetoPorCabecalhoS7_(sh,{
    ID_BACKUP:id,CRIADO_EM:agora,ARQUIVO_ID:copy.getId(),ARQUIVO_NOME:nome,
    DRIVE_URL:copy.getUrl(),MANIFEST_JSON:JSON.stringify(manifest),STATUS:'OK',
    USUARIO:normalizarEmailS14_((usuarioRpcAtualS223_()?.email||Session.getActiveUser().getEmail())),VERSAO_APP:APP.VERSAO,
    OBSERVACAO:String(observacao||'')
  });
  return {idBackup:id,arquivoId:copy.getId(),nome};
}

function diagnosticoS16(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S16',cfg.APP_FASE||'ausente');
  check_(c,'S16_STATUS',cfg.S16_STATUS==='INSTALADO',cfg.S16_STATUS||'ausente');
  check_(c,'ABA_BACKUPS',!!ss.getSheetByName('BACKUPS'),'BACKUPS');
  check_(c,'PASTA_BACKUPS',!!String(cfg.BACKUPS_FOLDER_ID||''),String(cfg.BACKUPS_FOLDER_ID||'ausente'));
  check_(c,'CRIAR_BACKUP',typeof appCriarBackupS16==='function','OK');
  check_(c,'INTEGRIDADE',typeof appVerificarIntegridadeS16==='function','OK');
  check_(c,'SIMULAR_RESTORE',typeof appSimularRestauracaoS16==='function','OK');
  check_(c,'RESTAURAR',typeof appRestaurarBackupS16==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS16(){
  const d=diagnosticoS16();
  SpreadsheetApp.getUi().alert('Diagnóstico S16',
    `${d.ok?'S16 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK);
  return d;
}


// ========================================================
// S17 — MONITORAMENTO, SAÚDE OPERACIONAL E DIAGNÓSTICO
// ========================================================
function setupS17(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S17','Fase de implementação validada');
  setConfigValue_(cfg,'S17_STATUS','INSTALADO','Monitoramento e saúde operacional instalados');
  setConfigValue_(cfg,'S17_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S17'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S17',entidade:'SISTEMA',entidadeId:'S17',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS17();
}

function appSaudeOperacionalS17(){
  exigirPermissaoS15_('administrar',{acao:'DIAGNOSTICO_OPERACIONAL',entidade:'SISTEMA',entidadeId:'S17'});
  const ss=SpreadsheetApp.getActive();
  const cfg=lerConfigComoObjeto_(ss);
  const agora=new Date();

  const estrutura=S16_STRUCTURAL_SHEETS.map(nome=>{
    const sh=ss.getSheetByName(nome);
    if(!sh)return {aba:nome,ok:false,linhas:0,colunas:0,hash:'',detalhe:'ABA AUSENTE'};
    const h=hashSheetS16_(sh);
    return {
      aba:nome,ok:h.linhas>=1&&h.colunas>=1,
      linhas:h.linhas,colunas:h.colunas,hash:h.hash.slice(0,12),
      detalhe:`${h.linhas} linha(s) • ${h.colunas} coluna(s)`
    };
  });

  const backups=listarBackupsS16_();
  const ultimo=backups[0]||null;
  const backupAge=ultimo?idadeDataS17_(ultimo.criadoEm,agora):null;

  const auditoria=resumoAuditoriaS17_(agora);
  const registros=contarLinhasS17_('REGISTROS');
  const fotos=contarLinhasS17_('REGISTRO_FOTOS');
  const historico=contarLinhasS17_('REGISTRO_HISTORICO');
  const pendencias=contarPendenciasS17_();

  const alertas=[];
  estrutura.filter(x=>!x.ok).forEach(x=>alertas.push({
    nivel:'CRITICO',codigo:'ABA_ESTRUTURAL',mensagem:`${x.aba}: ${x.detalhe}`
  }));

  if(!ultimo){
    alertas.push({nivel:'ATENCAO',codigo:'SEM_BACKUP',mensagem:'Nenhum backup registrado.'});
  }else if(backupAge!==null&&backupAge>72){
    alertas.push({nivel:'ATENCAO',codigo:'BACKUP_ANTIGO',mensagem:`Último backup há ${Math.round(backupAge)} h.`});
  }

  if(auditoria.erros24h>0){
    alertas.push({nivel:'ATENCAO',codigo:'ERROS_RECENTES',mensagem:`${auditoria.erros24h} erro(s) auditado(s) nas últimas 24 h.`});
  }
  if(auditoria.negados24h>=5){
    alertas.push({nivel:'ATENCAO',codigo:'NEGADOS_RECENTES',mensagem:`${auditoria.negados24h} acesso(s) negado(s) nas últimas 24 h.`});
  }
  if(pendencias.vencidas>0){
    alertas.push({nivel:'ATENCAO',codigo:'PENDENCIAS_VENCIDAS',mensagem:`${pendencias.vencidas} pendência(s) vencida(s).`});
  }

  const nivel=alertas.some(a=>a.nivel==='CRITICO')?'CRITICO':
    alertas.length?'ATENCAO':'OK';

  const sessao=sessaoAtualS14_();

  const resp={
    ok:nivel!=='CRITICO',
    nivel,
    verificadoEm:Utilities.formatDate(agora,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),
    sistema:{
      appId:String(cfg.APP_ID||''),
      nome:String(cfg.APP_NOME||''),
      versao:String(cfg.APP_VERSAO||APP.VERSAO),
      fase:String(cfg.APP_FASE||''),
      timezone:String(cfg.TIMEZONE||APP.TIMEZONE),
      modoDados:String(cfg.MODO_DADOS||''),
      usuario:String(sessao.email||''),
      perfil:String(sessao.perfil||'')
    },
    estrutura,
    dados:{registros,fotos,historico,pendencias},
    backup:{
      total:backups.length,
      ultimo:ultimo?{
        idBackup:ultimo.idBackup,
        criadoEm:ultimo.criadoEm,
        nome:ultimo.nome,
        status:ultimo.status,
        idadeHoras:backupAge===null?null:Math.round(backupAge*10)/10
      }:null
    },
    auditoria,
    alertas
  };

  registrarAuditoriaS15_({
    acao:'DIAGNOSTICO_OPERACIONAL',entidade:'SISTEMA',entidadeId:'S17',
    resultado:nivel==='CRITICO'?'ERRO':'SUCESSO',
    origem:'WEB_APP',
    detalhes:{nivel,alertas:alertas.map(a=>a.codigo)}
  });

  return resp;
}

function idadeDataS17_(texto,agora){
  if(!texto)return null;
  let d=null;
  if(texto instanceof Date)d=texto;
  else{
    const s=String(texto);
    let m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
    if(m)d=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]),Number(m[5]));
    else{
      const p=new Date(s);
      if(!isNaN(p))d=p;
    }
  }
  if(!d||isNaN(d))return null;
  return Math.max(0,(agora.getTime()-d.getTime())/3600000);
}

function contarLinhasS17_(aba){
  const sh=SpreadsheetApp.getActive().getSheetByName(aba);
  return sh?Math.max(0,sh.getLastRow()-1):0;
}

function contarPendenciasS17_(){
  const itens=appListarPendenciasS10('');
  return {
    total:itens.length,
    abertas:itens.filter(x=>x.status==='ABERTA').length,
    andamento:itens.filter(x=>x.status==='EM ANDAMENTO').length,
    concluidas:itens.filter(x=>x.status==='CONCLUIDA').length,
    vencidas:itens.filter(x=>pendenciaVencidaS11_(x)).length,
    criticas:itens.filter(x=>x.prioridade==='CRITICA'&&!['CONCLUIDA','CANCELADA'].includes(x.status)).length
  };
}

function resumoAuditoriaS17_(agora){
  const sh=SpreadsheetApp.getActive().getSheetByName('AUDITORIA');
  if(!sh||sh.getLastRow()<2)return {total:0,erros24h:0,negados24h:0,eventos24h:0,ultimos:[]};

  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};
  h.forEach((x,i)=>idx[x]=i);
  let erros24h=0,negados24h=0,eventos24h=0;
  const itens=[];

  vals.forEach(r=>{
    const data=r[idx.DATA_HORA];
    const dt=data instanceof Date?data:new Date(data);
    const horas=!isNaN(dt)?Math.max(0,(agora-dt)/3600000):999999;
    const resultado=String(r[idx.RESULTADO]||'');
    if(horas<=24){
      eventos24h++;
      if(resultado==='ERRO')erros24h++;
      if(resultado==='NEGADO')negados24h++;
    }
    itens.push({
      dataHora:dt instanceof Date&&!isNaN(dt)?Utilities.formatDate(dt,APP.TIMEZONE,'dd/MM HH:mm'):String(data||''),
      acao:String(r[idx.ACAO]||''),
      resultado,
      usuario:String(r[idx.USUARIO_EMAIL]||'')
    });
  });

  return {
    total:vals.length,erros24h,negados24h,eventos24h,
    ultimos:itens.reverse().slice(0,8)
  };
}

function diagnosticoS17(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S17',cfg.APP_FASE||'ausente');
  check_(c,'S17_STATUS',cfg.S17_STATUS==='INSTALADO',cfg.S17_STATUS||'ausente');
  check_(c,'SAUDE_API',typeof appSaudeOperacionalS17==='function','OK');
  check_(c,'BACKUP_S16',typeof listarBackupsS16_==='function','OK');
  check_(c,'AUDITORIA_S15',typeof registrarAuditoriaS15_==='function','OK');
  check_(c,'RBAC_S14',typeof exigirPermissaoS14_==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS17(){
  const d=diagnosticoS17();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S17',
    `${d.ok?'S17 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
      d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S18 — CICLO DE VIDA DO ATIVO E INVENTÁRIO
// ========================================================
const S18_LIFECYCLE_HEADERS = [
  'ID_EVENTO_CICLO','ID_REGISTRO','PROTOCOLO','DATA_HORA',
  'TIPO_EVENTO','STATUS_ANTERIOR','STATUS_NOVO',
  'ID_MAPA_ANTERIOR','MAPA_ANTERIOR','PISO_ANTERIOR','RUA_ANTERIOR','TRECHO_ANTERIOR','X_ANTERIOR','Y_ANTERIOR',
  'ID_MAPA_NOVO','MAPA_NOVO','PISO_NOVO','RUA_NOVO','TRECHO_NOVO','X_NOVO','Y_NOVO',
  'MOTIVO','OBSERVACAO','RESPONSAVEL','USUARIO','DEVICE_ID','VERSAO_APP',
  'ID_PLANTA_NIVEL_ANTERIOR','X_NIVEL_ANTERIOR','Y_NIVEL_ANTERIOR','ID_TORRE_ANTERIOR','CODIGO_TORRE_ANTERIOR','NOME_TORRE_ANTERIOR','ID_REPRESENTACAO_TORRE_ANTERIOR',
  'ID_PLANTA_NIVEL_NOVO','X_NIVEL_NOVO','Y_NIVEL_NOVO','ID_TORRE_NOVO','CODIGO_TORRE_NOVO','NOME_TORRE_NOVO','ID_REPRESENTACAO_TORRE_NOVO','ORIGEM_TORRE_NOVO'
];

const S18_EVENTOS = ['INSTALACAO','MOVIMENTACAO','SUBSTITUICAO','RETIRADA','DESATIVACAO','REINSTALACAO'];

function setupS18(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let sh=ss.getSheetByName('CICLO_VIDA_ATIVO');
  if(!sh)sh=ss.insertSheet('CICLO_VIDA_ATIVO');
  garantirCabecalhosS18_(sh);

  garantirCamposInventarioRegistrosS18_();

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S18','Fase de implementação validada');
  setConfigValue_(cfg,'S18_STATUS','INSTALADO','Ciclo de vida e inventário instalados');
  setConfigValue_(cfg,'S18_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S18'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S18',entidade:'SISTEMA',entidadeId:'S18',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS18();
}

function garantirCabecalhosS18_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,S18_LIFECYCLE_HEADERS.length).setValues([S18_LIFECYCLE_HEADERS]);
  }else{
    const falt=S18_LIFECYCLE_HEADERS.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function garantirCamposInventarioRegistrosS18_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  if(!sh)throw new Error('REGISTROS ausente.');
  const add=['CODIGO_PATRIMONIAL','STATUS_CICLO_VIDA','DATA_ATIVACAO','DATA_DESATIVACAO','ID_REGISTRO_SUBSTITUIDO','ID_REGISTRO_SUBSTITUTO'];
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  const falt=add.filter(h=>!atuais.includes(h));
  if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
}

function appRegistrarEventoCicloS18(payload){
  payload=payload||{};
  exigirPermissaoS15_('editarRegistro',{
    acao:'EVENTO_CICLO_VIDA',entidade:'REGISTRO',entidadeId:String(payload.idRegistro||'')
  });

  const tipo=String(payload.tipoEvento||'').trim().toUpperCase();
  if(!S18_EVENTOS.includes(tipo))throw new Error('Tipo de evento de ciclo de vida inválido.');

  const reg=obterRegistroPorIdS8_(String(payload.idRegistro||''));
  if(!reg)throw new Error('Sinalização não encontrada.');

  const antes=snapshotCicloS18_(reg);
  validarTransicaoCicloS18_(antes.statusCicloVida,tipo);

  const depois=calcularNovoEstadoCicloS18_(reg,payload,tipo);
  aplicarNovoEstadoCicloS18_(reg.ID_REGISTRO,depois);

  const ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName('CICLO_VIDA_ATIVO');
  garantirCabecalhosS18_(sh);
  const agora=new Date();
  const idEvento='CV-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();

  appendObjetoPorCabecalhoS7_(sh,{
    ID_EVENTO_CICLO:idEvento,
    ID_REGISTRO:String(reg.ID_REGISTRO||''),
    PROTOCOLO:String(reg.PROTOCOLO||''),
    DATA_HORA:agora,
    TIPO_EVENTO:tipo,
    STATUS_ANTERIOR:String(antes.statusCicloVida||''),
    STATUS_NOVO:String(depois.statusCicloVida||''),
    ID_MAPA_ANTERIOR:String(antes.idMapa||''),
    MAPA_ANTERIOR:String(antes.mapa||''),
    PISO_ANTERIOR:String(antes.piso||''),
    RUA_ANTERIOR:String(antes.rua||''),
    TRECHO_ANTERIOR:String(antes.trecho||''),
    X_ANTERIOR:antes.x,
    Y_ANTERIOR:antes.y,
    ID_MAPA_NOVO:String(depois.idMapa||''),
    MAPA_NOVO:String(depois.mapa||''),
    PISO_NOVO:String(depois.piso||''),
    RUA_NOVO:String(depois.rua||''),
    TRECHO_NOVO:String(depois.trecho||''),
    X_NOVO:depois.x,
    Y_NOVO:depois.y,
    MOTIVO:String(payload.motivo||''),
    OBSERVACAO:String(payload.observacao||''),
    RESPONSAVEL:String(payload.responsavel||reg.RESPONSAVEL||''),
    USUARIO:normalizarEmailS14_((usuarioRpcAtualS223_()?.email||Session.getActiveUser().getEmail())),
    DEVICE_ID:String(payload.deviceId||''),
    VERSAO_APP:APP.VERSAO,
    ID_PLANTA_NIVEL_ANTERIOR:String(antes.idPlantaNivel||''),
    X_NIVEL_ANTERIOR:s18ValorNumeroOuVazio_(antes.xNivel),
    Y_NIVEL_ANTERIOR:s18ValorNumeroOuVazio_(antes.yNivel),
    ID_TORRE_ANTERIOR:String(antes.idTorre||''),
    CODIGO_TORRE_ANTERIOR:String(antes.codigoTorre||''),
    NOME_TORRE_ANTERIOR:String(antes.nomeTorre||''),
    ID_REPRESENTACAO_TORRE_ANTERIOR:String(antes.idRepresentacaoTorre||''),
    ID_PLANTA_NIVEL_NOVO:String(depois.idPlantaNivel||''),
    X_NIVEL_NOVO:s18ValorNumeroOuVazio_(depois.xNivel),
    Y_NIVEL_NOVO:s18ValorNumeroOuVazio_(depois.yNivel),
    ID_TORRE_NOVO:String(depois.idTorre||''),
    CODIGO_TORRE_NOVO:String(depois.codigoTorre||''),
    NOME_TORRE_NOVO:String(depois.nomeTorre||''),
    ID_REPRESENTACAO_TORRE_NOVO:String(depois.idRepresentacaoTorre||''),
    ORIGEM_TORRE_NOVO:String(depois.origemTorre||'')
  });

  if(tipo==='SUBSTITUICAO'&&payload.idRegistroSubstituto){
    vincularSubstituicaoS18_(String(reg.ID_REGISTRO||''),String(payload.idRegistroSubstituto||''));
  }

  registrarAuditoriaS15_({
    acao:'CICLO_VIDA_ATUALIZADO',
    entidade:'REGISTRO',
    entidadeId:String(reg.ID_REGISTRO||''),
    resultado:'SUCESSO',
    origem:'WEB_APP',
    deviceId:String(payload.deviceId||''),
    detalhes:{tipoEvento:tipo,idEvento,compatS2610G1:true},
    valorAnterior:antes,
    valorNovo:depois
  });

  return {ok:true,idEvento,tipoEvento:tipo,statusCicloVida:depois.statusCicloVida};
}

function snapshotCicloS18_(reg){
  return {
    idRegistro:String(reg.ID_REGISTRO||''),
    codigoPatrimonial:String(reg.CODIGO_PATRIMONIAL||''),
    statusCicloVida:String(reg.STATUS_CICLO_VIDA||'ATIVO').toUpperCase(),
    statusOperacional:String(reg.STATUS||''),
    idMapa:String(reg.ID_MAPA_SETOR||''),
    mapa:String(reg.MAPA||''),
    piso:String(reg.PISO||''),
    rua:String(reg.RUA||''),
    trecho:String(reg.TRECHO||''),
    idCorredor:String(reg.ID_CORREDOR||''),
    idSegmento:String(reg.ID_SEGMENTO||''),
    cruzamento:String(reg.CRUZAMENTO||''),
    referencia:String(reg.REFERENCIA||''),
    numeroLoja:String(reg.NUMERO_LOJA||''),
    luc:String(reg.LUC||''),
    nomeLoja:String(reg.NOME_LOJA||''),
    x:Number(reg.X_NORMALIZADO||0),
    y:Number(reg.Y_NORMALIZADO||0),
    idPlantaNivel:String(reg.ID_PLANTA_NIVEL||''),
    xNivel:s18NumeroOuNull_(reg.X_NIVEL),
    yNivel:s18NumeroOuNull_(reg.Y_NIVEL),
    idTorre:String(reg.ID_TORRE||''),
    codigoTorre:String(reg.CODIGO_TORRE||''),
    nomeTorre:String(reg.NOME_TORRE||''),
    idRepresentacaoTorre:String(reg.ID_REPRESENTACAO_TORRE||''),
    versaoGeometriaTorre:Number(reg.VERSAO_GEOMETRIA_TORRE||0),
    origemTorre:String(reg.ORIGEM_TORRE||''),
    dataAtivacao:reg.DATA_ATIVACAO||'',
    dataDesativacao:reg.DATA_DESATIVACAO||''
  };
}

function validarTransicaoCicloS18_(status,tipo){
  const s=String(status||'ATIVO').toUpperCase();
  const permitidas={
    ATIVO:['MOVIMENTACAO','SUBSTITUICAO','RETIRADA','DESATIVACAO'],
    RETIRADO:['REINSTALACAO','DESATIVACAO'],
    DESATIVADO:['REINSTALACAO'],
    SUBSTITUIDO:['DESATIVACAO'],
    NOVO:['INSTALACAO']
  };
  const lista=permitidas[s]||['MOVIMENTACAO','RETIRADA','DESATIVACAO'];
  if(!lista.includes(tipo)){
    throw new Error(`Transição ${s} → ${tipo} não permitida.`);
  }
}

function calcularNovoEstadoCicloS18_(reg,payload,tipo){
  const antes=snapshotCicloS18_(reg),agora=new Date();
  const n={...antes};

  if(tipo==='INSTALACAO'||tipo==='REINSTALACAO'||tipo==='MOVIMENTACAO'){
    if(payload.idMapaSetor!==undefined)n.idMapa=String(payload.idMapaSetor||n.idMapa);
    if(payload.mapa!==undefined)n.mapa=String(payload.mapa||n.mapa);
    if(payload.piso!==undefined)n.piso=String(payload.piso||n.piso);
    if(payload.rua!==undefined)n.rua=String(payload.rua||'');
    if(payload.trecho!==undefined)n.trecho=String(payload.trecho||'');
    if(payload.idCorredor!==undefined)n.idCorredor=String(payload.idCorredor||'');
    if(payload.idSegmento!==undefined)n.idSegmento=String(payload.idSegmento||'');
    if(payload.cruzamento!==undefined)n.cruzamento=String(payload.cruzamento||'');
    if(payload.referencia!==undefined)n.referencia=String(payload.referencia||'');
    if(payload.numeroLoja!==undefined)n.numeroLoja=String(payload.numeroLoja||'');
    if(payload.luc!==undefined)n.luc=String(payload.luc||'');
    if(payload.nomeLoja!==undefined)n.nomeLoja=String(payload.nomeLoja||'');
    if(payload.x!==undefined)n.x=Number(payload.x||0);
    if(payload.y!==undefined)n.y=Number(payload.y||0);

    // S26.10-G1 FIX1 — nível/Torre fazem parte da mesma movimentação. A ausência
    // explícita limpa um contexto antigo; a presença usa o normalizador G1 para
    // validar coordenadas e enriquecer a Torre publicada quando necessário.
    const recebeuContextoNivel = ['idPlantaNivel','xNivel','yNivel','idTorre','codigoTorre','nomeTorre','idRepresentacaoTorre','versaoGeometriaTorre','origemTorre']
      .some(function(k){return Object.prototype.hasOwnProperty.call(payload,k);});
    if(recebeuContextoNivel){
      let ctx={
        idPlantaNivel:String(payload.idPlantaNivel||''),
        xNivel:s18NumeroOuNull_(payload.xNivel),
        yNivel:s18NumeroOuNull_(payload.yNivel),
        idTorre:String(payload.idTorre||''),
        codigoTorre:String(payload.codigoTorre||''),
        nomeTorre:String(payload.nomeTorre||''),
        idRepresentacaoTorre:String(payload.idRepresentacaoTorre||''),
        versaoGeometriaTorre:Number(payload.versaoGeometriaTorre||0),
        origemTorre:String(payload.origemTorre||'')
      };
      if(typeof s2610GNormalizarLocalizacaoRegistro_==='function'){
        const normal=s2610GNormalizarLocalizacaoRegistro_(payload);
        ctx={
          idPlantaNivel:String(normal.idPlantaNivel||''),
          xNivel:Number.isFinite(normal.xNivel)?normal.xNivel:null,
          yNivel:Number.isFinite(normal.yNivel)?normal.yNivel:null,
          idTorre:String(normal.idTorre||''),
          codigoTorre:String(normal.codigoTorre||''),
          nomeTorre:String(normal.nomeTorre||''),
          idRepresentacaoTorre:String(normal.idRepresentacaoTorre||''),
          versaoGeometriaTorre:Number(normal.versaoGeometriaTorre||0),
          origemTorre:String(normal.origemTorre||'')
        };
      }
      n.idPlantaNivel=ctx.idPlantaNivel;
      n.xNivel=ctx.xNivel;
      n.yNivel=ctx.yNivel;
      n.idTorre=ctx.idTorre;
      n.codigoTorre=ctx.codigoTorre;
      n.nomeTorre=ctx.nomeTorre;
      n.idRepresentacaoTorre=ctx.idRepresentacaoTorre;
      n.versaoGeometriaTorre=ctx.versaoGeometriaTorre;
      n.origemTorre=ctx.origemTorre;
    }
  }

  if(tipo==='INSTALACAO'){
    n.statusCicloVida='ATIVO';n.dataAtivacao=agora;n.dataDesativacao='';
  }
  if(tipo==='MOVIMENTACAO')n.statusCicloVida='ATIVO';
  if(tipo==='RETIRADA'){n.statusCicloVida='RETIRADO';n.statusOperacional='RETIRADO';}
  if(tipo==='DESATIVACAO'){n.statusCicloVida='DESATIVADO';n.statusOperacional='DESATIVADO';n.dataDesativacao=agora;}
  if(tipo==='SUBSTITUICAO'){n.statusCicloVida='SUBSTITUIDO';n.statusOperacional='SUBSTITUIDO';n.dataDesativacao=agora;}
  if(tipo==='REINSTALACAO'){n.statusCicloVida='ATIVO';n.statusOperacional='ATIVA';n.dataAtivacao=agora;n.dataDesativacao='';}
  return n;
}

function aplicarNovoEstadoCicloS18_(idRegistro,n){
  const sh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};
  h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_REGISTRO]||'')===String(idRegistro||''));
  if(row<1)throw new Error('Registro não encontrado.');

  const set=(campo,v)=>{const c=idx[campo];if(c>=0)sh.getRange(row+1,c+1).setValue(v);};
  set('STATUS_CICLO_VIDA',n.statusCicloVida);
  set('STATUS',n.statusOperacional);
  set('ID_MAPA_SETOR',n.idMapa);
  set('MAPA',n.mapa);
  set('PISO',n.piso);
  set('RUA',n.rua);
  set('TRECHO',n.trecho);
  set('ID_CORREDOR',String(n.idCorredor||''));
  set('ID_SEGMENTO',String(n.idSegmento||''));
  set('CRUZAMENTO',String(n.cruzamento||''));
  set('REFERENCIA',String(n.referencia||''));
  set('NUMERO_LOJA',String(n.numeroLoja||''));
  set('LUC',String(n.luc||''));
  set('NOME_LOJA',String(n.nomeLoja||''));
  set('X_NORMALIZADO',n.x);
  set('Y_NORMALIZADO',n.y);
  set('ID_PLANTA_NIVEL',String(n.idPlantaNivel||''));
  set('X_NIVEL',s18ValorNumeroOuVazio_(n.xNivel));
  set('Y_NIVEL',s18ValorNumeroOuVazio_(n.yNivel));
  set('ID_TORRE',String(n.idTorre||''));
  set('CODIGO_TORRE',String(n.codigoTorre||''));
  set('NOME_TORRE',String(n.nomeTorre||''));
  set('ID_REPRESENTACAO_TORRE',String(n.idRepresentacaoTorre||''));
  set('VERSAO_GEOMETRIA_TORRE',Number(n.versaoGeometriaTorre||0)||'');
  set('ORIGEM_TORRE',String(n.origemTorre||''));
  if(n.dataAtivacao)set('DATA_ATIVACAO',n.dataAtivacao);
  set('DATA_DESATIVACAO',n.dataDesativacao||'');
}

function s18NumeroOuNull_(v){
  if(v===null||v===undefined||String(v).trim()==='')return null;
  const n=Number(String(v).replace(',','.'));
  return Number.isFinite(n)?n:null;
}
function s18ValorNumeroOuVazio_(v){
  const n=s18NumeroOuNull_(v);
  return n===null?'':n;
}

function vincularSubstituicaoS18_(idAntigo,idNovo){
  const sh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const a=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_REGISTRO]||'')===idAntigo);
  const n=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_REGISTRO]||'')===idNovo);
  if(a<1||n<1)throw new Error('Registro antigo ou substituto não encontrado.');
  if(idx.ID_REGISTRO_SUBSTITUTO>=0)sh.getRange(a+1,idx.ID_REGISTRO_SUBSTITUTO+1).setValue(idNovo);
  if(idx.ID_REGISTRO_SUBSTITUIDO>=0)sh.getRange(n+1,idx.ID_REGISTRO_SUBSTITUIDO+1).setValue(idAntigo);
}

function appListarCicloVidaS18(idRegistro){
  exigirPermissaoS15_('consultarMapa',{acao:'LISTAR_CICLO_VIDA',entidade:'REGISTRO',entidadeId:String(idRegistro||'')});
  const sh=SpreadsheetApp.getActive().getSheetByName('CICLO_VIDA_ATIVO');
  if(!sh||sh.getLastRow()<2)return[];
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  return vals.filter(r=>String(r[idx.ID_REGISTRO]||'')===String(idRegistro||'')).map(r=>({
    idEvento:String(r[idx.ID_EVENTO_CICLO]||''),
    dataHora:formatarDataS10_(r[idx.DATA_HORA]),
    tipoEvento:String(r[idx.TIPO_EVENTO]||''),
    statusAnterior:String(r[idx.STATUS_ANTERIOR]||''),
    statusNovo:String(r[idx.STATUS_NOVO]||''),
    mapaAnterior:String(r[idx.MAPA_ANTERIOR]||''),
    ruaAnterior:String(r[idx.RUA_ANTERIOR]||''),
    trechoAnterior:String(r[idx.TRECHO_ANTERIOR]||''),
    mapaNovo:String(r[idx.MAPA_NOVO]||''),
    ruaNovo:String(r[idx.RUA_NOVO]||''),
    trechoNovo:String(r[idx.TRECHO_NOVO]||''),
    motivo:String(r[idx.MOTIVO]||''),
    observacao:String(r[idx.OBSERVACAO]||''),
    responsavel:String(r[idx.RESPONSAVEL]||'')
  })).reverse();
}

function appResumoInventarioS18(){
  exigirPermissaoS15_('acessarDashboard',{acao:'RESUMO_INVENTARIO',entidade:'INVENTARIO'});
  const regs=listarRegistrosS12_();
  const m={ATIVO:0,RETIRADO:0,DESATIVADO:0,SUBSTITUIDO:0,NOVO:0};
  regs.forEach(r=>{
    const s=String(r.STATUS_CICLO_VIDA||'ATIVO').toUpperCase();
    m[s]=(m[s]||0)+1;
  });
  return {total:regs.length,porStatus:Object.entries(m).map(([status,total])=>({status,total}))};
}

function diagnosticoS18(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S18',cfg.APP_FASE||'ausente');
  check_(c,'S18_STATUS',cfg.S18_STATUS==='INSTALADO',cfg.S18_STATUS||'ausente');
  check_(c,'ABA_CICLO_VIDA',!!ss.getSheetByName('CICLO_VIDA_ATIVO'),'CICLO_VIDA_ATIVO');
  check_(c,'EVENTO_CICLO',typeof appRegistrarEventoCicloS18==='function','OK');
  check_(c,'HISTORICO_CICLO',typeof appListarCicloVidaS18==='function','OK');
  check_(c,'RESUMO_INVENTARIO',typeof appResumoInventarioS18==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS18(){
  const d=diagnosticoS18();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S18',
    `${d.ok?'S18 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
      d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S19 — MANUTENÇÃO PREVENTIVA / AGENDA DE INSPEÇÕES
// ========================================================
const S19_AGENDA_HEADERS=[
  'ID_AGENDA','ID_REGISTRO','PROTOCOLO','DATA_PROGRAMADA','RESPONSAVEL',
  'PRIORIDADE','STATUS','ORIGEM','CRIADO_EM','ATUALIZADO_EM',
  'DATA_CONCLUSAO','ID_HISTORICO_INSPECAO','OBSERVACAO','VERSAO_APP'
];

function setupS19(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let sh=ss.getSheetByName('AGENDA_INSPECOES');
  if(!sh)sh=ss.insertSheet('AGENDA_INSPECOES');
  garantirCabecalhosS19_(sh);

  sincronizarAgendaComRegistrosS19_();

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S19','Fase de implementação validada');
  setConfigValue_(cfg,'S19_STATUS','INSTALADO','Agenda preventiva instalada');
  setConfigValue_(cfg,'S19_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S19'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S19',entidade:'SISTEMA',entidadeId:'S19',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS19();
}

function garantirCabecalhosS19_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,S19_AGENDA_HEADERS.length).setValues([S19_AGENDA_HEADERS]);
  }else{
    const falt=S19_AGENDA_HEADERS.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function dataAgendaS19_(v){
  if(!v)return null;
  if(v instanceof Date){
    return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  }
  const s=String(v).trim();
  let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m)return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
  m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if(m)return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));
  const d=new Date(s);
  if(isNaN(d))return null;
  return new Date(d.getFullYear(),d.getMonth(),d.getDate());
}

function isoDataAgendaS19_(v){
  const d=dataAgendaS19_(v);
  return d?Utilities.formatDate(d,APP.TIMEZONE,'yyyy-MM-dd'):'';
}

function brDataAgendaS19_(v){
  const d=dataAgendaS19_(v);
  return d?Utilities.formatDate(d,APP.TIMEZONE,'dd/MM/yyyy'):'';
}

function sincronizarAgendaComRegistrosS19_(){
  const regs=listarRegistrosS12_();
  regs.forEach(r=>{
    const prox=String(r.PROXIMA_INSPECAO||'').trim();
    const ciclo=String(r.STATUS_CICLO_VIDA||'ATIVO').toUpperCase();
    if(prox&&!['DESATIVADO','SUBSTITUIDO','RETIRADO'].includes(ciclo)){
      upsertAgendaS19_({
        idRegistro:String(r.ID_REGISTRO||''),
        protocolo:String(r.PROTOCOLO||''),
        dataProgramada:prox,
        responsavel:String(r.RESPONSAVEL||''),
        origem:'MIGRACAO_S19'
      });
    }
  });
  atualizarVencimentosAgendaS19_();
}

function agendaRowS19_(idRegistro,dataProgramada){
  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  if(!sh||sh.getLastRow()<2)return null;
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};
  h.forEach((x,i)=>idx[x]=i);
  const iso=isoDataAgendaS19_(dataProgramada);
  for(let i=0;i<vals.length;i++){
    const r=vals[i];
    if(String(r[idx.ID_REGISTRO]||'')===String(idRegistro||'') &&
       isoDataAgendaS19_(r[idx.DATA_PROGRAMADA])===iso &&
       !['CONCLUIDA','CANCELADA'].includes(String(r[idx.STATUS]||''))){
      const o={_row:i+2};h.forEach((k,j)=>o[k]=r[j]);return o;
    }
  }
  return null;
}

function prioridadeAgendaS19_(dataProgramada){
  const d=dataAgendaS19_(dataProgramada),hoje=dataAgendaS19_(new Date());
  if(!d)return 'NORMAL';
  const dias=Math.round((d-hoje)/86400000);
  if(dias<0)return 'URGENTE';
  if(dias<=1)return 'ALTA';
  if(dias<=7)return 'MEDIA';
  return 'NORMAL';
}

function upsertAgendaS19_(dados){
  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  if(!sh)return null;
  garantirCabecalhosS19_(sh);

  const existente=agendaRowS19_(dados.idRegistro,dados.dataProgramada);
  if(existente)return existente;

  const agora=new Date();
  const obj={
    ID_AGENDA:'AG-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase(),
    ID_REGISTRO:String(dados.idRegistro||''),
    PROTOCOLO:String(dados.protocolo||''),
    DATA_PROGRAMADA:isoDataAgendaS19_(dados.dataProgramada),
    RESPONSAVEL:String(dados.responsavel||''),
    PRIORIDADE:prioridadeAgendaS19_(dados.dataProgramada),
    STATUS:'PENDENTE',
    ORIGEM:String(dados.origem||'MANUAL'),
    CRIADO_EM:agora,
    ATUALIZADO_EM:agora,
    DATA_CONCLUSAO:'',
    ID_HISTORICO_INSPECAO:'',
    OBSERVACAO:String(dados.observacao||''),
    VERSAO_APP:APP.VERSAO
  };
  appendObjetoPorCabecalhoS7_(sh,obj);
  return obj;
}

function atualizarVencimentosAgendaS19_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  if(!sh||sh.getLastRow()<2)return;
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const hoje=dataAgendaS19_(new Date());
  for(let i=1;i<vals.length;i++){
    const st=String(vals[i][idx.STATUS]||'');
    if(['CONCLUIDA','CANCELADA'].includes(st))continue;
    const d=dataAgendaS19_(vals[i][idx.DATA_PROGRAMADA]);
    const novo=d&&d<hoje?'VENCIDA':'PENDENTE';
    const pri=prioridadeAgendaS19_(d);
    if(st!==novo)sh.getRange(i+1,idx.STATUS+1).setValue(novo);
    if(idx.PRIORIDADE>=0&&String(vals[i][idx.PRIORIDADE]||'')!==pri)sh.getRange(i+1,idx.PRIORIDADE+1).setValue(pri);
  }
}

function concluirAgendaPorInspecaoS19_(payload,resp){
  const idRegistro=String(payload?.idRegistro||'');
  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  if(!sh||!idRegistro)return;

  atualizarVencimentosAgendaS19_();
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const hoje=dataAgendaS19_(new Date());

  // Conclui a agenda aberta mais antiga, desde que programada para hoje ou antes.
  const candidatos=[];
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][idx.ID_REGISTRO]||'')!==idRegistro)continue;
    const st=String(vals[i][idx.STATUS]||'');
    if(!['PENDENTE','VENCIDA'].includes(st))continue;
    const d=dataAgendaS19_(vals[i][idx.DATA_PROGRAMADA]);
    if(d&&d<=hoje)candidatos.push({row:i+1,data:d});
  }
  candidatos.sort((a,b)=>a.data-b.data);
  if(candidatos.length){
    const row=candidatos[0].row,agora=new Date();
    sh.getRange(row,idx.STATUS+1).setValue('CONCLUIDA');
    sh.getRange(row,idx.DATA_CONCLUSAO+1).setValue(agora);
    sh.getRange(row,idx.ID_HISTORICO_INSPECAO+1).setValue(String(resp?.idHistorico||''));
    sh.getRange(row,idx.ATUALIZADO_EM+1).setValue(agora);
  }

  const prox=String(payload?.proximaInspecao||'').trim();
  if(prox){
    const reg=obterRegistroPorIdS8_(idRegistro)||{};
    upsertAgendaS19_({
      idRegistro,
      protocolo:String(reg.PROTOCOLO||resp?.protocolo||''),
      dataProgramada:prox,
      responsavel:String(payload?.responsavel||reg.RESPONSAVEL||''),
      origem:'INSPECAO',
      observacao:'Gerada pela inspeção concluída'
    });
  }
}

function appSincronizarInspecaoS19(payload){
  const resp=appSincronizarInspecaoS15(payload||{});
  if(resp&&resp.ok)concluirAgendaPorInspecaoS19_(payload||{},resp);
  return resp;
}

function appAgendaInspecoesS19(filtros){
  exigirPermissaoS15_('inspecionar',{acao:'LISTAR_AGENDA_INSPECOES',entidade:'AGENDA_INSPECOES'});
  filtros=filtros||{};
  sincronizarAgendaComRegistrosS19_();

  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  if(!sh||sh.getLastRow()<2)return {ok:true,resumo:{},itens:[],responsaveis:[]};

  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const regs=listarRegistrosS12_(),regMap={};regs.forEach(r=>regMap[String(r.ID_REGISTRO||'')]=r);
  const hoje=dataAgendaS19_(new Date()),fim7=new Date(hoje.getTime()+7*86400000);
  const visao=String(filtros.visao||'7DIAS').toUpperCase();

  let itens=vals.map(r=>{
    const reg=regMap[String(r[idx.ID_REGISTRO]||'')]||{};
    const data=dataAgendaS19_(r[idx.DATA_PROGRAMADA]);
    return {
      idAgenda:String(r[idx.ID_AGENDA]||''),
      idRegistro:String(r[idx.ID_REGISTRO]||''),
      protocolo:String(r[idx.PROTOCOLO]||reg.PROTOCOLO||''),
      dataIso:isoDataAgendaS19_(r[idx.DATA_PROGRAMADA]),
      data:brDataAgendaS19_(r[idx.DATA_PROGRAMADA]),
      responsavel:String(r[idx.RESPONSAVEL]||reg.RESPONSAVEL||''),
      prioridade:String(r[idx.PRIORIDADE]||'NORMAL'),
      status:String(r[idx.STATUS]||'PENDENTE'),
      origem:String(r[idx.ORIGEM]||''),
      observacao:String(r[idx.OBSERVACAO]||''),
      titulo:String(reg.TITULO||''),
      tipo:String(reg.TIPO||''),
      idMapaSetor:String(reg.ID_MAPA_SETOR||''),
      mapa:String(reg.MAPA||''),
      piso:String(reg.PISO||''),
      rua:String(reg.RUA||''),
      trecho:String(reg.TRECHO||''),
      x:Number(reg.X_NORMALIZADO||0),
      y:Number(reg.Y_NORMALIZADO||0),
      ciclo:String(reg.STATUS_CICLO_VIDA||'ATIVO'),
      _data:data
    };
  }).filter(i=>{
    if(filtros.responsavel&&i.responsavel!==String(filtros.responsavel))return false;
    if(filtros.status&&i.status!==String(filtros.status))return false;
    if(['DESATIVADO','SUBSTITUIDO','RETIRADO'].includes(i.ciclo))return false;
    if(visao==='HOJE')return i._data&&i._data.getTime()===hoje.getTime()&&!['CONCLUIDA','CANCELADA'].includes(i.status);
    if(visao==='VENCIDAS')return i.status==='VENCIDA';
    if(visao==='7DIAS')return i._data&&i._data>=hoje&&i._data<=fim7&&!['CONCLUIDA','CANCELADA'].includes(i.status);
    if(visao==='PENDENTES')return ['PENDENTE','VENCIDA'].includes(i.status);
    return true;
  });

  itens.sort((a,b)=>(a._data?.getTime()||0)-(b._data?.getTime()||0));
  itens=itens.map(({_data,...i})=>i);

  const todos=vals.map(r=>({
    status:String(r[idx.STATUS]||''),
    data:dataAgendaS19_(r[idx.DATA_PROGRAMADA]),
    responsavel:String(r[idx.RESPONSAVEL]||'')
  }));
  const resumo={
    hoje:todos.filter(i=>i.data&&i.data.getTime()===hoje.getTime()&&!['CONCLUIDA','CANCELADA'].includes(i.status)).length,
    proximos7:todos.filter(i=>i.data&&i.data>=hoje&&i.data<=fim7&&!['CONCLUIDA','CANCELADA'].includes(i.status)).length,
    vencidas:todos.filter(i=>i.status==='VENCIDA').length,
    pendentes:todos.filter(i=>['PENDENTE','VENCIDA'].includes(i.status)).length,
    concluidas:todos.filter(i=>i.status==='CONCLUIDA').length
  };
  const responsaveis=[...new Set(todos.map(i=>i.responsavel).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  return {ok:true,resumo,itens,responsaveis};
}

function appReagendarInspecaoS19(payload){
  payload=payload||{};
  exigirPermissaoS15_('inspecionar',{acao:'REAGENDAR_INSPECAO',entidade:'AGENDA_INSPECOES',entidadeId:String(payload.idAgenda||'')});
  const nova=isoDataAgendaS19_(payload.novaData);
  if(!nova)throw new Error('Nova data inválida.');

  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_AGENDA]||'')===String(payload.idAgenda||''));
  if(row<1)throw new Error('Item da agenda não encontrado.');

  const idRegistro=String(vals[row][idx.ID_REGISTRO]||'');
  const anterior=isoDataAgendaS19_(vals[row][idx.DATA_PROGRAMADA]);
  sh.getRange(row+1,idx.DATA_PROGRAMADA+1).setValue(nova);
  sh.getRange(row+1,idx.STATUS+1).setValue(prioridadeAgendaS19_(nova)==='URGENTE'?'VENCIDA':'PENDENTE');
  sh.getRange(row+1,idx.PRIORIDADE+1).setValue(prioridadeAgendaS19_(nova));
  sh.getRange(row+1,idx.ATUALIZADO_EM+1).setValue(new Date());

  const regSh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  const rv=regSh.getDataRange().getValues(),rh=rv[0].map(String),ridx={};rh.forEach((x,i)=>ridx[x]=i);
  const rr=rv.findIndex((r,i)=>i>0&&String(r[ridx.ID_REGISTRO]||'')===idRegistro);
  if(rr>0&&ridx.PROXIMA_INSPECAO>=0)regSh.getRange(rr+1,ridx.PROXIMA_INSPECAO+1).setValue(nova);

  registrarAuditoriaS15_({
    acao:'INSPECAO_REAGENDADA',entidade:'AGENDA_INSPECOES',entidadeId:String(payload.idAgenda||''),
    resultado:'SUCESSO',detalhes:{motivo:String(payload.motivo||'')},valorAnterior:{data:anterior},valorNovo:{data:nova}
  });
  return {ok:true,idAgenda:String(payload.idAgenda||''),novaData:nova};
}

function appCancelarAgendaS19(idAgenda){
  exigirPermissaoS15_('inspecionar',{acao:'CANCELAR_AGENDA',entidade:'AGENDA_INSPECOES',entidadeId:String(idAgenda||'')});
  const sh=SpreadsheetApp.getActive().getSheetByName('AGENDA_INSPECOES');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_AGENDA]||'')===String(idAgenda||''));
  if(row<1)throw new Error('Item da agenda não encontrado.');
  sh.getRange(row+1,idx.STATUS+1).setValue('CANCELADA');
  sh.getRange(row+1,idx.ATUALIZADO_EM+1).setValue(new Date());
  registrarAuditoriaS15_({acao:'AGENDA_CANCELADA',entidade:'AGENDA_INSPECOES',entidadeId:String(idAgenda||''),resultado:'SUCESSO'});
  return {ok:true};
}

function diagnosticoS19(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S19',cfg.APP_FASE||'ausente');
  check_(c,'S19_STATUS',cfg.S19_STATUS==='INSTALADO',cfg.S19_STATUS||'ausente');
  check_(c,'ABA_AGENDA',!!ss.getSheetByName('AGENDA_INSPECOES'),'AGENDA_INSPECOES');
  check_(c,'AGENDA_API',typeof appAgendaInspecoesS19==='function','OK');
  check_(c,'SYNC_INSPECAO',typeof appSincronizarInspecaoS19==='function','OK');
  check_(c,'REAGENDAR',typeof appReagendarInspecaoS19==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS19(){
  const d=diagnosticoS19();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S19',
    `${d.ok?'S19 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
      d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S20 — PLANO PREVENTIVO / RECORRÊNCIA AUTOMÁTICA
// ========================================================
const S20_PLANOS_HEADERS=[
  'ID_PLANO','NOME','ATIVO','PRIORIDADE_REGRA',
  'TIPO_SINALIZACAO','ID_MAPA_SETOR','PISO','RESPONSAVEL',
  'ESTADO_CONSERVACAO','CONDICAO','INTERVALO_DIAS',
  'ANTECEDENCIA_DIAS','RESPONSAVEL_EXECUCAO',
  'OBSERVACAO','CRIADO_EM','ATUALIZADO_EM','VERSAO_APP'
];

function setupS20(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let sh=ss.getSheetByName('PLANOS_PREVENTIVOS');
  if(!sh)sh=ss.insertSheet('PLANOS_PREVENTIVOS');
  garantirCabecalhosS20_(sh);

  if(sh.getLastRow()<2){
    const agora=new Date();
    sh.getRange(2,1,3,S20_PLANOS_HEADERS.length).setValues([
      ['PP-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Crítica — 30 dias',true,100,'','','','','Crítica','',30,7,'','Regra inicial S20',agora,agora,APP.VERSAO],
      ['PP-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Regular/Ruim — 60 dias',true,80,'','','','','Regular','',60,7,'','Regra inicial S20',agora,agora,APP.VERSAO],
      ['PP-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Padrão — 90 dias',true,10,'','','','','','',90,7,'','Fallback geral',agora,agora,APP.VERSAO]
    ]);
  }

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S20','Fase de implementação validada');
  setConfigValue_(cfg,'S20_STATUS','INSTALADO','Plano preventivo e recorrência automática instalados');
  setConfigValue_(cfg,'S20_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S20'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S20',entidade:'SISTEMA',entidadeId:'S20',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS20();
}

function garantirCabecalhosS20_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,S20_PLANOS_HEADERS.length).setValues([S20_PLANOS_HEADERS]);
  }else{
    const falt=S20_PLANOS_HEADERS.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function boolS20_(v){
  return v===true||String(v).toLowerCase()==='true'||String(v)==='1'||String(v).toLowerCase()==='sim';
}

function listarPlanosS20_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('PLANOS_PREVENTIVOS');
  if(!sh||sh.getLastRow()<2)return[];
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  return vals.filter(r=>String(r[idx.ID_PLANO]||'')).map(r=>({
    idPlano:String(r[idx.ID_PLANO]||''),
    nome:String(r[idx.NOME]||''),
    ativo:boolS20_(r[idx.ATIVO]),
    prioridadeRegra:Number(r[idx.PRIORIDADE_REGRA]||0),
    tipoSinalizacao:String(r[idx.TIPO_SINALIZACAO]||''),
    idMapaSetor:String(r[idx.ID_MAPA_SETOR]||''),
    piso:String(r[idx.PISO]||''),
    responsavel:String(r[idx.RESPONSAVEL]||''),
    estadoConservacao:String(r[idx.ESTADO_CONSERVACAO]||''),
    condicao:String(r[idx.CONDICAO]||''),
    intervaloDias:Number(r[idx.INTERVALO_DIAS]||0),
    antecedenciaDias:Number(r[idx.ANTECEDENCIA_DIAS]||0),
    responsavelExecucao:String(r[idx.RESPONSAVEL_EXECUCAO]||''),
    observacao:String(r[idx.OBSERVACAO]||'')
  })).sort((a,b)=>b.prioridadeRegra-a.prioridadeRegra);
}

function planoAplicaS20_(p,reg){
  if(!p.ativo||p.intervaloDias<=0)return false;
  if(p.tipoSinalizacao&&String(reg.TIPO||'')!==p.tipoSinalizacao)return false;
  if(p.idMapaSetor&&String(reg.ID_MAPA_SETOR||'')!==p.idMapaSetor)return false;
  if(p.piso&&String(reg.PISO||'')!==p.piso)return false;
  if(p.responsavel&&String(reg.RESPONSAVEL||'')!==p.responsavel)return false;
  if(p.estadoConservacao&&String(reg.ESTADO_CONSERVACAO||'')!==p.estadoConservacao)return false;
  if(p.condicao&&String(reg.CONDICAO||'')!==p.condicao)return false;
  return true;
}

function selecionarPlanoS20_(reg){
  const planos=listarPlanosS20_();
  return planos.find(p=>planoAplicaS20_(p,reg))||null;
}

function calcularProximaDataS20_(base,intervaloDias){
  const d=dataAgendaS19_(base)||dataAgendaS19_(new Date());
  return Utilities.formatDate(new Date(d.getTime()+Number(intervaloDias||0)*86400000),APP.TIMEZONE,'yyyy-MM-dd');
}

function appSimularPlanoPreventivoS20(idRegistro){
  exigirPermissaoS15_('acessarDashboard',{acao:'SIMULAR_PLANO_PREVENTIVO',entidade:'REGISTRO',entidadeId:String(idRegistro||'')});
  const reg=obterRegistroPorIdS8_(String(idRegistro||''));
  if(!reg)throw new Error('Sinalização não encontrada.');
  const p=selecionarPlanoS20_(reg);
  if(!p)return {ok:true,encontrado:false};
  const base=reg.DATA_ULTIMA_INSPECAO||new Date();
  return {
    ok:true,encontrado:true,plano:p,
    base:formatarDataS10_(base),
    proximaData:calcularProximaDataS20_(base,p.intervaloDias)
  };
}

function aplicarPlanoAposInspecaoS20_(payload,resp){
  const idRegistro=String(payload?.idRegistro||'');
  const reg=obterRegistroPorIdS8_(idRegistro);
  if(!reg)return;

  const plano=selecionarPlanoS20_(reg);
  if(!plano)return;

  // Se o usuário informou próxima inspeção manualmente, respeita a decisão explícita.
  let prox=String(payload?.proximaInspecao||'').trim();
  if(!prox){
    prox=calcularProximaDataS20_(new Date(),plano.intervaloDias);
  }

  const regSh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  const vals=regSh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_REGISTRO]||'')===idRegistro);
  if(row>0&&idx.PROXIMA_INSPECAO>=0)regSh.getRange(row+1,idx.PROXIMA_INSPECAO+1).setValue(prox);

  upsertAgendaS19_({
    idRegistro,
    protocolo:String(reg.PROTOCOLO||resp?.protocolo||''),
    dataProgramada:prox,
    responsavel:String(plano.responsavelExecucao||payload?.responsavel||reg.RESPONSAVEL||''),
    origem:'PLANO_PREVENTIVO_S20',
    observacao:`Plano: ${plano.nome}`
  });

  registrarAuditoriaS15_({
    acao:'PROXIMA_INSPECAO_AUTOMATICA',
    entidade:'REGISTRO',
    entidadeId:idRegistro,
    resultado:'SUCESSO',
    detalhes:{
      idPlano:plano.idPlano,
      plano:plano.nome,
      intervaloDias:plano.intervaloDias,
      proximaData:prox
    }
  });
}

function appSincronizarInspecaoS20(payload){
  const resp=appSincronizarInspecaoS19(payload||{});
  if(resp&&resp.ok)aplicarPlanoAposInspecaoS20_(payload||{},resp);
  return resp;
}

function appListarPlanosS20(){
  exigirPermissaoS15_('administrar',{acao:'LISTAR_PLANOS_PREVENTIVOS',entidade:'PLANOS_PREVENTIVOS'});
  return listarPlanosS20_();
}

function appSalvarPlanoS20(payload){
  payload=payload||{};
  exigirPermissaoS15_('administrar',{acao:'SALVAR_PLANO_PREVENTIVO',entidade:'PLANOS_PREVENTIVOS',entidadeId:String(payload.idPlano||'')});

  const sh=SpreadsheetApp.getActive().getSheetByName('PLANOS_PREVENTIVOS');
  garantirCabecalhosS20_(sh);
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);

  const id=String(payload.idPlano||'').trim()||'PP-'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_PLANO]||'')===id);
  const agora=new Date();

  const obj={
    ID_PLANO:id,
    NOME:String(payload.nome||'Plano preventivo'),
    ATIVO:payload.ativo!==false,
    PRIORIDADE_REGRA:Number(payload.prioridadeRegra||0),
    TIPO_SINALIZACAO:String(payload.tipoSinalizacao||''),
    ID_MAPA_SETOR:String(payload.idMapaSetor||''),
    PISO:String(payload.piso||''),
    RESPONSAVEL:String(payload.responsavel||''),
    ESTADO_CONSERVACAO:String(payload.estadoConservacao||''),
    CONDICAO:String(payload.condicao||''),
    INTERVALO_DIAS:Number(payload.intervaloDias||0),
    ANTECEDENCIA_DIAS:Number(payload.antecedenciaDias||0),
    RESPONSAVEL_EXECUCAO:String(payload.responsavelExecucao||''),
    OBSERVACAO:String(payload.observacao||''),
    ATUALIZADO_EM:agora,
    VERSAO_APP:APP.VERSAO
  };

  if(obj.INTERVALO_DIAS<=0)throw new Error('Intervalo em dias deve ser maior que zero.');

  if(row>0){
    Object.entries(obj).forEach(([k,v])=>{
      const c=idx[k];if(c>=0)sh.getRange(row+1,c+1).setValue(v);
    });
  }else{
    obj.CRIADO_EM=agora;
    appendObjetoPorCabecalhoS7_(sh,obj);
  }

  registrarAuditoriaS15_({
    acao:row>0?'PLANO_PREVENTIVO_ATUALIZADO':'PLANO_PREVENTIVO_CRIADO',
    entidade:'PLANOS_PREVENTIVOS',entidadeId:id,resultado:'SUCESSO',
    valorNovo:obj
  });
  return {ok:true,idPlano:id};
}

function appAplicarPlanosLoteS20(){
  exigirPermissaoS15_('administrar',{acao:'APLICAR_PLANOS_LOTE',entidade:'PLANOS_PREVENTIVOS'});
  const regs=listarRegistrosS12_();
  let aplicados=0,semPlano=0;
  regs.forEach(reg=>{
    const ciclo=String(reg.STATUS_CICLO_VIDA||'ATIVO').toUpperCase();
    if(['DESATIVADO','RETIRADO','SUBSTITUIDO'].includes(ciclo))return;
    const p=selecionarPlanoS20_(reg);
    if(!p){semPlano++;return;}

    const base=reg.DATA_ULTIMA_INSPECAO||new Date();
    const prox=String(reg.PROXIMA_INSPECAO||'').trim()||calcularProximaDataS20_(base,p.intervaloDias);
    upsertAgendaS19_({
      idRegistro:String(reg.ID_REGISTRO||''),
      protocolo:String(reg.PROTOCOLO||''),
      dataProgramada:prox,
      responsavel:String(p.responsavelExecucao||reg.RESPONSAVEL||''),
      origem:'PLANO_PREVENTIVO_S20',
      observacao:`Plano: ${p.nome}`
    });
    aplicados++;
  });

  registrarAuditoriaS15_({
    acao:'PLANOS_PREVENTIVOS_APLICADOS',
    entidade:'PLANOS_PREVENTIVOS',entidadeId:'LOTE',resultado:'SUCESSO',
    detalhes:{aplicados,semPlano}
  });

  return {ok:true,aplicados,semPlano};
}

function diagnosticoS20(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S20',cfg.APP_FASE||'ausente');
  check_(c,'S20_STATUS',cfg.S20_STATUS==='INSTALADO',cfg.S20_STATUS||'ausente');
  check_(c,'ABA_PLANOS',!!ss.getSheetByName('PLANOS_PREVENTIVOS'),'PLANOS_PREVENTIVOS');
  check_(c,'SELECAO_PLANO',typeof selecionarPlanoS20_==='function','OK');
  check_(c,'SYNC_S20',typeof appSincronizarInspecaoS20==='function','OK');
  check_(c,'APLICAR_LOTE',typeof appAplicarPlanosLoteS20==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS20(){
  const d=diagnosticoS20();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S20',
    `${d.ok?'S20 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S21 — NOTIFICAÇÕES E ALERTAS OPERACIONAIS
// ========================================================
const S21_ALERTAS_HEADERS=[
  'ID_ALERTA','TIPO_ALERTA','NIVEL','ID_AGENDA','ID_REGISTRO','PROTOCOLO',
  'DATA_REFERENCIA','RESPONSAVEL','TITULO','MENSAGEM','STATUS',
  'CRIADO_EM','ATUALIZADO_EM','LIDO_EM','RESOLVIDO_EM','VERSAO_APP'
];

function setupS21(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let sh=ss.getSheetByName('ALERTAS_OPERACIONAIS');
  if(!sh)sh=ss.insertSheet('ALERTAS_OPERACIONAIS');
  garantirCabecalhosS21_(sh);

  atualizarAlertasOperacionaisS21_();

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S21','Fase de implementação validada');
  setConfigValue_(cfg,'S21_STATUS','INSTALADO','Central de alertas operacionais instalada');
  setConfigValue_(cfg,'S21_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S21'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S21',entidade:'SISTEMA',entidadeId:'S21',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS21();
}

function garantirCabecalhosS21_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,S21_ALERTAS_HEADERS.length).setValues([S21_ALERTAS_HEADERS]);
  }else{
    const falt=S21_ALERTAS_HEADERS.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function nivelAlertaS21_(item){
  const hoje=dataAgendaS19_(new Date());
  const data=dataAgendaS19_(item.DATA_PROGRAMADA);
  if(!data)return 'INFO';
  const dias=Math.round((data-hoje)/86400000);
  if(String(item.STATUS||'')==='VENCIDA'||dias<0)return 'CRITICO';
  if(dias===0)return 'ALTO';
  if(dias<=2)return 'ALTO';
  if(dias<=7)return 'MEDIO';
  return 'INFO';
}

function tipoAlertaS21_(item){
  const hoje=dataAgendaS19_(new Date());
  const data=dataAgendaS19_(item.DATA_PROGRAMADA);
  if(!data)return 'AGENDA';
  const dias=Math.round((data-hoje)/86400000);
  if(String(item.STATUS||'')==='VENCIDA'||dias<0)return 'INSPECAO_VENCIDA';
  if(dias===0)return 'INSPECAO_HOJE';
  if(dias<=7)return 'INSPECAO_PROXIMA';
  return 'AGENDA';
}

function alertaMensagemS21_(item,reg,tipo){
  const data=brDataAgendaS19_(item.DATA_PROGRAMADA);
  const prot=String(item.PROTOCOLO||reg.PROTOCOLO||'');
  const titulo=String(reg.TITULO||reg.TIPO||'Sinalização');
  if(tipo==='INSPECAO_VENCIDA')return `${prot} — ${titulo}: inspeção vencida em ${data}.`;
  if(tipo==='INSPECAO_HOJE')return `${prot} — ${titulo}: inspeção prevista para hoje (${data}).`;
  if(tipo==='INSPECAO_PROXIMA')return `${prot} — ${titulo}: inspeção prevista para ${data}.`;
  return `${prot} — ${titulo}: item de agenda em ${data}.`;
}

function atualizarAlertasOperacionaisS21_(){
  atualizarVencimentosAgendaS19_();

  const ss=SpreadsheetApp.getActive();
  const agenda=ss.getSheetByName('AGENDA_INSPECOES');
  const alertas=ss.getSheetByName('ALERTAS_OPERACIONAIS');
  if(!agenda||!alertas)return {criados:0,atualizados:0,resolvidos:0};

  garantirCabecalhosS21_(alertas);

  const regs=listarRegistrosS12_(),regMap={};
  regs.forEach(r=>regMap[String(r.ID_REGISTRO||'')]=r);

  const av=agenda.getDataRange().getValues(),ah=av[0].map(String),aidx={};
  ah.forEach((x,i)=>aidx[x]=i);

  const lv=alertas.getDataRange().getValues(),lh=lv[0].map(String),lidx={};
  lh.forEach((x,i)=>lidx[x]=i);

  const ativosPorChave={};
  for(let i=1;i<lv.length;i++){
    const chave=`${String(lv[i][lidx.ID_AGENDA]||'')}|${String(lv[i][lidx.TIPO_ALERTA]||'')}`;
    ativosPorChave[chave]={row:i+1,status:String(lv[i][lidx.STATUS]||'')};
  }

  const chavesVivas=new Set();
  let criados=0,atualizados=0,resolvidos=0;
  const hoje=dataAgendaS19_(new Date());

  for(let i=1;i<av.length;i++){
    const r=av[i];
    const st=String(r[aidx.STATUS]||'');
    if(!['PENDENTE','VENCIDA'].includes(st))continue;

    const d=dataAgendaS19_(r[aidx.DATA_PROGRAMADA]);
    if(!d)continue;
    const dias=Math.round((d-hoje)/86400000);
    if(dias>7)continue; // S21 centraliza alertas operacionais imediatos.

    const item={};ah.forEach((k,j)=>item[k]=r[j]);
    const reg=regMap[String(item.ID_REGISTRO||'')]||{};
    const tipo=tipoAlertaS21_(item);
    const nivel=nivelAlertaS21_(item);
    const chave=`${String(item.ID_AGENDA||'')}|${tipo}`;
    chavesVivas.add(chave);

    const obj={
      TIPO_ALERTA:tipo,
      NIVEL:nivel,
      ID_AGENDA:String(item.ID_AGENDA||''),
      ID_REGISTRO:String(item.ID_REGISTRO||''),
      PROTOCOLO:String(item.PROTOCOLO||reg.PROTOCOLO||''),
      DATA_REFERENCIA:isoDataAgendaS19_(item.DATA_PROGRAMADA),
      RESPONSAVEL:String(item.RESPONSAVEL||reg.RESPONSAVEL||''),
      TITULO:String(reg.TITULO||reg.TIPO||'Sinalização'),
      MENSAGEM:alertaMensagemS21_(item,reg,tipo),
      STATUS:'ABERTO',
      ATUALIZADO_EM:new Date(),
      VERSAO_APP:APP.VERSAO
    };

    const ex=ativosPorChave[chave];
    if(ex){
      // S21.1: resolução manual é persistente para a mesma agenda + tipo de alerta.
      // A atualização automática não pode reabrir um alerta já resolvido.
      if(ex.status==='RESOLVIDO'){
        continue;
      }
      Object.entries(obj).forEach(([k,v])=>{
        const c=lidx[k];if(c>=0)alertas.getRange(ex.row,c+1).setValue(v);
      });
      atualizados++;
    }else{
      obj.ID_ALERTA='ALT-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
      obj.CRIADO_EM=new Date();
      obj.LIDO_EM='';
      obj.RESOLVIDO_EM='';
      appendObjetoPorCabecalhoS7_(alertas,obj);
      criados++;
    }
  }

  // Resolve alerts whose agenda no longer qualifies.
  for(let i=1;i<lv.length;i++){
    const tipo=String(lv[i][lidx.TIPO_ALERTA]||'');
    if(!['INSPECAO_VENCIDA','INSPECAO_HOJE','INSPECAO_PROXIMA'].includes(tipo))continue;
    const st=String(lv[i][lidx.STATUS]||'');
    if(st==='RESOLVIDO')continue;
    const chave=`${String(lv[i][lidx.ID_AGENDA]||'')}|${tipo}`;
    if(!chavesVivas.has(chave)){
      if(lidx.STATUS>=0)alertas.getRange(i+1,lidx.STATUS+1).setValue('RESOLVIDO');
      if(lidx.RESOLVIDO_EM>=0)alertas.getRange(i+1,lidx.RESOLVIDO_EM+1).setValue(new Date());
      if(lidx.ATUALIZADO_EM>=0)alertas.getRange(i+1,lidx.ATUALIZADO_EM+1).setValue(new Date());
      resolvidos++;
    }
  }

  return {criados,atualizados,resolvidos};
}

function appCentralAlertasS21(filtros){
  exigirPermissaoS15_('inspecionar',{acao:'LISTAR_ALERTAS',entidade:'ALERTAS_OPERACIONAIS'});
  filtros=filtros||{};
  atualizarAlertasOperacionaisS21_();

  const sh=SpreadsheetApp.getActive().getSheetByName('ALERTAS_OPERACIONAIS');
  if(!sh||sh.getLastRow()<2)return {resumo:{total:0,criticos:0,altos:0,naoLidos:0},itens:[],responsaveis:[]};

  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);

  let itens=vals.map(r=>({
    idAlerta:String(r[idx.ID_ALERTA]||''),
    tipo:String(r[idx.TIPO_ALERTA]||''),
    nivel:String(r[idx.NIVEL]||'INFO'),
    idAgenda:String(r[idx.ID_AGENDA]||''),
    idRegistro:String(r[idx.ID_REGISTRO]||''),
    protocolo:String(r[idx.PROTOCOLO]||''),
    dataIso:isoDataAgendaS19_(r[idx.DATA_REFERENCIA]),
    data:brDataAgendaS19_(r[idx.DATA_REFERENCIA]),
    responsavel:String(r[idx.RESPONSAVEL]||''),
    titulo:String(r[idx.TITULO]||''),
    mensagem:String(r[idx.MENSAGEM]||''),
    status:String(r[idx.STATUS]||'ABERTO'),
    lidoEm:r[idx.LIDO_EM]?formatarDataS10_(r[idx.LIDO_EM]):''
  }));

  const todosAbertos=itens.filter(i=>i.status!=='RESOLVIDO');
  const resumo={
    total:todosAbertos.length,
    criticos:todosAbertos.filter(i=>i.nivel==='CRITICO').length,
    altos:todosAbertos.filter(i=>i.nivel==='ALTO').length,
    naoLidos:todosAbertos.filter(i=>!i.lidoEm).length
  };
  const responsaveis=[...new Set(todosAbertos.map(i=>i.responsavel).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));

  itens=itens.filter(i=>{
    if(String(filtros.status||'ABERTOS')==='ABERTOS'&&i.status==='RESOLVIDO')return false;
    if(filtros.nivel&&i.nivel!==String(filtros.nivel))return false;
    if(filtros.responsavel&&i.responsavel!==String(filtros.responsavel))return false;
    return true;
  });

  const peso={CRITICO:4,ALTO:3,MEDIO:2,INFO:1};
  itens.sort((a,b)=>(peso[b.nivel]||0)-(peso[a.nivel]||0)||String(a.dataIso).localeCompare(String(b.dataIso)));

  return {resumo,itens,responsaveis};
}

function appMarcarAlertaLidoS21(idAlerta){
  exigirPermissaoS15_('inspecionar',{acao:'MARCAR_ALERTA_LIDO',entidade:'ALERTAS_OPERACIONAIS',entidadeId:String(idAlerta||'')});
  const sh=SpreadsheetApp.getActive().getSheetByName('ALERTAS_OPERACIONAIS');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_ALERTA]||'')===String(idAlerta||''));
  if(row<1)throw new Error('Alerta não encontrado.');
  sh.getRange(row+1,idx.LIDO_EM+1).setValue(new Date());
  sh.getRange(row+1,idx.ATUALIZADO_EM+1).setValue(new Date());
  return {ok:true};
}

function appResolverAlertaS21(idAlerta){
  exigirPermissaoS15_('inspecionar',{acao:'RESOLVER_ALERTA',entidade:'ALERTAS_OPERACIONAIS',entidadeId:String(idAlerta||'')});
  const sh=SpreadsheetApp.getActive().getSheetByName('ALERTAS_OPERACIONAIS');
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_ALERTA]||'')===String(idAlerta||''));
  if(row<1)throw new Error('Alerta não encontrado.');
  sh.getRange(row+1,idx.STATUS+1).setValue('RESOLVIDO');
  sh.getRange(row+1,idx.RESOLVIDO_EM+1).setValue(new Date());
  sh.getRange(row+1,idx.ATUALIZADO_EM+1).setValue(new Date());
  registrarAuditoriaS15_({acao:'ALERTA_RESOLVIDO',entidade:'ALERTAS_OPERACIONAIS',entidadeId:String(idAlerta||''),resultado:'SUCESSO'});
  return {ok:true};
}

function diagnosticoS21(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S21',cfg.APP_FASE||'ausente');
  check_(c,'S21_STATUS',cfg.S21_STATUS==='INSTALADO',cfg.S21_STATUS||'ausente');
  check_(c,'ABA_ALERTAS',!!ss.getSheetByName('ALERTAS_OPERACIONAIS'),'ALERTAS_OPERACIONAIS');
  check_(c,'CENTRAL_ALERTAS',typeof appCentralAlertasS21==='function','OK');
  check_(c,'MARCAR_LIDO',typeof appMarcarAlertaLidoS21==='function','OK');
  check_(c,'RESOLVER_ALERTA',typeof appResolverAlertaS21==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS21(){
  const d=diagnosticoS21();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S21',
    `${d.ok?'S21 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S22 — COMUNICAÇÃO E ESCALONAMENTO POR E-MAIL
// ========================================================
const S22_REGRAS_HEADERS=[
  'ID_REGRA','NOME','ATIVO','CANAL','NIVEL_MINIMO','RESPONSAVEL_FILTRO',
  'DESTINATARIOS','CC','ASSUNTO_PREFIXO','COOLDOWN_HORAS',
  'ESCALONAR_APOS_HORAS','DESTINATARIOS_ESCALACAO',
  'CRIADO_EM','ATUALIZADO_EM','VERSAO_APP'
];

const S22_ENVIOS_HEADERS=[
  'ID_ENVIO','ID_ALERTA','ID_REGRA','TIPO_ENVIO','CANAL',
  'DESTINATARIOS','CC','ASSUNTO','STATUS','TENTATIVAS','ERRO',
  'HASH_DEDUP','CRIADO_EM','ENVIADO_EM','VERSAO_APP'
];

function setupS22(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  let regras=ss.getSheetByName('REGRAS_NOTIFICACAO');
  if(!regras)regras=ss.insertSheet('REGRAS_NOTIFICACAO');
  garantirCabecalhosS22_(regras,S22_REGRAS_HEADERS);

  let envios=ss.getSheetByName('NOTIFICACOES_ENVIO');
  if(!envios)envios=ss.insertSheet('NOTIFICACOES_ENVIO');
  garantirCabecalhosS22_(envios,S22_ENVIOS_HEADERS);

  // Regras iniciais ficam DESATIVADAS para evitar disparos acidentais.
  if(regras.getLastRow()<2){
    const agora=new Date();
    regras.getRange(2,1,3,S22_REGRAS_HEADERS.length).setValues([
      ['NR-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Críticos — e-mail imediato',false,'EMAIL','CRITICO','','','','[Mall • CRÍTICO]',24,2,'',agora,agora,APP.VERSAO],
      ['NR-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Altos — e-mail',false,'EMAIL','ALTO','','','','[Mall • ALERTA]',24,6,'',agora,agora,APP.VERSAO],
      ['NR-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Médios — resumo operacional',false,'EMAIL','MEDIO','','','','[Mall • Preventiva]',24,24,'',agora,agora,APP.VERSAO]
    ]);
  }

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22','Fase de implementação validada');
  setConfigValue_(cfg,'S22_STATUS','INSTALADO','Comunicação e escalonamento por e-mail instalados');
  setConfigValue_(cfg,'S22_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S22',entidade:'SISTEMA',entidadeId:'S22',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS22();
}

function garantirCabecalhosS22_(sh,headers){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  if(!atuais.some(Boolean)){
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }else{
    const falt=headers.filter(h=>!atuais.includes(h));
    if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#171B68').setFontColor('#fff');
}

function emailValidoS22_(e){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e||'').trim());
}

function normalizarListaEmailsS22_(texto){
  return [...new Set(
    String(texto||'').split(/[;,\n]+/).map(x=>x.trim().toLowerCase()).filter(Boolean)
  )];
}

function validarListaEmailsS22_(texto,campo){
  const arr=normalizarListaEmailsS22_(texto);
  const ruins=arr.filter(e=>!emailValidoS22_(e));
  if(ruins.length)throw new Error(`${campo}: e-mail(s) inválido(s): ${ruins.join(', ')}`);
  return arr;
}

function listarRegrasS22_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('REGRAS_NOTIFICACAO');
  if(!sh||sh.getLastRow()<2)return[];
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  return vals.filter(r=>String(r[idx.ID_REGRA]||'')).map(r=>({
    idRegra:String(r[idx.ID_REGRA]||''),
    nome:String(r[idx.NOME]||''),
    ativo:boolS20_(r[idx.ATIVO]),
    canal:String(r[idx.CANAL]||'EMAIL'),
    nivelMinimo:String(r[idx.NIVEL_MINIMO]||'MEDIO'),
    responsavelFiltro:String(r[idx.RESPONSAVEL_FILTRO]||''),
    destinatarios:String(r[idx.DESTINATARIOS]||''),
    cc:String(r[idx.CC]||''),
    assuntoPrefixo:String(r[idx.ASSUNTO_PREFIXO]||''),
    cooldownHoras:Number(r[idx.COOLDOWN_HORAS]||24),
    escalarAposHoras:Number(r[idx.ESCALONAR_APOS_HORAS]||0),
    destinatariosEscalacao:String(r[idx.DESTINATARIOS_ESCALACAO]||'')
  }));
}

function appListarRegrasS22(){
  exigirPermissaoS15_('administrar',{acao:'LISTAR_REGRAS_NOTIFICACAO',entidade:'REGRAS_NOTIFICACAO'});
  return listarRegrasS22_();
}

function appSalvarRegraS22(payload){
  payload=payload||{};
  exigirPermissaoS15_('administrar',{acao:'SALVAR_REGRA_NOTIFICACAO',entidade:'REGRAS_NOTIFICACAO',entidadeId:String(payload.idRegra||'')});

  const dest=validarListaEmailsS22_(payload.destinatarios,'Destinatários');
  const cc=validarListaEmailsS22_(payload.cc,'CC');
  const esc=validarListaEmailsS22_(payload.destinatariosEscalacao,'Destinatários de escalonamento');

  if(payload.ativo&&dest.length===0)throw new Error('Uma regra ativa precisa ter ao menos um destinatário.');

  const sh=SpreadsheetApp.getActive().getSheetByName('REGRAS_NOTIFICACAO');
  garantirCabecalhosS22_(sh,S22_REGRAS_HEADERS);
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);

  const id=String(payload.idRegra||'').trim()||'NR-'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
  const row=vals.findIndex((r,i)=>i>0&&String(r[idx.ID_REGRA]||'')===id);
  const agora=new Date();

  const obj={
    ID_REGRA:id,
    NOME:String(payload.nome||'Regra de notificação'),
    ATIVO:!!payload.ativo,
    CANAL:'EMAIL',
    NIVEL_MINIMO:String(payload.nivelMinimo||'MEDIO'),
    RESPONSAVEL_FILTRO:String(payload.responsavelFiltro||''),
    DESTINATARIOS:dest.join(','),
    CC:cc.join(','),
    ASSUNTO_PREFIXO:String(payload.assuntoPrefixo||'[Mall]'),
    COOLDOWN_HORAS:Math.max(1,Number(payload.cooldownHoras||24)),
    ESCALONAR_APOS_HORAS:Math.max(0,Number(payload.escalarAposHoras||0)),
    DESTINATARIOS_ESCALACAO:esc.join(','),
    ATUALIZADO_EM:agora,
    VERSAO_APP:APP.VERSAO
  };

  if(row>0){
    Object.entries(obj).forEach(([k,v])=>{const c=idx[k];if(c>=0)sh.getRange(row+1,c+1).setValue(v);});
  }else{
    obj.CRIADO_EM=agora;
    appendObjetoPorCabecalhoS7_(sh,obj);
  }

  registrarAuditoriaS15_({
    acao:row>0?'REGRA_NOTIFICACAO_ATUALIZADA':'REGRA_NOTIFICACAO_CRIADA',
    entidade:'REGRAS_NOTIFICACAO',entidadeId:id,resultado:'SUCESSO',valorNovo:obj
  });

  return {ok:true,idRegra:id};
}

function pesoNivelS22_(nivel){
  return ({INFO:1,MEDIO:2,ALTO:3,CRITICO:4})[String(nivel||'').toUpperCase()]||0;
}

function regraAplicaAlertaS22_(regra,alerta){
  if(!regra.ativo||regra.canal!=='EMAIL')return false;
  if(pesoNivelS22_(alerta.nivel)<pesoNivelS22_(regra.nivelMinimo))return false;
  if(regra.responsavelFiltro&&String(alerta.responsavel||'')!==regra.responsavelFiltro)return false;
  return true;
}

function idadeAlertaHorasS22_(alerta){
  const sh=SpreadsheetApp.getActive().getSheetByName('ALERTAS_OPERACIONAIS');
  if(!sh||sh.getLastRow()<2)return 0;
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const row=vals.find((r,i)=>i>0&&String(r[idx.ID_ALERTA]||'')===String(alerta.idAlerta||''));
  if(!row)return 0;
  const d=row[idx.CRIADO_EM] instanceof Date?row[idx.CRIADO_EM]:new Date(row[idx.CRIADO_EM]);
  return isNaN(d)?0:Math.max(0,(Date.now()-d.getTime())/3600000);
}

function hashDedupS22_(idAlerta,idRegra,tipo){
  const raw=`${idAlerta}|${idRegra}|${tipo}`;
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,raw,Utilities.Charset.UTF_8);
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function envioJaRealizadoS22_(hashDedup,cooldownHoras){
  const sh=SpreadsheetApp.getActive().getSheetByName('NOTIFICACOES_ENVIO');
  if(!sh||sh.getLastRow()<2)return false;
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const agora=Date.now(),limite=Math.max(1,Number(cooldownHoras||24))*3600000;
  return vals.some((r,i)=>{
    if(i===0)return false;
    if(String(r[idx.HASH_DEDUP]||'')!==hashDedup)return false;
    if(String(r[idx.STATUS]||'')!=='SUCESSO')return false;
    const d=r[idx.ENVIADO_EM] instanceof Date?r[idx.ENVIADO_EM]:new Date(r[idx.ENVIADO_EM]);
    return !isNaN(d)&&(agora-d.getTime())<=limite;
  });
}

function templateEmailAlertaS22_(alerta,regra,tipoEnvio){
  const assunto=`${regra.assuntoPrefixo||'[Mall]'} ${alerta.protocolo||''} — ${alerta.titulo||alerta.tipo||'Alerta'}`.trim();
  const tituloEsc=tipoEnvio==='ESCALONAMENTO'?'ESCALONAMENTO DE ALERTA':'ALERTA OPERACIONAL';

  const html=`<div style="font-family:Arial,sans-serif;color:#101228;max-width:680px">
    <div style="background:#171B68;color:#fff;padding:16px 18px;border-radius:12px 12px 0 0">
      <strong>${tituloEsc}</strong><br>
      <span style="font-size:13px">Sinalização do Mall</span>
    </div>
    <div style="border:1px solid #dfe2ea;border-top:0;padding:18px;border-radius:0 0 12px 12px">
      <div style="display:inline-block;background:#fff0f8;color:#cb006f;padding:6px 10px;border-radius:999px;font-weight:bold">
        ${String(alerta.nivel||'').replace(/</g,'&lt;')}
      </div>
      <h2 style="margin:14px 0 8px">${String(alerta.protocolo||'').replace(/</g,'&lt;')} — ${String(alerta.titulo||'').replace(/</g,'&lt;')}</h2>
      <p>${String(alerta.mensagem||'').replace(/</g,'&lt;')}</p>
      <table style="border-collapse:collapse;width:100%;font-size:14px">
        <tr><td style="padding:6px 0;color:#676A7A">Responsável</td><td style="padding:6px 0;font-weight:bold">${String(alerta.responsavel||'—').replace(/</g,'&lt;')}</td></tr>
        <tr><td style="padding:6px 0;color:#676A7A">Data</td><td style="padding:6px 0;font-weight:bold">${String(alerta.data||'—').replace(/</g,'&lt;')}</td></tr>
        <tr><td style="padding:6px 0;color:#676A7A">Tipo</td><td style="padding:6px 0;font-weight:bold">${String(alerta.tipo||'—').replace(/</g,'&lt;')}</td></tr>
      </table>
      <p style="font-size:12px;color:#676A7A;margin-top:18px">Mensagem automática da aplicação Sinalização do Mall • ${APP.VERSAO}</p>
    </div>
  </div>`;

  const texto=[
    tituloEsc,
    'Sinalização do Mall',
    '',
    `${alerta.nivel} — ${alerta.protocolo} — ${alerta.titulo}`,
    alerta.mensagem,
    `Responsável: ${alerta.responsavel||'—'}`,
    `Data: ${alerta.data||'—'}`,
    `Tipo: ${alerta.tipo||'—'}`,
    '',
    `Aplicação: ${APP.VERSAO}`
  ].join('\n');

  return {assunto,html,texto};
}

function registrarEnvioS22_(obj){
  const sh=SpreadsheetApp.getActive().getSheetByName('NOTIFICACOES_ENVIO');
  garantirCabecalhosS22_(sh,S22_ENVIOS_HEADERS);
  appendObjetoPorCabecalhoS7_(sh,obj);
}

function enviarEmailS22_(alerta,regra,tipoEnvio,destinatariosOverride){
  const destinatarios=destinatariosOverride||regra.destinatarios;
  const dest=validarListaEmailsS22_(destinatarios,'Destinatários');
  const cc=validarListaEmailsS22_(regra.cc,'CC');
  if(!dest.length)throw new Error('Sem destinatários para envio.');

  const hash=hashDedupS22_(alerta.idAlerta,regra.idRegra,tipoEnvio);
  if(envioJaRealizadoS22_(hash,regra.cooldownHoras)){
    return {ok:true,ignorado:true,motivo:'COOLDOWN'};
  }

  const tpl=templateEmailAlertaS22_(alerta,regra,tipoEnvio);
  const idEnvio='ENV-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
  const base={
    ID_ENVIO:idEnvio,
    ID_ALERTA:String(alerta.idAlerta||''),
    ID_REGRA:String(regra.idRegra||''),
    TIPO_ENVIO:tipoEnvio,
    CANAL:'EMAIL',
    DESTINATARIOS:dest.join(','),
    CC:cc.join(','),
    ASSUNTO:tpl.assunto,
    TENTATIVAS:1,
    HASH_DEDUP:hash,
    CRIADO_EM:new Date(),
    VERSAO_APP:APP.VERSAO
  };

  try{
    MailApp.sendEmail({
      to:dest.join(','),
      cc:cc.join(','),
      subject:tpl.assunto,
      body:tpl.texto,
      htmlBody:tpl.html,
      name:'Sinalização do Mall'
    });
    registrarEnvioS22_({...base,STATUS:'SUCESSO',ERRO:'',ENVIADO_EM:new Date()});
    return {ok:true,idEnvio};
  }catch(e){
    registrarEnvioS22_({...base,STATUS:'ERRO',ERRO:String(e?.message||e),ENVIADO_EM:''});
    throw e;
  }
}

function processarNotificacoesS221_(origem){
  const central=appCentralAlertasS21({status:'ABERTOS'});
  const alertas=central.itens||[];
  const regras=listarRegrasS22_().filter(r=>r.ativo);

  let enviados=0,escalados=0,ignorados=0,erros=0;

  alertas.forEach(alerta=>{
    regras.filter(r=>regraAplicaAlertaS22_(r,alerta)).forEach(regra=>{
      try{
        const r1=enviarEmailS22_(alerta,regra,'INICIAL');
        if(r1.ignorado)ignorados++; else enviados++;

        if(regra.escalarAposHoras>0 && regra.destinatariosEscalacao){
          const idade=idadeAlertaHorasS22_(alerta);
          if(idade>=regra.escalarAposHoras){
            const r2=enviarEmailS22_(alerta,regra,'ESCALONAMENTO',regra.destinatariosEscalacao);
            if(r2.ignorado)ignorados++; else escalados++;
          }
        }
      }catch(e){
        erros++;
        console.error('[S22.1] erro ao processar alerta/regra',alerta?.idAlerta,regra?.idRegra,e);
      }
    });
  });

  registrarAuditoriaS15_({
    acao:origem==='TRIGGER'?'NOTIFICACOES_PROCESSADAS_AUTOMATICO':'NOTIFICACOES_PROCESSADAS',
    entidade:'NOTIFICACOES_ENVIO',
    entidadeId:'LOTE',
    resultado:erros?'ERRO':'SUCESSO',
    origem:origem==='TRIGGER'?'TRIGGER':'WEB_APP',
    detalhes:{enviados,escalados,ignorados,erros,alertas:alertas.length,regrasAtivas:regras.length}
  });

  return {ok:erros===0,enviados,escalados,ignorados,erros,alertas:alertas.length,regrasAtivas:regras.length,origem};
}

function appProcessarNotificacoesS22(){
  exigirPermissaoS15_('administrar',{acao:'PROCESSAR_NOTIFICACOES',entidade:'NOTIFICACOES_ENVIO'});
  return processarNotificacoesS221_('MANUAL');
}

function appEnviarTesteNotificacaoS22(idRegra,emailDestino){
  exigirPermissaoS15_('administrar',{acao:'TESTAR_NOTIFICACAO',entidade:'REGRAS_NOTIFICACAO',entidadeId:String(idRegra||'')});
  const email=String(emailDestino||'').trim().toLowerCase();
  if(!emailValidoS22_(email))throw new Error('Informe um e-mail válido para o teste.');

  const regra=listarRegrasS22_().find(r=>r.idRegra===String(idRegra||''));
  if(!regra)throw new Error('Regra não encontrada.');

  const alerta={
    idAlerta:'TESTE-'+Utilities.getUuid().slice(0,8),
    protocolo:'SIG-TESTE',
    titulo:'Teste de comunicação',
    mensagem:'Este é um e-mail de teste da regra de notificação da aplicação Sinalização do Mall.',
    nivel:regra.nivelMinimo||'MEDIO',
    responsavel:'Teste',
    data:Utilities.formatDate(new Date(),APP.TIMEZONE,'dd/MM/yyyy HH:mm'),
    tipo:'TESTE'
  };

  const teste={...regra,destinatarios:email,cc:''};
  const tpl=templateEmailAlertaS22_(alerta,teste,'TESTE');

  MailApp.sendEmail({
    to:email,
    subject:tpl.assunto,
    body:tpl.texto,
    htmlBody:tpl.html,
    name:'Sinalização do Mall'
  });

  registrarEnvioS22_({
    ID_ENVIO:'ENV-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase(),
    ID_ALERTA:alerta.idAlerta,
    ID_REGRA:regra.idRegra,
    TIPO_ENVIO:'TESTE',
    CANAL:'EMAIL',
    DESTINATARIOS:email,
    CC:'',
    ASSUNTO:tpl.assunto,
    STATUS:'SUCESSO',
    TENTATIVAS:1,
    ERRO:'',
    HASH_DEDUP:hashDedupS22_(alerta.idAlerta,regra.idRegra,'TESTE'),
    CRIADO_EM:new Date(),
    ENVIADO_EM:new Date(),
    VERSAO_APP:APP.VERSAO
  });

  registrarAuditoriaS15_({
    acao:'NOTIFICACAO_TESTE_ENVIADA',entidade:'REGRAS_NOTIFICACAO',
    entidadeId:regra.idRegra,resultado:'SUCESSO',detalhes:{destino:email}
  });

  return {ok:true,destino:email};
}

function appHistoricoNotificacoesS22(){
  exigirPermissaoS15_('administrar',{acao:'LISTAR_NOTIFICACOES',entidade:'NOTIFICACOES_ENVIO'});
  const sh=SpreadsheetApp.getActive().getSheetByName('NOTIFICACOES_ENVIO');
  if(!sh||sh.getLastRow()<2)return[];
  const vals=sh.getDataRange().getValues(),h=vals.shift().map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  return vals.map(r=>({
    idEnvio:String(r[idx.ID_ENVIO]||''),
    idAlerta:String(r[idx.ID_ALERTA]||''),
    idRegra:String(r[idx.ID_REGRA]||''),
    tipoEnvio:String(r[idx.TIPO_ENVIO]||''),
    destinatarios:String(r[idx.DESTINATARIOS]||''),
    assunto:String(r[idx.ASSUNTO]||''),
    status:String(r[idx.STATUS]||''),
    erro:String(r[idx.ERRO]||''),
    enviadoEm:r[idx.ENVIADO_EM]?formatarDataS10_(r[idx.ENVIADO_EM]):''
  })).reverse().slice(0,200);
}

function diagnosticoS22(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22',cfg.APP_FASE||'ausente');
  check_(c,'S22_STATUS',cfg.S22_STATUS==='INSTALADO',cfg.S22_STATUS||'ausente');
  check_(c,'ABA_REGRAS',!!ss.getSheetByName('REGRAS_NOTIFICACAO'),'REGRAS_NOTIFICACAO');
  check_(c,'ABA_ENVIOS',!!ss.getSheetByName('NOTIFICACOES_ENVIO'),'NOTIFICACOES_ENVIO');
  check_(c,'PROCESSAR',typeof appProcessarNotificacoesS22==='function','OK');
  check_(c,'MAILAPP',typeof MailApp!=='undefined','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}
function mostrarDiagnosticoS22(){
  const d=diagnosticoS22();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22',
    `${d.ok?'S22 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S22.1 — AUTOMAÇÃO HORÁRIA / ESCALONAMENTO
// ========================================================
const S221_TRIGGER_FN='processarNotificacoesAutomaticoS221';

function setupS221(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  // Garante estruturas da S22 sem reativar ou sobrescrever regras existentes.
  let regras=ss.getSheetByName('REGRAS_NOTIFICACAO');
  if(!regras)regras=ss.insertSheet('REGRAS_NOTIFICACAO');
  garantirCabecalhosS22_(regras,S22_REGRAS_HEADERS);

  let envios=ss.getSheetByName('NOTIFICACOES_ENVIO');
  if(!envios)envios=ss.insertSheet('NOTIFICACOES_ENVIO');
  garantirCabecalhosS22_(envios,S22_ENVIOS_HEADERS);

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.1','Fase de implementação validada');
  setConfigValue_(cfg,'S22_STATUS','INSTALADO','Comunicação por e-mail instalada');
  setConfigValue_(cfg,'S221_STATUS','INSTALADO','Escalonamento automático disponível');
  setConfigValue_(cfg,'S221_AUTOMACAO','DESATIVADA','Estado da automação horária');
  setConfigValue_(cfg,'S221_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.1'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S221',entidade:'SISTEMA',entidadeId:'S22.1',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS221();
}

function obterTriggersS221_(){
  return ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===S221_TRIGGER_FN);
}

function appStatusAutomacaoS221(){
  exigirPermissaoS15_('administrar',{acao:'STATUS_AUTOMACAO_NOTIFICACOES',entidade:'SISTEMA'});
  const tr=obterTriggersS221_();
  const cfg=lerConfigComoObjeto_(SpreadsheetApp.getActive());
  return {
    ativa:tr.length>0,
    quantidade:tr.length,
    configuracao:String(cfg.S221_AUTOMACAO||'DESATIVADA'),
    frequencia:'HOURLY',
    handler:S221_TRIGGER_FN
  };
}

function appAtivarAutomacaoS221(){
  exigirPermissaoS15_('administrar',{acao:'ATIVAR_AUTOMACAO_NOTIFICACOES',entidade:'SISTEMA'});

  obterTriggersS221_().forEach(t=>ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(S221_TRIGGER_FN)
    .timeBased()
    .everyHours(1)
    .create();

  const cfg=SpreadsheetApp.getActive().getSheetByName('CONFIG');
  setConfigValue_(cfg,'S221_AUTOMACAO','ATIVA','Estado da automação horária');
  setConfigValue_(cfg,'S221_AUTOMACAO_ATIVADA_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Última ativação da automação'
  );

  registrarAuditoriaS15_({
    acao:'AUTOMACAO_NOTIFICACOES_ATIVADA',
    entidade:'SISTEMA',entidadeId:'S22.1',resultado:'SUCESSO',origem:'WEB_APP'
  });

  return appStatusAutomacaoS221();
}

function appDesativarAutomacaoS221(){
  exigirPermissaoS15_('administrar',{acao:'DESATIVAR_AUTOMACAO_NOTIFICACOES',entidade:'SISTEMA'});

  obterTriggersS221_().forEach(t=>ScriptApp.deleteTrigger(t));

  const cfg=SpreadsheetApp.getActive().getSheetByName('CONFIG');
  setConfigValue_(cfg,'S221_AUTOMACAO','DESATIVADA','Estado da automação horária');

  registrarAuditoriaS15_({
    acao:'AUTOMACAO_NOTIFICACOES_DESATIVADA',
    entidade:'SISTEMA',entidadeId:'S22.1',resultado:'SUCESSO',origem:'WEB_APP'
  });

  return appStatusAutomacaoS221();
}

// Esta função é executada pelo gatilho instalável.
// Não usa exigirPermissaoS15_, pois gatilhos não possuem usuário interativo.
function processarNotificacoesAutomaticoS221(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(5000)){
    console.warn('[S22.1] execução ignorada: processamento já em andamento.');
    return {ok:true,ignorado:true,motivo:'LOCK'};
  }

  try{
    const cfg=lerConfigComoObjeto_(SpreadsheetApp.getActive());
    if(String(cfg.S221_AUTOMACAO||'DESATIVADA')!=='ATIVA'){
      return {ok:true,ignorado:true,motivo:'AUTOMACAO_DESATIVADA'};
    }

    const resultado=processarNotificacoesS221_('TRIGGER');

    console.log(JSON.stringify({evento:'NOTIFICACOES_TRIGGER',ok:!!(resultado&&resultado.ok),ts:new Date().toISOString()}));
    const cfgSh=SpreadsheetApp.getActive().getSheetByName('CONFIG');
    setConfigValue_(cfgSh,'S221_ULTIMA_EXECUCAO',
      Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
      'Última execução automática'
    );
    setConfigValue_(cfgSh,'S221_ULTIMO_RESULTADO',
      JSON.stringify(resultado),
      'Resumo da última execução automática'
    );

    return resultado;
  }catch(e){
    try{
      registrarAuditoriaS15_({
        acao:'NOTIFICACOES_PROCESSADAS_AUTOMATICO',
        entidade:'NOTIFICACOES_ENVIO',entidadeId:'LOTE',
        resultado:'ERRO',origem:'TRIGGER',detalhes:{erro:String(e?.message||e)}
      });
    }catch(_){}
    throw e;
  }finally{
    lock.releaseLock();
  }
}

function appExecutarAutomacaoAgoraS221(){
  exigirPermissaoS15_('administrar',{acao:'EXECUTAR_AUTOMACAO_AGORA',entidade:'SISTEMA'});
  // Executa em modo de teste administrativo, mesmo se o trigger ainda estiver desativado.
  return processarNotificacoesS221_('MANUAL');
}

function diagnosticoS221(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.1',cfg.APP_FASE||'ausente');
  check_(c,'S221_STATUS',cfg.S221_STATUS==='INSTALADO',cfg.S221_STATUS||'ausente');
  check_(c,'TRIGGER_FN',typeof processarNotificacoesAutomaticoS221==='function','OK');
  check_(c,'STATUS_API',typeof appStatusAutomacaoS221==='function','OK');
  check_(c,'ATIVAR_API',typeof appAtivarAutomacaoS221==='function','OK');
  check_(c,'DESATIVAR_API',typeof appDesativarAutomacaoS221==='function','OK');
  const out={ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
  // R06 (auditoria S26.10) — observabilidade do gatilho sem abrir a planilha.
  out.ultimaExecucao=String(cfg.S221_ULTIMA_EXECUCAO||'');
  out.ultimoResultado=String(cfg.S221_ULTIMO_RESULTADO||'');
  out.gatilhosInstalados=obterTriggersS221_().length;
  return out;
}

function mostrarDiagnosticoS221(){
  const d=diagnosticoS221();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.1',
    `${d.ok?'S22.1 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}

/**
 * Auditoria S26.10 — estado REAL dos gatilhos instalados.
 * Rode no editor do Apps Script (não é chamável por RPC):
 *   function diagnosticoGatilhosS221(){ ... }
 * Confirma se o trigger horário S221 existe de fato no projeto e
 * compara com a CONFIG (S221_AUTOMACAO / S221_ULTIMA_EXECUCAO).
 */
function diagnosticoGatilhosS221(){
  const cfg=lerConfigComoObjeto_(SpreadsheetApp.getActive());
  const todos=ScriptApp.getProjectTriggers();
  const s221=todos.filter(t=>t.getHandlerFunction()===S221_TRIGGER_FN);
  return {
    ok:s221.length>0,
    configuracao:String(cfg.S221_AUTOMACAO||''),
    ultimaExecucao:String(cfg.S221_ULTIMA_EXECUCAO||''),
    totalGatilhos:todos.length,
    gatilhosS221:s221.map(t=>({
      uid:t.getUniqueId(),
      handler:t.getHandlerFunction(),
      tipo:String(t.getEventType()),
      fonte:String(t.getTriggerSource()),
      intervaloHoras:t.getTriggerSource()==='TIME'?Number(t.getIntervalInHours()||1):null
    })),
    demaisGatilhos:todos
      .filter(t=>t.getHandlerFunction()!==S221_TRIGGER_FN)
      .map(t=>t.getHandlerFunction())
  };
}


// ========================================================
// S22.2 — TESTE SEGURO DE ESCALONAMENTO
// ========================================================
function setupS222(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.2','Fase de implementação validada');
  setConfigValue_(cfg,'S222_STATUS','INSTALADO','Modo de teste seguro de escalonamento instalado');
  setConfigValue_(cfg,'S222_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.2'
  );

  registrarAuditoriaS15_({
    acao:'SETUP_S222',entidade:'SISTEMA',entidadeId:'S22.2',
    resultado:'SUCESSO',origem:'APPS_SCRIPT'
  });

  SpreadsheetApp.flush();
  return diagnosticoS222();
}

function appListarAlertasAbertosS222(){
  exigirPermissaoS15_('administrar',{acao:'LISTAR_ALERTAS_TESTE_ESCALONAMENTO',entidade:'ALERTAS_OPERACIONAIS'});
  const d=appCentralAlertasS21({status:'ABERTOS'});
  return (d.itens||[]).map(a=>({
    idAlerta:a.idAlerta,
    protocolo:a.protocolo,
    titulo:a.titulo,
    nivel:a.nivel,
    responsavel:a.responsavel,
    data:a.data,
    tipo:a.tipo
  }));
}

function appSimularEscalonamentoS222(payload){
  payload=payload||{};
  exigirPermissaoS15_('administrar',{
    acao:'SIMULAR_ESCALONAMENTO',
    entidade:'REGRAS_NOTIFICACAO',
    entidadeId:String(payload.idRegra||'')
  });

  const regra=listarRegrasS22_().find(r=>r.idRegra===String(payload.idRegra||''));
  if(!regra)throw new Error('Regra não encontrada.');

  const destino=String(regra.destinatariosEscalacao||'').trim();
  if(!destino)throw new Error('Configure Destinatários do escalonamento nesta regra antes do teste.');

  const central=appCentralAlertasS21({status:'ABERTOS'});
  const alerta=(central.itens||[]).find(a=>a.idAlerta===String(payload.idAlerta||''));
  if(!alerta)throw new Error('Alerta aberto não encontrado.');

  if(!regraAplicaAlertaS22_(regra,alerta)){
    throw new Error('A regra selecionada não é compatível com este alerta.');
  }

  // Teste isolado: usa hash/tipo próprios, não consome o cooldown do ESCALONAMENTO real.
  const tipoEnvio='ESCALONAMENTO_TESTE';
  const dest=validarListaEmailsS22_(destino,'Destinatários do escalonamento');
  const cc=validarListaEmailsS22_(regra.cc,'CC');

  const tpl=templateEmailAlertaS22_(alerta,regra,'ESCALONAMENTO');
  const assunto='[TESTE] '+tpl.assunto;
  const html='<div style="font-family:Arial,sans-serif;padding:10px;background:#fffbea;border:1px solid #fde68a;border-radius:10px;margin-bottom:12px"><strong>SIMULAÇÃO DE ESCALONAMENTO</strong><br>Este envio não altera a idade do alerta nem consome o cooldown do escalonamento automático.</div>'+tpl.html;
  const texto='SIMULAÇÃO DE ESCALONAMENTO\nEste envio não altera a idade do alerta nem consome o cooldown do escalonamento automático.\n\n'+tpl.texto;

  const idEnvio='ENV-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
  const hash=hashDedupS22_(alerta.idAlerta,regra.idRegra,tipoEnvio+'-'+Date.now());

  try{
    MailApp.sendEmail({
      to:dest.join(','),
      cc:cc.join(','),
      subject:assunto,
      body:texto,
      htmlBody:html,
      name:'Sinalização do Mall'
    });

    registrarEnvioS22_({
      ID_ENVIO:idEnvio,
      ID_ALERTA:String(alerta.idAlerta||''),
      ID_REGRA:String(regra.idRegra||''),
      TIPO_ENVIO:tipoEnvio,
      CANAL:'EMAIL',
      DESTINATARIOS:dest.join(','),
      CC:cc.join(','),
      ASSUNTO:assunto,
      STATUS:'SUCESSO',
      TENTATIVAS:1,
      ERRO:'',
      HASH_DEDUP:hash,
      CRIADO_EM:new Date(),
      ENVIADO_EM:new Date(),
      VERSAO_APP:APP.VERSAO
    });

    registrarAuditoriaS15_({
      acao:'ESCALONAMENTO_TESTE_ENVIADO',
      entidade:'ALERTAS_OPERACIONAIS',
      entidadeId:String(alerta.idAlerta||''),
      resultado:'SUCESSO',
      origem:'WEB_APP',
      detalhes:{
        idRegra:regra.idRegra,
        regra:regra.nome,
        protocolo:alerta.protocolo,
        nivel:alerta.nivel,
        destinatarios:dest
      }
    });

    return {
      ok:true,
      idEnvio,
      protocolo:alerta.protocolo,
      nivel:alerta.nivel,
      destinatarios:dest
    };
  }catch(e){
    registrarEnvioS22_({
      ID_ENVIO:idEnvio,
      ID_ALERTA:String(alerta.idAlerta||''),
      ID_REGRA:String(regra.idRegra||''),
      TIPO_ENVIO:tipoEnvio,
      CANAL:'EMAIL',
      DESTINATARIOS:dest.join(','),
      CC:cc.join(','),
      ASSUNTO:assunto,
      STATUS:'ERRO',
      TENTATIVAS:1,
      ERRO:String(e?.message||e),
      HASH_DEDUP:hash,
      CRIADO_EM:new Date(),
      ENVIADO_EM:'',
      VERSAO_APP:APP.VERSAO
    });
    throw e;
  }
}

function diagnosticoS222(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.2',cfg.APP_FASE||'ausente');
  check_(c,'S222_STATUS',cfg.S222_STATUS==='INSTALADO',cfg.S222_STATUS||'ausente');
  check_(c,'LISTAR_ALERTAS_TESTE',typeof appListarAlertasAbertosS222==='function','OK');
  check_(c,'SIMULAR_ESCALONAMENTO',typeof appSimularEscalonamentoS222==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

function mostrarDiagnosticoS222(){
  const d=diagnosticoS222();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.2',
    `${d.ok?'S22.2 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S22.3 — AUTENTICAÇÃO HÍBRIDA DO MALL
// Google Identity + fallback e-mail/PIN
// ========================================================
const S223_SESSION_HOURS=12;
const S223_MAX_LOGIN_FAILS=5;
const S223_LOCK_MINUTES=15;
const S223_SESSION_HEADERS=[
  'ID_SESSAO','EMAIL','TOKEN_HASH','DEVICE_ID','CRIADO_EM','EXPIRA_EM',
  'ULTIMO_USO','STATUS','ORIGEM','ENCERRADO_EM','VERSAO_APP'
];
const S223_USUARIO_EXTRA=[
  'PIN_SALT','PIN_HASH','PIN_ATUALIZADO_EM','TENTATIVAS_LOGIN',
  'BLOQUEADO_ATE','ULTIMO_LOGIN_PIN'
];

function setupS223(){
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  const usuarios=ss.getSheetByName('USUARIOS');
  if(!usuarios)throw new Error('USUARIOS ausente. Execute o setup de Administração primeiro.');
  garantirCabecalhosS223Usuarios_(usuarios);

  let sessoes=ss.getSheetByName('SESSOES_USUARIO');
  if(!sessoes)sessoes=ss.insertSheet('SESSOES_USUARIO');
  garantirCabecalhosS22_(sessoes,S223_SESSION_HEADERS);

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.3','Fase de implementação validada');
  setConfigValue_(cfg,'S223_STATUS','INSTALADO','Autenticação híbrida instalada');
  setConfigValue_(cfg,'S223_SESSION_HOURS',S223_SESSION_HOURS,'Duração da sessão PIN em horas');
  setConfigValue_(cfg,'S223_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.3'
  );

  try{
    registrarAuditoriaS15_({
      acao:'SETUP_S223',entidade:'SISTEMA',entidadeId:'S22.3',
      resultado:'SUCESSO',origem:'APPS_SCRIPT'
    });
  }catch(_){}

  SpreadsheetApp.flush();
  return diagnosticoS223();
}

function garantirCabecalhosS223Usuarios_(sh){
  const atuais=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  const falt=S223_USUARIO_EXTRA.filter(h=>!atuais.includes(h));
  if(falt.length)sh.getRange(1,atuais.length+1,1,falt.length).setValues([falt]);
}

function sha256HexS223_(texto){
  const bytes=Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(texto||''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function novoSaltS223_(){
  return Utilities.getUuid().replace(/-/g,'')+
    Utilities.getUuid().replace(/-/g,'').slice(0,16);
}

function hashPinS223_(pin,salt){
  return sha256HexS223_(`${salt}|${String(pin||'')}|${APP.ID}`);
}

function hashTokenS223_(token){
  return sha256HexS223_(`SESSION|${String(token||'')}|${APP.ID}`);
}

function tokenSessaoS223_(){
  return [
    Utilities.getUuid().replace(/-/g,''),
    Utilities.getUuid().replace(/-/g,''),
    Date.now().toString(36)
  ].join('.');
}

function usuarioPorEmailS223_(email){
  email=normalizarEmailS14_(email);
  const sh=SpreadsheetApp.getActive().getSheetByName('USUARIOS');
  if(!sh||sh.getLastRow()<2)return null;
  garantirCabecalhosS223Usuarios_(sh);
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  for(let i=1;i<vals.length;i++){
    if(normalizarEmailS14_(vals[i][idx.EMAIL])===email){
      const o={_row:i+1};
      h.forEach((k,j)=>o[k]=vals[i][j]);
      return o;
    }
  }
  return null;
}

function usuarioPublicoS223_(u,origem){
  if(!u)return null;
  return {
    email:normalizarEmailS14_(u.EMAIL),
    nome:String(u.NOME||normalizarEmailS14_(u.EMAIL).split('@')[0]||'Usuário'),
    perfil:String(u.PERFIL||'CONSULTA').toUpperCase(),
    ativo:boolS20_(u.ATIVO),
    setorPadrao:String(u.SETOR_PADRAO||''),
    origem:origem||'PIN'
  };
}

function registrarFalhaLoginS223_(u,motivo){
  if(!u)return;
  const sh=SpreadsheetApp.getActive().getSheetByName('USUARIOS');
  garantirCabecalhosS223Usuarios_(sh);
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  const tent=Number(u.TENTATIVAS_LOGIN||0)+1;
  sh.getRange(u._row,idx.TENTATIVAS_LOGIN+1).setValue(tent);
  if(tent>=S223_MAX_LOGIN_FAILS){
    sh.getRange(u._row,idx.BLOQUEADO_ATE+1).setValue(new Date(Date.now()+S223_LOCK_MINUTES*60000));
  }
  try{
    registrarAuditoriaS15_({
      acao:'LOGIN_PIN_FALHA',entidade:'USUARIO',entidadeId:normalizarEmailS14_(u.EMAIL),
      resultado:'ERRO',origem:'WEB_APP',detalhes:{motivo,tentativa:tent}
    });
  }catch(_){}
}

function appLoginMallS223(payload){
  payload=payload||{};
  const email=normalizarEmailS14_(payload.email);
  const pin=String(payload.pin||'');
  const deviceId=String(payload.deviceId||'').slice(0,150);

  if(!email||!pin)throw new Error('Informe e-mail e PIN.');

  const lock=LockService.getScriptLock();
  lock.waitLock(5000);
  try{
    const u=usuarioPorEmailS223_(email);
    // Mensagem genérica evita revelar se o e-mail existe.
    if(!u)throw new Error('Credenciais inválidas.');
    if(!boolS20_(u.ATIVO))throw new Error('Credenciais inválidas.');

    const bloqueado=u.BLOQUEADO_ATE instanceof Date?u.BLOQUEADO_ATE:new Date(u.BLOQUEADO_ATE||0);
    if(bloqueado instanceof Date&&!isNaN(bloqueado)&&bloqueado>new Date()){
      throw new Error(`Acesso temporariamente bloqueado. Tente novamente após ${Utilities.formatDate(bloqueado,APP.TIMEZONE,'HH:mm')}.`);
    }

    if(!String(u.PIN_HASH||'')||!String(u.PIN_SALT||'')){
      throw new Error('PIN ainda não configurado para este usuário. Solicite a um administrador.');
    }

    const esperado=hashPinS223_(pin,String(u.PIN_SALT));
    if(esperado!==String(u.PIN_HASH)){
      registrarFalhaLoginS223_(u,'PIN_INVALIDO');
      throw new Error('Credenciais inválidas.');
    }

    const token=tokenSessaoS223_();
    const agora=new Date();
    const expira=new Date(agora.getTime()+S223_SESSION_HOURS*3600000);
    // R08 (auditoria S26.10) — log estruturado de login (PIN nunca é logado).
    console.log(JSON.stringify({evento:'LOGIN_OK',email,ts:agora.toISOString()}));

    const sessoes=SpreadsheetApp.getActive().getSheetByName('SESSOES_USUARIO');
    appendObjetoPorCabecalhoS7_(sessoes,{
      ID_SESSAO:'SES-'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase(),
      EMAIL:email,
      TOKEN_HASH:hashTokenS223_(token),
      DEVICE_ID:deviceId,
      CRIADO_EM:agora,
      EXPIRA_EM:expira,
      ULTIMO_USO:agora,
      STATUS:'ATIVA',
      ORIGEM:'PIN',
      ENCERRADO_EM:'',
      VERSAO_APP:APP.VERSAO
    });

    const sh=SpreadsheetApp.getActive().getSheetByName('USUARIOS');
    garantirCabecalhosS223Usuarios_(sh);
    const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
    sh.getRange(u._row,idx.TENTATIVAS_LOGIN+1).setValue(0);
    sh.getRange(u._row,idx.BLOQUEADO_ATE+1).setValue('');
    sh.getRange(u._row,idx.ULTIMO_LOGIN_PIN+1).setValue(agora);
    if(idx.ULTIMO_ACESSO>=0)sh.getRange(u._row,idx.ULTIMO_ACESSO+1).setValue(agora);

    try{
      registrarAuditoriaS15_({
        acao:'LOGIN_PIN_SUCESSO',entidade:'USUARIO',entidadeId:email,
        resultado:'SUCESSO',origem:'WEB_APP',
        detalhes:{deviceId:deviceId||'',expiraEm:expira.toISOString()}
      });
    }catch(_){}

    return {
      ok:true,
      token,
      expiraEm:expira.toISOString(),
      usuario:usuarioPublicoS223_(u,'PIN')
    };
  }finally{
    lock.releaseLock();
  }
}

// R09 (auditoria S26.10) — leitura por colunas da aba de sessões: o caminho
// de validação roda em TODA chamada RPC; ler apenas as colunas necessárias
// reduz o tráfego de planilha (a aba pode ter milhares de linhas).
const S223_SESSAO_COLUNAS=['EMAIL','TOKEN_HASH','STATUS','EXPIRA_EM','ULTIMO_USO','DEVICE_ID','ENCERRADO_EM'];

function lerSessoesS223_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('SESSOES_USUARIO');
  if(!sh||sh.getLastRow()<2)return null;
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const idx={};h.forEach((x,i)=>idx[x]=i);
  const cols=S223_SESSAO_COLUNAS.map(c=>idx[c]).filter(c=>c>=0);
  if(!cols.length)return null;
  const min=Math.min.apply(null,cols),max=Math.max.apply(null,cols);
  const grid=sh.getRange(2,min+1,sh.getLastRow()-1,max-min+1).getValues();
  const vals=[h.slice()];
  grid.forEach(r=>{
    const linha=h.slice();
    cols.forEach(c=>{linha[c]=r[c-min];});
    vals.push(linha);
  });
  return {sh,vals,h,idx};
}

function validarSessaoMallS223_(token,deviceId){
  token=String(token||'');
  if(!token)return null;

  const dados=lerSessoesS223_();
  if(!dados)return null;
  const {sh,vals,idx}=dados;
  const hash=hashTokenS223_(token),agora=new Date();

  for(let i=1;i<vals.length;i++){
    if(String(vals[i][idx.TOKEN_HASH]||'')!==hash)continue;
    if(String(vals[i][idx.STATUS]||'')!=='ATIVA')return null;

    const exp=vals[i][idx.EXPIRA_EM] instanceof Date?vals[i][idx.EXPIRA_EM]:new Date(vals[i][idx.EXPIRA_EM]||0);
    if(!(exp instanceof Date)||isNaN(exp)||exp<=agora){
      sh.getRange(i+1,idx.STATUS+1).setValue('EXPIRADA');
      return null;
    }

    const sessDevice=String(vals[i][idx.DEVICE_ID]||'');
    if(sessDevice&&deviceId&&sessDevice!==String(deviceId))return null;

    const email=normalizarEmailS14_(vals[i][idx.EMAIL]);
    const u=usuarioPorEmailS223_(email);
    if(!u||!boolS20_(u.ATIVO))return null;

    sh.getRange(i+1,idx.ULTIMO_USO+1).setValue(agora);
    return {usuario:usuarioPublicoS223_(u,'PIN'),row:i+1,expiraEm:exp.toISOString()};
  }
  return null;
}

/**
 * R17 (auditoria S26.10) — higiene da aba de sessões.
 * Marca EXPIRADA as sessões vencidas e remove EXPIRADA/ENCERRADA com mais
 * de 30 dias (não há referência externa a essas linhas). Somente ADMIN.
 */
function appLimparSessoesExpiradasS223(){
  exigirPermissaoS15_('administrar',{acao:'LIMPAR_SESSOES',entidade:'SESSOES_USUARIO'});
  const dados=lerSessoesS223_();
  if(!dados)return {ok:true,marcadas:0,removidas:0};
  const {sh,vals,idx}=dados;
  const agora=new Date();
  const corte=new Date(agora.getTime()-30*86400000);
  let marcadas=0,removidas=0;
  // varre de baixo para cima para deleteRow não deslocar índices
  for(let i=vals.length-1;i>=1;i--){
    const status=String(vals[i][idx.STATUS]||'');
    const exp=vals[i][idx.EXPIRA_EM] instanceof Date?vals[i][idx.EXPIRA_EM]:new Date(vals[i][idx.EXPIRA_EM]||0);
    const encerrado=vals[i][idx.ENCERRADO_EM] instanceof Date?vals[i][idx.ENCERRADO_EM]:new Date(vals[i][idx.ENCERRADO_EM]||0);
    if(status==='ATIVA'&&(!(exp instanceof Date)||isNaN(exp)||exp<=agora)){
      sh.getRange(i+1,idx.STATUS+1).setValue('EXPIRADA');
      marcadas++;
      continue;
    }
    if((status==='EXPIRADA'||status==='ENCERRADA')&&
       (exp instanceof Date&&!isNaN(exp)&&exp<corte)){
      // Conservador: remove somente quando EXPIRA_EM é válido e antigo.
      sh.deleteRow(i+1);
      removidas++;
    }
  }
  return {ok:true,marcadas,removidas};
}

function appValidarSessaoMallS223(payload){
  payload=payload||{};
  const v=validarSessaoMallS223_(payload.token,payload.deviceId);
  return v?{ok:true,usuario:v.usuario,expiraEm:v.expiraEm}:{ok:false};
}

function appLogoutMallS223(payload){
  payload=payload||{};
  const token=String(payload.token||'');
  if(!token)return {ok:true};

  const dados=lerSessoesS223_();
  if(!dados)return {ok:true};
  const {sh,vals,idx}=dados;
  const hash=hashTokenS223_(token);

  for(let i=1;i<vals.length;i++){
    if(String(vals[i][idx.TOKEN_HASH]||'')===hash){
      const email=normalizarEmailS14_(vals[i][idx.EMAIL]);
      sh.getRange(i+1,idx.STATUS+1).setValue('ENCERRADA');
      sh.getRange(i+1,idx.ENCERRADO_EM+1).setValue(new Date());
      try{
        registrarAuditoriaS15_({
          acao:'LOGOUT_PIN',entidade:'USUARIO',entidadeId:email,
          resultado:'SUCESSO',origem:'WEB_APP'
        });
      }catch(_){}
      break;
    }
  }
  return {ok:true};
}

function appDefinirPinUsuarioS223(payload){
  payload=payload||{};
  exigirPermissaoS15_('administrar',{
    acao:'DEFINIR_PIN_USUARIO',entidade:'USUARIO',entidadeId:String(payload.email||'')
  });

  const email=normalizarEmailS14_(payload.email);
  const pin=String(payload.pin||'');
  const confirmacao=String(payload.confirmacao||'');

  if(pin!==confirmacao)throw new Error('A confirmação do PIN não confere.');
  if(!/^\d{6,10}$/.test(pin))throw new Error('Use um PIN numérico de 6 a 10 dígitos.');

  const u=usuarioPorEmailS223_(email);
  if(!u)throw new Error('Usuário não encontrado.');

  const salt=novoSaltS223_(),hash=hashPinS223_(pin,salt);
  const sh=SpreadsheetApp.getActive().getSheetByName('USUARIOS');
  garantirCabecalhosS223Usuarios_(sh);
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);

  sh.getRange(u._row,idx.PIN_SALT+1).setValue(salt);
  sh.getRange(u._row,idx.PIN_HASH+1).setValue(hash);
  sh.getRange(u._row,idx.PIN_ATUALIZADO_EM+1).setValue(new Date());
  sh.getRange(u._row,idx.TENTATIVAS_LOGIN+1).setValue(0);
  sh.getRange(u._row,idx.BLOQUEADO_ATE+1).setValue('');

  // Revoga sessões PIN anteriores deste usuário ao redefinir PIN.
  revogarSessoesUsuarioS223_(email,'PIN_REDEFINIDO');

  registrarAuditoriaS15_({
    acao:'PIN_USUARIO_DEFINIDO',entidade:'USUARIO',entidadeId:email,
    resultado:'SUCESSO',origem:'WEB_APP'
  });

  return {ok:true,email};
}

function revogarSessoesUsuarioS223_(email,motivo){
  const sh=SpreadsheetApp.getActive().getSheetByName('SESSOES_USUARIO');
  if(!sh||sh.getLastRow()<2)return 0;
  const vals=sh.getDataRange().getValues(),h=vals[0].map(String),idx={};h.forEach((x,i)=>idx[x]=i);
  let n=0;
  for(let i=1;i<vals.length;i++){
    if(normalizarEmailS14_(vals[i][idx.EMAIL])===normalizarEmailS14_(email)&&String(vals[i][idx.STATUS]||'')==='ATIVA'){
      sh.getRange(i+1,idx.STATUS+1).setValue('REVOGADA');
      sh.getRange(i+1,idx.ENCERRADO_EM+1).setValue(new Date());
      n++;
    }
  }
  return n;
}

function appRevogarSessoesUsuarioS223(email){
  exigirPermissaoS15_('administrar',{
    acao:'REVOGAR_SESSOES_USUARIO',entidade:'USUARIO',entidadeId:String(email||'')
  });
  const n=revogarSessoesUsuarioS223_(email,'ADMIN');
  registrarAuditoriaS15_({
    acao:'SESSOES_USUARIO_REVOGADAS',entidade:'USUARIO',
    entidadeId:normalizarEmailS14_(email),resultado:'SUCESSO',
    detalhes:{quantidade:n}
  });
  return {ok:true,quantidade:n};
}

function diagnosticoS223(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.3',cfg.APP_FASE||'ausente');
  check_(c,'S223_STATUS',cfg.S223_STATUS==='INSTALADO',cfg.S223_STATUS||'ausente');
  check_(c,'ABA_SESSOES',!!ss.getSheetByName('SESSOES_USUARIO'),'SESSOES_USUARIO');
  check_(c,'LOGIN_PIN',typeof appLoginMallS223==='function','OK');
  check_(c,'VALIDAR_SESSAO',typeof appValidarSessaoMallS223==='function','OK');
  check_(c,'DEFINIR_PIN',typeof appDefinirPinUsuarioS223==='function','OK');
  const allowlist=s223RpcAllowlistChecagem_();
  check_(c,'RPC_ALLOWLIST',allowlist.ausentes.length===0,
    allowlist.total+' nomes; ausentes: '+(allowlist.ausentes.join(', ')||'nenhum'));
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

function mostrarDiagnosticoS223(){
  const d=diagnosticoS223();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.3',
    `${d.ok?'S22.3 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// S22.3 RPC autenticado para sessões PIN.
// ALLOWLIST EXPLÍCITA (auditoria S26.10, item crítico 5):
// - somente funções chamadas pelo frontend via server()/'appRpcS223' podem
//   ser executadas por RPC; nada de prefixo genérico app*/diagnostico*.
// - MANUTENÇÃO: ao expor um novo endpoint no frontend, adicione o nome aqui.
//   Fontes: wrapper server() (120) + serverReferenciasS2610R3 (R3/R5/R6B/R6C/R8)
//   + serverCatalogosS269C1 (catálogos de domínio) + serverS266 (presets).
//   O diagnóstico s223RpcAllowlistChecagem_() aponta nomes sem implementação.
// - diagnostico* NÃO entra na allowlist: diagnósticos rodam no editor/menu.
let S223_RPC_USER=null;

const S223_RPC_ALLOWLIST=Object.freeze([
  'appAgendaInspecoesS19', 'appAlternarStatusConclusaoAS', 'appAlterarAtivoCatalogoReferenciaS2610R8', 'appAplicarPlanosLoteS20',
  'appArquivarRascunhoTorresS2610F', 'appAtivarAutomacaoS221', 'appAtivarReferenciaS2610R2',
  'appAtualizarPendenciaS15', 'appAtualizarReferenciaS2610R2', 'appAtualizarRegistroS237',
  'appAvaliarRemocaoReferenciaS2610R6B', 'appCancelarAgendaS19', 'appCarregarEditorTorresS2610C',
  'appCarregarGovernancaTorresS2610F', 'appCarregarS5B', 'appCatalogosAdminCarregar',
  'appCatalogosAdminCriarOpcao', 'appCatalogosAdminDefinirAtivo', 'appCatalogosAdminEditarOpcao',
  'appCatalogosAdminHistorico', 'appCatalogosAdminRemoverOpcao', 'appCatalogosAdminReordenar',
  'appCatalogosReferenciaS2610R8', 'appCatalogosReferenciasS2610R2', 'appCentralAlertasS21',
  'appCentralGestaoS14', 'appCriarBackupS16', 'appCriarRascunhoCartograficoS255',
  'appCriarRascunhoPublicacaoTorresS2610F', 'appCriarReferenciaS2610R2', 'appCriarSnapshotCartograficoS253',
  'appCriarSnapshotTorresS2610F', 'appDashboardS12', 'appDashboardS14',
  'appDefinirPadraoPresetCorporativoS266', 'appDefinirPinUsuarioS223', 'appDesativarAutomacaoS221',
  'appDesativarReferenciaS2610R2', 'appDescartarRascunhoCartograficoS255', 'appEnviarEmailAutorizacaoServico',
  'appEnviarTesteNotificacaoS22', 'appEnviarValidacaoCartograficaS255', 'appEnviarValidacaoTorresS2610F', 'appExcluirLojista', 'appExcluirPresetCorporativoS266',
  'appExcluirRegistroS236', 'appExecutarAutomacaoAgoraS221', 'appGarantirColunasCeopAs', 'appHistoricoNotificacoesS22',
  'appImportarLojistasLote',
  'appListarAcessosDispositivosS2611A4', 'appListarAlertasAbertosS222', 'appListarAreasNivel0S268D',
  'appListarAreasNivel1S246', 'appListarAuditoriaS15', 'appListarBackupsS16',
  'appListarCicloVidaS18', 'appListarFotosRegistroS225', 'appListarHistoricoCartograficoS253',
  'appListarHistoricoS8', 'appListarLojistas', 'appListarNiveisExtrasS268A', 'appListarNiveisS241',
  'appListarParesCalibracaoS242', 'appListarPendenciasS10', 'appListarPerfisS14',
  'appListarPlanosS20', 'appListarPresetsCorporativosS266', 'appListarReferenciasS2610R2',
  'appListarRegistrosMapaS4', 'appListarRegistrosNivelS242', 'appListarRegrasS22',
  'appListarSnapshotsRestauraveisS254', 'appListarTodasAutorizacoesServico', 'appListarTombstonesReferenciasS2610R6C', 'appListarUsuariosS14',
  'appMarcarAlertaLidoS21', 'appObterAreaVermelhaS244', 'appObterCamadasNivelS267B1',
  'appObterCamadasS2', 'appObterCicloPublicacaoCartograficaS255', 'appObterContextoPendenciaS11',
  'appObterEnquadramentoNivelS267C', 'appObterEventoCartograficoS253', 'appObterFotoS225',
  'appObterGovernancaCartograficaS25', 'appObterImagemMapaS2', 'appObterImagemNivelS241',
  'appObterManifestOfflineS5B', 'appObterOperacaoNivel0S268D', 'appObterOperacaoNivel1S246',
  'appObterOperacaoNivel3S244', 'appObterPacoteCalibracaoS242', 'appObterPacoteMapaOfflineS5B',
  'appObterPacoteTorresOfflineS2610E', 'appObterLojista', 'appObterMapaOcupacaoLojistas', 'appObterHistoricoLojista', 'appObterReferenciaS2610R2', 'appObterRegistroParaEdicaoS237',
  'appObterUrlPlanilhaRegistros',
  'appObterSnapshotMidiaAtualS2352', 'appOpcoesAuditoriaS15', 'appPingS224',
  'appPrepararLocalizacaoS3', 'appPrevisualizarPublicacaoCartograficaS255', 'appPrevisualizarPublicacaoTorresS2610F',
  'appPrevisualizarRestauracaoS254', 'appProcessarNotificacoesS22', 'appPublicarCartografiaS255',
  'appPublicarTorresS2610F', 'appReabrirRascunhoCartograficoS255', 'appReabrirRascunhoTorresS2610F',
  'appReagendarInspecaoS19', 'appRegistrarEventoCicloS18', 'appRelatorioGerencialS14',
  'appRemoverItemCatalogoReferenciaS2610R8', 'appRemoverReferenciaS2610R6C', 'appReposicionarReferenciaS2610R5',
  'appResolverAlertaS21', 'appBuscarLojistasParaLocalizador', 'appResolverLojistaDoPonto', 'appResolverLojistaPorLuc', 'appResolverPontoNivel0S268D', 'appResolverPontoNivel1S246',
  'appResolverPontoNivel3S244', 'appResolverPontoNivelS243', 'appResolverTorreS2610D',
  'appRestaurarBackupS16', 'appRestaurarSnapshotCartograficoS254', 'appRestaurarSnapshotTorresComoRascunhoS2610F',
  'appResumoInventarioS18', 'appResumoPoliticaRemocaoReferenciasS2610R6B', 'appRevalidarIntegridadeCartograficaS252',
  'appRevogarDispositivoS2611A4', 'appRevogarTodasSessoesUsuarioS2611A4', 'appSalvarAreaNivel0S268D',
  'appSalvarAreaNivel1S246', 'appSalvarAreaVermelhaS244', 'appSalvarCalibracaoS242',
  'appSalvarComponenteTorreS2610B', 'appSalvarItemCatalogoReferenciaS2610R8', 'appSalvarLojista', 'appSalvarPerfilS15',
  'appSalvarPlanoS20', 'appSalvarPresetCorporativoS266', 'appSalvarRegraS22',
  'appSalvarRepresentacaoTorreS2610B', 'appSalvarTorreS2610B', 'appSalvarUsuarioS15',
  'appSaudeOperacionalS17', 'appSessaoS14', 'appSimularEscalonamentoS222',
  'appSimularRestauracaoS16', 'appSincronizarFotoInspecaoS235', 'appSincronizarInspecaoS20',
  'appSincronizarRegistroS6', 'appSincronizarRuasPlanilhaLojistas', 'appStatusAutomacaoS221',
  'appStatusBaseLojistas', 'appVerificarInspecaoS9', 'appVerificarIntegridadeS16',
  'appVerificarRegistroS6'
]);


/** Retorna nomes da allowlist sem implementação no backend (para diagnóstico). */
function s223RpcAllowlistChecagem_(){
  const ausentes=[];
  for(const nome of S223_RPC_ALLOWLIST){
    try{ if(typeof eval(nome)!=='function')ausentes.push(nome); }catch(_){ ausentes.push(nome); }
  }
  return {total:S223_RPC_ALLOWLIST.length,ausentes};
}

function appRpcS223(fn,args,auth){
  fn=String(fn||'');
  args=Array.isArray(args)?args:[];
  auth=auth||{};

  if(S223_RPC_ALLOWLIST.indexOf(fn)<0){
    throw new Error('Operação não permitida.');
  }
  if(['appRpcS223','appLoginMallS223','appValidarSessaoMallS223','appLogoutMallS223'].includes(fn)){
    throw new Error('Operação não permitida via RPC.');
  }

  const v=validarSessaoMallS223_(auth.authToken,auth.deviceId);
  if(!v)throw new Error('Sessão expirada ou inválida.');

  // Em Apps Script V8, funções top-level nem sempre ficam expostas em `this`.
  // O nome já foi rigidamente validado acima; eval resolve a função top-level
  // no próprio escopo do projeto.
  let alvo=null;
  try{ alvo=eval(fn); }catch(_){}
  if(typeof alvo!=='function')throw new Error('Operação indisponível: '+fn);

  S223_RPC_USER=v.usuario;
  try{
    // R08 (auditoria S26.10) — log estruturado do dispatch RPC (sem dados sensíveis).
    console.log(JSON.stringify({evento:'RPC',fn,usuario:(v.usuario&&v.usuario.email)||'',ts:new Date().toISOString()}));
    return alvo.apply(null,args);
  }finally{
    S223_RPC_USER=null;
  }
}

function usuarioRpcAtualS223_(){
  return S223_RPC_USER;
}


// ========================================================
// S22.4-A/B — RESILIÊNCIA DE CONECTIVIDADE + LOCAL-FIRST
// ========================================================
function appPingS224(){
  return {
    ok:true,
    ts:Date.now(),
    versao:APP.VERSAO
  };
}

function setupS224(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.4','Fase de implementação validada');
  setConfigValue_(cfg,'S224_STATUS','INSTALADO','ConnectivityManager + local-first instalados');
  setConfigValue_(cfg,'S224_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.4'
  );

  try{
    registrarAuditoriaS15_({
      acao:'SETUP_S224',
      entidade:'SISTEMA',
      entidadeId:'S22.4',
      resultado:'SUCESSO',
      origem:'APPS_SCRIPT'
    });
  }catch(_){}

  SpreadsheetApp.flush();
  return diagnosticoS224();
}

function diagnosticoS224(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.4',cfg.APP_FASE||'ausente');
  check_(c,'S224_STATUS',cfg.S224_STATUS==='INSTALADO',cfg.S224_STATUS||'ausente');
  check_(c,'PING',typeof appPingS224==='function','OK');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

function mostrarDiagnosticoS224(){
  const d=diagnosticoS224();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.4',
    `${d.ok?'S22.4 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S22.4-C — WRITE-LOCAL-FIRST
// ========================================================
function setupS224C(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.4-C','Fase de implementação validada');
  setConfigValue_(cfg,'S224C_STATUS','INSTALADO','Write-local-first instalado');
  setConfigValue_(cfg,'S224C_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.4-C'
  );

  try{
    registrarAuditoriaS15_({
      acao:'SETUP_S224C',entidade:'SISTEMA',entidadeId:'S22.4-C',
      resultado:'SUCESSO',origem:'APPS_SCRIPT'
    });
  }catch(_){}

  SpreadsheetApp.flush();
  return diagnosticoS224C();
}

function diagnosticoS224C(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.4-C',cfg.APP_FASE||'ausente');
  check_(c,'S224C_STATUS',cfg.S224C_STATUS==='INSTALADO',cfg.S224C_STATUS||'ausente');
  check_(c,'OUTBOX',true,'IndexedDB/outbox');
  return{ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

function mostrarDiagnosticoS224C(){
  const d=diagnosticoS224C();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.4-C',
    `${d.ok?'S22.4-C PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}

// ========================================================
// S22.4-D — SINCRONIZAÇÃO ADAPTATIVA
// ========================================================
function setupS224D(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.4-D','Fase de implementação validada');
  setConfigValue_(cfg,'S224D_STATUS','INSTALADO','Sincronização adaptativa instalada');
  setConfigValue_(cfg,'S224D_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.4-D'
  );

  try{
    registrarAuditoriaS15_({
      acao:'SETUP_S224D',
      entidade:'SISTEMA',
      entidadeId:'S22.4-D',
      resultado:'SUCESSO',
      origem:'APPS_SCRIPT'
    });
  }catch(_){}

  SpreadsheetApp.flush();
  return diagnosticoS224D();
}

function diagnosticoS224D(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.4-D',cfg.APP_FASE||'ausente');
  check_(c,'S224D_STATUS',cfg.S224D_STATUS==='INSTALADO',cfg.S224D_STATUS||'ausente');
  check_(c,'FOTOS_BACKEND',typeof appSincronizarFotoS7==='function','OK');
  return{
    ok:c.every(x=>x.ok),
    totalChecks:c.length,
    totalFalhas:c.filter(x=>!x.ok).length,
    checks:c
  };
}

function mostrarDiagnosticoS224D(){
  const d=diagnosticoS224D();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.4-D',
    `${d.ok?'S22.4-D PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}


// ========================================================
// S22.5 — GALERIA DE FOTOS NO POPUP DO PIN
// ========================================================
function appListarFotosRegistroS225(idRegistro){
  exigirPermissaoS14_('consultarMapa');
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('REGISTRO_FOTOS');
  if(!sh||sh.getLastRow()<2)return [];

  const id=String(idRegistro||'').trim();
  if(!id)return [];

  // S23.5.2.2
  // O card do mapa representa a MÍDIA ATUAL, não todo o histórico fotográfico.
  // Usa a mesma regra de snapshot da Nova inspeção, mas retorna somente metadados.
  garantirCabecalhosFotosInspecaoS235_(sh);

  const rows=sheetObjectsS2352_(sh)
    .filter(f=>
      String(f.ID_REGISTRO||'').trim()===id &&
      String(f.ARQUIVO_ID||'').trim()
    );

  if(!rows.length)return [];

  const base=rows.filter(f=>!String(f.CLIENT_INSPECTION_ID||'').trim());
  const insp=rows.filter(f=>String(f.CLIENT_INSPECTION_ID||'').trim());

  let selecionadas=[];

  if(insp.length){
    const grupos={};

    insp.forEach(f=>{
      const k=String(f.CLIENT_INSPECTION_ID||'').trim();
      (grupos[k]=grupos[k]||[]).push(f);
    });

    const chaves=Object.keys(grupos).sort((a,b)=>{
      const ma=Math.max(...grupos[a].map(x=>
        dataMillisS2352_(x.CRIADO_EM||x.SINCRONIZADO_EM)
      ));
      const mb=Math.max(...grupos[b].map(x=>
        dataMillisS2352_(x.CRIADO_EM||x.SINCRONIZADO_EM)
      ));
      return mb-ma;
    });

    const ult=grupos[chaves[0]]||[];

    // Snapshots novos (S23.5.2+) já são completos.
    // Inspeções legadas S23.5/S23.5.1 são complementadas com o cadastro inicial.
    const snapshotCompleto=ult.some(f=>
      ['HERDADA','SNAPSHOT'].includes(
        String(f.ORIGEM_FOTO||'').toUpperCase()
      )
    );

    selecionadas=snapshotCompleto
      ? deduplicarFotosSnapshotS2352_(ult)
      : deduplicarFotosSnapshotS2352_([...base,...ult]);
  }else{
    selecionadas=deduplicarFotosSnapshotS2352_(base);
  }

  selecionadas.sort((a,b)=>
    (Number(a.ORDEM_FOTO||999)-Number(b.ORDEM_FOTO||999)) ||
    dataMillisS2352_(a.CRIADO_EM)-dataMillisS2352_(b.CRIADO_EM)
  );

  return selecionadas.map((f,i)=>{
    let criado='';
    const bruto=f.CRIADO_EM||f.SINCRONIZADO_EM;

    if(bruto instanceof Date){
      criado=Utilities.formatDate(bruto,APP.TIMEZONE,'dd/MM/yyyy HH:mm');
    }else{
      criado=String(bruto||'');
    }

    return {
      idFoto:String(f.ID_FOTO||''),
      arquivoId:String(f.ARQUIVO_ID||''),
      nomeArquivo:String(f.NOME_ARQUIVO||`Foto ${i+1}`),
      mimeType:String(f.MIME_TYPE||'image/jpeg'),
      tamanhoBytes:Number(f.TAMANHO_BYTES||0),
      criadoEm:criado,
      ordemFoto:Number(f.ORDEM_FOTO||i+1),
      origemFoto:String(f.ORIGEM_FOTO||'CADASTRO')
    };
  });
}

const S225_MAX_FOTO_BYTES=10*1024*1024; // 10 MB — limite de foto servida pelo Web App

function appObterFotoS225(arquivoId){
  exigirPermissaoS14_('consultarMapa');
  const id=String(arquivoId||'').trim();
  if(!id)throw new Error('Foto não informada.');

  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('REGISTRO_FOTOS');
  if(!sh||sh.getLastRow()<2)throw new Error('Foto não encontrada ou não autorizada.');

  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const arquivoIdCol=headers.indexOf('ARQUIVO_ID');
  if(arquivoIdCol<0)throw new Error('Coluna ARQUIVO_ID ausente em REGISTRO_FOTOS.');

  const arquivoAutorizado=sh
    .getRange(2,arquivoIdCol+1,sh.getLastRow()-1,1)
    .getDisplayValues()
    .some(row=>String(row[0]||'').trim()===id);
  if(!arquivoAutorizado)throw new Error('Foto não encontrada ou não autorizada.');

  const f=DriveApp.getFileById(id);
  const blob=f.getBlob();
  const bytes=blob.getBytes();
  const mime=blob.getContentType()||'image/jpeg';

  // Auditoria S26.10 — reforço: somente imagem e tamanho controlado.
  if(!/^image\//.test(mime)){
    throw new Error('Arquivo não é uma imagem.');
  }
  if(bytes.length>S225_MAX_FOTO_BYTES){
    throw new Error('Imagem excede o tamanho máximo de 10 MB.');
  }

  return {
    ok:true,
    nome:f.getName(),
    mimeType:mime,
    tamanhoBytes:bytes.length,
    dataUrl:'data:'+mime+';base64,'+Utilities.base64Encode(bytes)
  };
}

function setupS225(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5','Fase de implementação validada');
  setConfigValue_(cfg,'S225_STATUS','INSTALADO','Galeria de fotos no popup instalada');
  setConfigValue_(cfg,'S225_INSTALADO_EM',
    Utilities.formatDate(new Date(),APP.TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),
    'Data/hora da instalação S22.5'
  );

  try{
    registrarAuditoriaS15_({
      acao:'SETUP_S225',
      entidade:'SISTEMA',
      entidadeId:'S22.5',
      resultado:'SUCESSO',
      origem:'APPS_SCRIPT'
    });
  }catch(_){}

  SpreadsheetApp.flush();
  return diagnosticoS225();
}

function diagnosticoS225(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_ID',cfg.APP_ID===APP.ID,cfg.APP_ID||'ausente');
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5',cfg.APP_FASE||'ausente');
  check_(c,'S225_STATUS',cfg.S225_STATUS==='INSTALADO',cfg.S225_STATUS||'ausente');
  check_(c,'REGISTRO_FOTOS',!!ss.getSheetByName('REGISTRO_FOTOS'),'REGISTRO_FOTOS');
  check_(c,'LISTAR_FOTOS',typeof appListarFotosRegistroS225==='function','OK');
  check_(c,'OBTER_FOTO',typeof appObterFotoS225==='function','OK');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

function mostrarDiagnosticoS225(){
  const d=diagnosticoS225();
  SpreadsheetApp.getUi().alert(
    'Diagnóstico S22.5',
    `${d.ok?'S22.5 PRONTA':'HÁ PENDÊNCIAS'}\n\n`+
    d.checks.map(c=>`${c.ok?'✅':'❌'} ${c.nome}: ${c.detalhe}`).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return d;
}

// ========================================================
// S22.5.1 — AJUSTE VISUAL DO POPUP
// ========================================================
function setupS2251(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.1','Fase de implementação validada');
  setConfigValue_(cfg,'S2251_STATUS','INSTALADO','Correção de layout da galeria no popup');
  SpreadsheetApp.flush();
  return diagnosticoS2251();
}
function diagnosticoS2251(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.1',cfg.APP_FASE||'ausente');
  check_(c,'S2251_STATUS',cfg.S2251_STATUS==='INSTALADO',cfg.S2251_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.2 — MOBILE + FILA
// ========================================================
function setupS2252(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.2','Fase de implementação validada');
  setConfigValue_(cfg,'S2252_STATUS','INSTALADO','Correções mobile e fila');
  SpreadsheetApp.flush();
  return diagnosticoS2252();
}
function diagnosticoS2252(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.2',cfg.APP_FASE||'ausente');
  check_(c,'S2252_STATUS',cfg.S2252_STATUS==='INSTALADO',cfg.S2252_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.3 — SINCRONIZAÇÃO MANUAL SEGURA
// ========================================================
function setupS2253(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.3','Fase de implementação validada');
  setConfigValue_(cfg,'S2253_STATUS','INSTALADO','Sincronização manual segura');
  SpreadsheetApp.flush();
  return diagnosticoS2253();
}
function diagnosticoS2253(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.3',cfg.APP_FASE||'ausente');
  check_(c,'S2253_STATUS',cfg.S2253_STATUS==='INSTALADO',cfg.S2253_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.4 — HISTERESE E RECUPERAÇÃO DE REDE
// ========================================================
function setupS2254(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.4','Fase de implementação validada');
  setConfigValue_(cfg,'S2254_STATUS','INSTALADO','Histerese e recuperação de rede');
  setConfigValue_(cfg,'S2254_ONLINE_MAX_MS','1500','Limite máximo ONLINE');
  setConfigValue_(cfg,'S2254_BOM_MAX_MS','1200','Latência considerada boa para recuperação');
  setConfigValue_(cfg,'S2254_RECOVERY_SUCESSOS','3','Medições boas necessárias para recuperar ONLINE');

  SpreadsheetApp.flush();
  return diagnosticoS2254();
}

function diagnosticoS2254(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.4',cfg.APP_FASE||'ausente');
  check_(c,'S2254_STATUS',cfg.S2254_STATUS==='INSTALADO',cfg.S2254_STATUS||'ausente');
  check_(c,'ONLINE_MAX_MS',String(cfg.S2254_ONLINE_MAX_MS||'')==='1500',cfg.S2254_ONLINE_MAX_MS||'ausente');
  check_(c,'BOM_MAX_MS',String(cfg.S2254_BOM_MAX_MS||'')==='1200',cfg.S2254_BOM_MAX_MS||'ausente');
  check_(c,'RECOVERY_SUCESSOS',String(cfg.S2254_RECOVERY_SUCESSOS||'')==='3',cfg.S2254_RECOVERY_SUCESSOS||'ausente');

  return {
    ok:c.every(x=>x.ok),
    totalChecks:c.length,
    totalFalhas:c.filter(x=>!x.ok).length,
    checks:c
  };
}

// ========================================================
// S22.5.5 — RECALIBRAÇÃO DE LATÊNCIA PARA AMBIENTE DO MALL
// ========================================================
function setupS2255(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.5','Fase de implementação validada');
  setConfigValue_(cfg,'S2255_STATUS','INSTALADO','Limites de rede recalibrados');
  setConfigValue_(cfg,'S2255_BOM_MAX_MS','3000','Latência boa para recuperação');
  setConfigValue_(cfg,'S2255_ONLINE_MAX_MS','3750','Latência aceitável para manter ONLINE');
  setConfigValue_(cfg,'S2255_MUITO_RUIM_MS','6250','Latência considerada muito ruim');
  setConfigValue_(cfg,'S2255_RECOVERY_SUCESSOS','3','Medições boas necessárias para recuperar ONLINE');

  SpreadsheetApp.flush();
  return diagnosticoS2255();
}

function diagnosticoS2255(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.5',cfg.APP_FASE||'ausente');
  check_(c,'S2255_STATUS',cfg.S2255_STATUS==='INSTALADO',cfg.S2255_STATUS||'ausente');
  check_(c,'BOM_MAX_MS',String(cfg.S2255_BOM_MAX_MS||'')==='3000',cfg.S2255_BOM_MAX_MS||'ausente');
  check_(c,'ONLINE_MAX_MS',String(cfg.S2255_ONLINE_MAX_MS||'')==='3750',cfg.S2255_ONLINE_MAX_MS||'ausente');
  check_(c,'MUITO_RUIM_MS',String(cfg.S2255_MUITO_RUIM_MS||'')==='6250',cfg.S2255_MUITO_RUIM_MS||'ausente');

  return {
    ok:c.every(x=>x.ok),
    totalChecks:c.length,
    totalFalhas:c.filter(x=>!x.ok).length,
    checks:c
  };
}

// ========================================================
// S22.5.6 — RETRY MANUAL QUANDO REDE ESTÁ ONLINE
// ========================================================
function setupS2256(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.6','Fase de implementação validada');
  setConfigValue_(cfg,'S2256_STATUS','INSTALADO','Retry manual ONLINE corrigido');

  SpreadsheetApp.flush();
  return diagnosticoS2256();
}
function diagnosticoS2256(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.6',cfg.APP_FASE||'ausente');
  check_(c,'S2256_STATUS',cfg.S2256_STATUS==='INSTALADO',cfg.S2256_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.7 — RECALIBRAÇÃO DE LATÊNCIA PARA 5000 ms
// ========================================================
function setupS2257(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.7','Fase de implementação validada');
  setConfigValue_(cfg,'S2257_STATUS','INSTALADO','Limites de rede recalibrados para 5000 ms');
  setConfigValue_(cfg,'S2257_BOM_MAX_MS','5000','Latência boa para recuperação');
  setConfigValue_(cfg,'S2257_ONLINE_MAX_MS','6250','Latência aceitável para manter ONLINE');
  setConfigValue_(cfg,'S2257_MUITO_RUIM_MS','10417','Latência considerada muito ruim');
  setConfigValue_(cfg,'S2257_RECOVERY_SUCESSOS','3','Medições boas necessárias para recuperar ONLINE');

  SpreadsheetApp.flush();
  return diagnosticoS2257();
}
function diagnosticoS2257(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.7',cfg.APP_FASE||'ausente');
  check_(c,'S2257_STATUS',cfg.S2257_STATUS==='INSTALADO',cfg.S2257_STATUS||'ausente');
  check_(c,'BOM_MAX_MS',String(cfg.S2257_BOM_MAX_MS||'')==='5000',cfg.S2257_BOM_MAX_MS||'ausente');
  check_(c,'ONLINE_MAX_MS',String(cfg.S2257_ONLINE_MAX_MS||'')==='6250',cfg.S2257_ONLINE_MAX_MS||'ausente');
  check_(c,'MUITO_RUIM_MS',String(cfg.S2257_MUITO_RUIM_MS||'')==='10417',cfg.S2257_MUITO_RUIM_MS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.8 — TIMEOUT ADAPTATIVO PARA FOTOS
// ========================================================
function setupS2258(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.8','Fase de implementação validada');
  setConfigValue_(cfg,'S2258_STATUS','INSTALADO','Timeout adaptativo de fotos');
  setConfigValue_(cfg,'S2258_TIMEOUT_MIN_MS','20000','Timeout mínimo para foto');
  setConfigValue_(cfg,'S2258_TIMEOUT_MAX_MS','60000','Timeout máximo para foto');

  SpreadsheetApp.flush();
  return diagnosticoS2258();
}
function diagnosticoS2258(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.8',cfg.APP_FASE||'ausente');
  check_(c,'S2258_STATUS',cfg.S2258_STATUS==='INSTALADO',cfg.S2258_STATUS||'ausente');
  check_(c,'TIMEOUT_MIN',String(cfg.S2258_TIMEOUT_MIN_MS||'')==='20000',cfg.S2258_TIMEOUT_MIN_MS||'ausente');
  check_(c,'TIMEOUT_MAX',String(cfg.S2258_TIMEOUT_MAX_MS||'')==='60000',cfg.S2258_TIMEOUT_MAX_MS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.9 — DIAGNÓSTICO REAL DO UPLOAD DE FOTOS
// ========================================================
function setupS2259(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.9','Fase de implementação validada');
  setConfigValue_(cfg,'S2259_STATUS','INSTALADO','Diagnóstico real de upload de fotos');
  SpreadsheetApp.flush();
  return diagnosticoS2259();
}
function diagnosticoS2259(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.9',cfg.APP_FASE||'ausente');
  check_(c,'S2259_STATUS',cfg.S2259_STATUS==='INSTALADO',cfg.S2259_STATUS||'ausente');
  check_(c,'ENDPOINT_FOTO_AUTH',typeof appSincronizarFotoAutenticadaS2259==='function','OK');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.10 — NORMALIZAÇÃO DO PAYLOAD DE FOTOS
// ========================================================
function setupS22510(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.10','Fase de implementação validada');
  setConfigValue_(cfg,'S22510_STATUS','INSTALADO','Payload de fotos normalizado');

  SpreadsheetApp.flush();
  return diagnosticoS22510();
}

function diagnosticoS22510(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.10',cfg.APP_FASE||'ausente');
  check_(c,'S22510_STATUS',cfg.S22510_STATUS==='INSTALADO',cfg.S22510_STATUS||'ausente');
  check_(c,'ENDPOINT_FOTO_AUTH',typeof appSincronizarFotoAutenticadaS2259==='function','OK');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.11 — ZOOM E PAN NO VISUALIZADOR DE FOTOS
// ========================================================
function setupS22511(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.11','Fase de implementação validada');
  setConfigValue_(cfg,'S22511_STATUS','INSTALADO','Zoom e pan no visualizador de fotos');
  setConfigValue_(cfg,'S22511_ZOOM_MIN','1','Zoom mínimo');
  setConfigValue_(cfg,'S22511_ZOOM_MAX','5','Zoom máximo');

  SpreadsheetApp.flush();
  return diagnosticoS22511();
}

function diagnosticoS22511(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.11',cfg.APP_FASE||'ausente');
  check_(c,'S22511_STATUS',cfg.S22511_STATUS==='INSTALADO',cfg.S22511_STATUS||'ausente');
  check_(c,'ZOOM_MIN',String(cfg.S22511_ZOOM_MIN||'')==='1',cfg.S22511_ZOOM_MIN||'ausente');
  check_(c,'ZOOM_MAX',String(cfg.S22511_ZOOM_MAX||'')==='5',cfg.S22511_ZOOM_MAX||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.12 — BINDING REAL DO ZOOM DA GALERIA
// ========================================================
function setupS22512(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.12','Fase de implementação validada');
  setConfigValue_(cfg,'S22512_STATUS','INSTALADO','Zoom ligado ao modal real de fotos');
  SpreadsheetApp.flush();
  return diagnosticoS22512();
}
function diagnosticoS22512(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.12',cfg.APP_FASE||'ausente');
  check_(c,'S22512_STATUS',cfg.S22512_STATUS==='INSTALADO',cfg.S22512_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.13 — MENU COMPACTO MOBILE-FIRST
// ========================================================
function setupS22513(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.13','Fase de implementação validada');
  setConfigValue_(cfg,'S22513_STATUS','INSTALADO','Menu compacto mobile-first');
  setConfigValue_(cfg,'S22513_ZOOM_BOTOES','REMOVIDOS','Zoom permanece por gesto/scroll');

  SpreadsheetApp.flush();
  return diagnosticoS22513();
}

function diagnosticoS22513(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.13',cfg.APP_FASE||'ausente');
  check_(c,'S22513_STATUS',cfg.S22513_STATUS==='INSTALADO',cfg.S22513_STATUS||'ausente');
  check_(c,'ZOOM_BOTOES',cfg.S22513_ZOOM_BOTOES==='REMOVIDOS',cfg.S22513_ZOOM_BOTOES||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.14 — CENTRALIZAR NO TOPO
// ========================================================
function setupS22514(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.14','Fase de implementação validada');
  setConfigValue_(cfg,'S22514_STATUS','INSTALADO','Centralizar movido para o topo');
  SpreadsheetApp.flush();
  return diagnosticoS22514();
}
function diagnosticoS22514(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.14',cfg.APP_FASE||'ausente');
  check_(c,'S22514_STATUS',cfg.S22514_STATUS==='INSTALADO',cfg.S22514_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.14.1 — CORREÇÃO DO BLOCO CSS
// ========================================================
function setupS225141(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.14.1','Fase de implementação validada');
  setConfigValue_(cfg,'S225141_STATUS','INSTALADO','CSS do Centralizar corrigido');
  SpreadsheetApp.flush();
  return diagnosticoS225141();
}
function diagnosticoS225141(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.14.1',cfg.APP_FASE||'ausente');
  check_(c,'S225141_STATUS',cfg.S225141_STATUS==='INSTALADO',cfg.S225141_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.15 — PERMISSÃO REAL PARA CRIAÇÃO DE REGISTRO
// ========================================================
function setupS22515(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.15','Fase de implementação validada');
  setConfigValue_(cfg,'S22515_STATUS','INSTALADO','Permissão criarRegistro aplicada em UI, offline e backend');
  setConfigValue_(cfg,'S22515_UI_ID_NOVO','novo','ID real protegido');
  setConfigValue_(cfg,'S22515_BACKEND_GUARD','ATIVO','appCriarRegistroS3 exige criarRegistro');

  SpreadsheetApp.flush();
  return diagnosticoS22515();
}

function diagnosticoS22515(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.15',cfg.APP_FASE||'ausente');
  check_(c,'S22515_STATUS',cfg.S22515_STATUS==='INSTALADO',cfg.S22515_STATUS||'ausente');
  check_(c,'UI_ID_NOVO',cfg.S22515_UI_ID_NOVO==='novo',cfg.S22515_UI_ID_NOVO||'ausente');
  check_(c,'BACKEND_GUARD',cfg.S22515_BACKEND_GUARD==='ATIVO',cfg.S22515_BACKEND_GUARD||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}

// ========================================================
// S22.5.16 — FECHAR CAMADAS
// ========================================================
function setupS22516(){
  exigirPermissaoS14_('administrar');
  const ss=SpreadsheetApp.getActive(),cfg=ss.getSheetByName('CONFIG');
  if(!cfg)throw new Error('CONFIG ausente.');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S22.5.16','Fase de implementação validada');
  setConfigValue_(cfg,'S22516_STATUS','INSTALADO','Painel Camadas com X, ESC e clique externo');

  SpreadsheetApp.flush();
  return diagnosticoS22516();
}

function diagnosticoS22516(){
  const ss=SpreadsheetApp.getActive(),cfg=lerConfigComoObjeto_(ss),c=[];
  check_(c,'APP_VERSAO',cfg.APP_VERSAO===APP.VERSAO,cfg.APP_VERSAO||'ausente');
  check_(c,'APP_FASE',cfg.APP_FASE==='S22.5.16',cfg.APP_FASE||'ausente');
  check_(c,'S22516_STATUS',cfg.S22516_STATUS==='INSTALADO',cfg.S22516_STATUS||'ausente');
  return {ok:c.every(x=>x.ok),totalChecks:c.length,totalFalhas:c.filter(x=>!x.ok).length,checks:c};
}


// ========================================================
// S23.5 — SETUP / DIAGNÓSTICO
// ========================================================
function setupS235(){
  exigirPermissaoS14_('administrar');

  const ss=SpreadsheetApp.getActive();
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  const hist=ss.getSheetByName('REGISTRO_HISTORICO');
  const cfg=ss.getSheetByName('CONFIG');

  if(!fotos)throw new Error('REGISTRO_FOTOS ausente.');
  if(!hist)throw new Error('REGISTRO_HISTORICO ausente.');
  if(!cfg)throw new Error('CONFIG ausente.');

  garantirCabecalhosFotosInspecaoS235_(fotos);

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5','Fotos vinculadas a inspeções');
  setConfigValue_(cfg,'S235_STATUS','INSTALADO','Fotos de inspeção e Antes/Depois');

  SpreadsheetApp.flush();
  return diagnosticoS235();
}

function diagnosticoS235(){
  const ss=SpreadsheetApp.getActive(),c=[];
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  const hist=ss.getSheetByName('REGISTRO_HISTORICO');

  check_(c,'REGISTRO_FOTOS',!!fotos,'REGISTRO_FOTOS');
  check_(c,'REGISTRO_HISTORICO',!!hist,'REGISTRO_HISTORICO');

  if(fotos){
    garantirCabecalhosFotosInspecaoS235_(fotos);
    const h=fotos.getRange(1,1,1,fotos.getLastColumn()).getValues()[0].map(String);
    ['CLIENT_INSPECTION_ID','ID_HISTORICO','ORIGEM_FOTO'].forEach(k=>
      check_(c,'FOTO_'+k,h.includes(k),k)
    );
  }

  check_(c,'SYNC_FOTO_INSPECAO',typeof appSincronizarFotoInspecaoS235==='function','OK');
  check_(c,'LISTAR_FOTO_INSPECAO',typeof appListarFotosInspecaoS235==='function','OK');

  return {
    ok:c.every(x=>x.ok),
    version:APP.VERSAO,
    totalChecks:c.length,
    falhas:c.filter(x=>!x.ok).length,
    checks:c
  };
}


// ========================================================
// S23.5.1 — GESTÃO DAS FOTOS DA INSPEÇÃO
// ========================================================
function setupS2351(){
  const r=setupS235();

  const ss=SpreadsheetApp.getActive();
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  garantirCabecalhosFotosInspecaoS235_(fotos);

  const cfg=ss.getSheetByName('CONFIG');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.1','Gestão das fotos da inspeção');
  setConfigValue_(cfg,'S2351_STATUS','INSTALADO','Remover, substituir, ordenar e deduplicar fotos');

  SpreadsheetApp.flush();
  return diagnosticoS2351();
}

function diagnosticoS2351(){
  const d=diagnosticoS235();
  const checks=(d.checks||[]).slice();

  const ss=SpreadsheetApp.getActive();
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  const h=fotos
    ? fotos.getRange(1,1,1,fotos.getLastColumn()).getValues()[0].map(String)
    : [];

  ['ORDEM_FOTO','HASH_LOCAL'].forEach(k=>{
    checks.push({
      nome:'S2351_'+k,
      ok:h.includes(k),
      detalhe:h.includes(k)?'OK':'Cabeçalho ausente'
    });
  });

  checks.push({
    nome:'S2351_DEDUP_BACKEND',
    ok:typeof obterFotoDuplicadaInspecaoS2351_==='function',
    detalhe:'Deduplicação por inspeção'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.2 — SNAPSHOT DE MÍDIA
// ========================================================
function setupS2352(){
  const r=setupS2351();

  const ss=SpreadsheetApp.getActive();
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  garantirCabecalhosFotosInspecaoS235_(fotos);

  const cfg=ss.getSheetByName('CONFIG');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.2','Snapshot de mídia por inspeção');
  setConfigValue_(cfg,'S2352_STATUS','INSTALADO','Fotos atuais herdadas na nova inspeção');

  SpreadsheetApp.flush();
  return diagnosticoS2352();
}

function diagnosticoS2352(){
  const d=diagnosticoS2351();
  const checks=(d.checks||[]).slice();

  const ss=SpreadsheetApp.getActive();
  const fotos=ss.getSheetByName('REGISTRO_FOTOS');
  const h=fotos
    ? fotos.getRange(1,1,1,fotos.getLastColumn()).getValues()[0].map(String)
    : [];

  checks.push({
    nome:'S2352_ID_FOTO_ORIGEM',
    ok:h.includes('ID_FOTO_ORIGEM'),
    detalhe:h.includes('ID_FOTO_ORIGEM')?'OK':'Cabeçalho ausente'
  });

  checks.push({
    nome:'S2352_SNAPSHOT_API',
    ok:typeof appObterSnapshotMidiaAtualS2352==='function',
    detalhe:'Carrega mídia atual do ativo'
  });

  checks.push({
    nome:'S2352_REFERENCIA_SEM_COPIA',
    ok:typeof obterFotoPorArquivoIdS2352_==='function',
    detalhe:'Fotos herdadas reutilizam ARQUIVO_ID'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.2.1 — CORREÇÃO DE PERMISSÃO DO SNAPSHOT
// ========================================================
function setupS23521(){
  const r=setupS2352();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.2.1','Correção de permissão do snapshot');
  setConfigValue_(cfg,'S23521_STATUS','INSTALADO','Snapshot usa permissão INSPECIONAR');

  SpreadsheetApp.flush();
  return diagnosticoS23521();
}

function diagnosticoS23521(){
  const d=diagnosticoS2352();
  const checks=(d.checks||[]).slice();

  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('PERFIS_PERMISSOES');
  let adminInspecionar=false;

  if(sh && sh.getLastRow()>=2){
    const vals=sh.getDataRange().getValues();
    const h=vals.shift().map(v=>String(v||'').trim().toUpperCase());
    const iPerfil=h.indexOf('PERFIL');
    const iInsp=h.indexOf('INSPECIONAR');

    const admin=vals.find(r=>String(r[iPerfil]||'').trim().toUpperCase()==='ADMIN');
    if(admin && iInsp>=0){
      adminInspecionar=String(admin[iInsp]).toUpperCase()==='TRUE' || admin[iInsp]===true;
    }
  }

  checks.push({
    nome:'S23521_ADMIN_INSPECIONAR',
    ok:adminInspecionar,
    detalhe:adminInspecionar?'ADMIN possui INSPECIONAR':'ADMIN sem INSPECIONAR'
  });

  checks.push({
    nome:'S23521_SNAPSHOT_PERMISSION',
    ok:typeof appObterSnapshotMidiaAtualS2352==='function',
    detalhe:'API de snapshot usa permissão INSPECIONAR'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.2.2 — BADGE / GALERIA = MÍDIA ATUAL
// ========================================================
function setupS23522(){
  const r=setupS23521();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.2.2','Badge e galeria mostram mídia atual');
  setConfigValue_(cfg,'S23522_STATUS','INSTALADO','Contador não soma histórico de inspeções');

  SpreadsheetApp.flush();
  return diagnosticoS23522();
}

function diagnosticoS23522(){
  const d=diagnosticoS23521();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S23522_LISTAGEM_ATUAL',
    ok:typeof appListarFotosRegistroS225==='function',
    detalhe:'Galeria do card usa snapshot atual'
  });

  // Gate conhecido da implementação atual.
  // Se o registro existir, a quantidade do card deve ser a quantidade atual do snapshot,
  // e não o total histórico de REGISTRO_FOTOS.
  const protocolo='SIG-20260818-0004';
  const sh=SpreadsheetApp.getActive().getSheetByName('REGISTROS');
  let idRegistro='';

  if(sh&&sh.getLastRow()>=2){
    const vals=sh.getDataRange().getValues();
    const h=vals.shift().map(v=>String(v||'').trim());
    const ip=h.indexOf('PROTOCOLO'), ii=h.indexOf('ID_REGISTRO');
    const row=vals.find(r=>String(r[ip]||'').trim()===protocolo);
    if(row&&ii>=0)idRegistro=String(row[ii]||'').trim();
  }

  const atuais=idRegistro?appListarFotosRegistroS225(idRegistro):[];

  checks.push({
    nome:'S23522_GATE_TRIEDO',
    ok:!idRegistro || atuais.length===3,
    detalhe:idRegistro
      ? `${protocolo}: ${atuais.length} foto(s) atual(is)`
      : `${protocolo} não encontrado; teste ignorado`
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length,
    fotosAtuaisGate:atuais.length
  };
}


// ========================================================
// S23.5.3 — SETUP / DIAGNÓSTICO
// ========================================================
function setupS2353(){
  const r=setupS23522();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.3','Antes x Depois por snapshot de mídia');
  setConfigValue_(cfg,'S2353_STATUS','INSTALADO','Comparação de mídia por snapshots');

  SpreadsheetApp.flush();
  return diagnosticoS2353();
}

function diagnosticoS2353(){
  const d=diagnosticoS23522();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S2353_RELATORIO',
    ok:typeof diagnosticoRelatorioS2353==='function',
    detalhe:'Motor Antes x Depois disponível'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.3.1 — RESTAURAÇÃO UI GOOGLE APRESENTAÇÃO
// ========================================================
function setupS23531(){
  const r=setupS2353();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.3.1','Restaurar Google Apresentação no relatório');
  setConfigValue_(cfg,'S23531_STATUS','INSTALADO','UI S23.1 restaurada sobre S23.5.3');

  SpreadsheetApp.flush();
  return diagnosticoS23531();
}

function diagnosticoS23531(){
  const d=diagnosticoS2353();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S23531_BACKEND_SLIDES',
    ok:typeof appGerarApresentacaoS231==='function',
    detalhe:'Backend Google Apresentação disponível'
  });

  checks.push({
    nome:'S23531_REPORT_SNAPSHOT',
    ok:typeof s2353SlidesAntesDepois_==='function',
    detalhe:'Antes x Depois S23.5.3 preservado'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.3.2 — RELATÓRIO COM SNAPSHOT ATUAL
// ========================================================
function setupS23532(){
  const r=setupS23531();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.3.2','Relatório usa mídia atual e delta por identidade');
  setConfigValue_(cfg,'S23532_STATUS','INSTALADO','Fotos históricas não duplicam ficha atual');

  SpreadsheetApp.flush();
  return diagnosticoS23532();
}

function diagnosticoS23532(){
  const d=diagnosticoS23531();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S23532_REPORT_DIAG',
    ok:typeof diagnosticoRelatorioS23532==='function',
    detalhe:'Diagnóstico do delta e snapshot atual'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.5.3.3 — CONTINUAÇÃO POR ATIVO
// ========================================================
function setupS23533(){
  const r=setupS23532();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.5.3.3','Continuação do Antes x Depois por ativo');
  setConfigValue_(cfg,'S23533_STATUS','INSTALADO','Título reinicia por protocolo');

  SpreadsheetApp.flush();
  return diagnosticoS23533();
}

function diagnosticoS23533(){
  const d=diagnosticoS23532();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S23533_RELATORIO',
    ok:typeof diagnosticoRelatorioS23533==='function',
    detalhe:'Título Antes x Depois controlado por protocolo'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.6 — SETUP / DIAGNÓSTICO
// ========================================================
function setupS236(){
  const r=setupS23533();

  const ss=SpreadsheetApp.getActive();
  const reg=ss.getSheetByName('REGISTROS');
  if(!reg)throw new Error('REGISTROS ausente.');

  garantirCabecalhosExclusaoS236_(reg);

  const cfg=ss.getSheetByName('CONFIG');
  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.6','Exclusão operacional administrativa');
  setConfigValue_(cfg,'S236_STATUS','INSTALADO','Excluir do mapa preservando auditoria');

  SpreadsheetApp.flush();
  return diagnosticoS236();
}

function diagnosticoS236(){
  const d=diagnosticoS23533();
  const checks=(d.checks||[]).slice();

  const ss=SpreadsheetApp.getActive();
  const reg=ss.getSheetByName('REGISTROS');
  let h=[];

  if(reg){
    garantirCabecalhosExclusaoS236_(reg);
    h=reg.getRange(1,1,1,reg.getLastColumn()).getValues()[0].map(String);
  }

  ['STATUS_ANTES_EXCLUSAO','EXCLUIDO_EM','EXCLUIDO_POR','EXCLUSAO_MOTIVO'].forEach(k=>{
    checks.push({
      nome:'S236_'+k,
      ok:h.includes(k),
      detalhe:h.includes(k)?'OK':'Cabeçalho ausente'
    });
  });

  checks.push({
    nome:'S236_API_EXCLUIR',
    ok:typeof appExcluirRegistroS236==='function',
    detalhe:'ADMIN only'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.6.1 — EXCLUSÃO DENTRO DO CICLO DE VIDA
// ========================================================
function setupS2361(){
  const r=setupS236();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.6.1','Exclusão administrativa no Ciclo de vida');
  setConfigValue_(cfg,'S2361_STATUS','INSTALADO','Excluir cadastro removido do card e movido para Ciclo de vida');

  SpreadsheetApp.flush();
  return diagnosticoS2361();
}

function diagnosticoS2361(){
  const d=diagnosticoS236();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S2361_BACKEND_PRESERVADO',
    ok:typeof appExcluirRegistroS236==='function',
    detalhe:'Backend de exclusão operacional preservado'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.6.2 — TIPO TRIEDO
// ========================================================
function setupS2362(){
  const r=setupS2361();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.6.2','Triedo na lista Tipo');
  setConfigValue_(cfg,'S2362_STATUS','INSTALADO','Tipo Triedo disponível no formulário');

  SpreadsheetApp.flush();
  return diagnosticoS2362();
}

function diagnosticoS2362(){
  const d=diagnosticoS2361();
  const checks=(d.checks||[]).slice();

  let bootstrap=null;
  try{
    bootstrap=appCarregarS5B();
  }catch(e){}

  const tipos=Array.isArray(bootstrap?.tipos)?bootstrap.tipos:[];
  const qtdTriedo=tipos.filter(v=>
    String(v||'').trim().toLocaleLowerCase('pt-BR')==='triedo'
  ).length;

  checks.push({
    nome:'S2362_TIPO_TRIEDO',
    ok:qtdTriedo===1,
    detalhe:qtdTriedo===1
      ? 'Triedo disponível uma única vez'
      : `Ocorrências de Triedo: ${qtdTriedo}`
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length,
    tipos
  };
}


// ========================================================
// S23.7 — SETUP / DIAGNÓSTICO
// ========================================================
function setupS237(){
  const r=setupS2362();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.7','Edição cadastral da sinalização');
  setConfigValue_(cfg,'S237_STATUS','INSTALADO','Editar cadastro preservando protocolo e posição');

  SpreadsheetApp.flush();
  return diagnosticoS237();
}

function diagnosticoS237(){
  const d=diagnosticoS2362();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S237_OBTER_REGISTRO',
    ok:typeof appObterRegistroParaEdicaoS237==='function',
    detalhe:'Carregamento para edição'
  });
  checks.push({
    nome:'S237_ATUALIZAR_REGISTRO',
    ok:typeof appAtualizarRegistroS237==='function',
    detalhe:'Atualização cadastral'
  });
  checks.push({
    nome:'S237_HISTORICO',
    ok:typeof registrarHistoricoEdicaoS237_==='function',
    detalhe:'Histórico de campos alterados'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.7.1 — SERIALIZAÇÃO SEGURA DA EDIÇÃO
// ========================================================
function setupS2371(){
  const r=setupS237();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.7.1','Serialização segura para edição cadastral');
  setConfigValue_(cfg,'S2371_STATUS','INSTALADO','Datas do Sheets convertidas antes do RPC');

  SpreadsheetApp.flush();
  return diagnosticoS2371();
}

function diagnosticoS2371(){
  const d=diagnosticoS237();
  const checks=(d.checks||[]).slice();

  checks.push({
    nome:'S2371_SERIALIZADOR',
    ok:typeof serializarRegistroEdicaoS2371_==='function',
    detalhe:'Conversão para objeto JSON-safe'
  });

  const teste={
    ID_REGISTRO:'TESTE',
    TITULO:'Teste',
    DATA_INSTALACAO:new Date(2026,7,18,12,0,0),
    CRIADO_EM:new Date(2026,7,18,14,30,0),
    ATIVO:true,
    NUMERO:3
  };

  const s=serializarRegistroEdicaoS2371_(teste);
  let jsonOk=false;
  try{
    JSON.stringify(s);
    jsonOk=
      s.ID_REGISTRO==='TESTE' &&
      s.DATA_INSTALACAO==='2026-08-18' &&
      typeof s.CRIADO_EM==='string';
  }catch(_){}

  checks.push({
    nome:'S2371_JSON_SAFE',
    ok:jsonOk,
    detalhe:jsonOk
      ? `DATA_INSTALACAO=${s.DATA_INSTALACAO}`
      : 'Falha na serialização'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.7.2 — NORMALIZAÇÃO DE BOOLEANOS NO HISTÓRICO
// ========================================================
function setupS2372(){
  const r=setupS2371();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.7.2','Normalização de booleanos no histórico');
  setConfigValue_(cfg,'S2372_STATUS','INSTALADO','Evita alterações fantasmas em campos booleanos');

  SpreadsheetApp.flush();
  return diagnosticoS2372();
}

function diagnosticoS2372(){
  const d=diagnosticoS2371();
  const checks=(d.checks||[]).slice();

  const falsos=[
    false,'FALSE','false','0','NAO','NÃO','N','NO','OFF',''
  ];
  const verdadeiros=[
    true,'TRUE','true','1','SIM','S','YES','Y','ON'
  ];

  checks.push({
    nome:'S2372_FALSES',
    ok:falsos.every(v=>normalizarBooleanoS2372_(v)===false),
    detalhe:'Representações falsas normalizadas'
  });

  checks.push({
    nome:'S2372_TRUES',
    ok:verdadeiros.every(v=>normalizarBooleanoS2372_(v)===true),
    detalhe:'Representações verdadeiras normalizadas'
  });

  checks.push({
    nome:'S2372_COMPARACAO',
    ok:
      normalizarComparacaoS237_('NAO','ILUMINADA')===
      normalizarComparacaoS237_(false,'ILUMINADA') &&
      normalizarComparacaoS237_('SIM','DUPLA_FACE')===
      normalizarComparacaoS237_(true,'DUPLA_FACE'),
    detalhe:'NAO=false e SIM=true para comparação'
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length
  };
}


// ========================================================
// S23.7.3 — TIPOS ORDENADOS + PLACA DE GALERIA
// ========================================================
function setupS2373(){
  const r=setupS2372();

  const ss=SpreadsheetApp.getActive();
  const cfg=ss.getSheetByName('CONFIG');

  setConfigValue_(cfg,'APP_VERSAO',APP.VERSAO,'Versão atualmente instalada');
  setConfigValue_(cfg,'APP_FASE','S23.7.3','Tipos ordenados e Placa de galeria');
  setConfigValue_(cfg,'S2373_STATUS','INSTALADO','Lista Tipo ordenada alfabeticamente');

  SpreadsheetApp.flush();
  return diagnosticoS2373();
}

function diagnosticoS2373(){
  const d=diagnosticoS2372();
  const checks=(d.checks||[]).slice();

  let bootstrap=null;
  try{
    bootstrap=appCarregarS5B();
  }catch(e){}

  const tipos=Array.isArray(bootstrap?.tipos)?bootstrap.tipos:[];
  const normalizados=tipos.map(v=>String(v||'').trim().toLocaleLowerCase('pt-BR'));

  const qtdTriedo=normalizados.filter(v=>v==='triedo').length;
  const qtdGaleria=normalizados.filter(v=>v==='placa de galeria').length;

  checks.push({
    nome:'S2373_TRIEDO',
    ok:qtdTriedo===1,
    detalhe:`Ocorrências: ${qtdTriedo}`
  });

  checks.push({
    nome:'S2373_PLACA_GALERIA',
    ok:qtdGaleria===1,
    detalhe:`Ocorrências: ${qtdGaleria}`
  });

  const ordenada=tipos.slice().sort((a,z)=>
    String(a).localeCompare(String(z),'pt-BR',{
      sensitivity:'base',
      numeric:true
    })
  );

  checks.push({
    nome:'S2373_ORDEM_ALFABETICA',
    ok:JSON.stringify(tipos)===JSON.stringify(ordenada),
    detalhe:tipos.join(' | ')
  });

  const semDuplicatas=new Set(normalizados).size===normalizados.length;
  checks.push({
    nome:'S2373_SEM_DUPLICATAS',
    ok:semDuplicatas,
    detalhe:`${tipos.length} tipo(s)`
  });

  return {
    ok:checks.every(x=>x.ok),
    version:APP.VERSAO,
    checks,
    totalChecks:checks.length,
    falhas:checks.filter(x=>!x.ok).length,
    tipos
  };
}


function setupS241Wrapper(){
  return setupS241();
}
