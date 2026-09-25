/**
 * build_showroom.js
 * Compilador do Showroom Interativo das Telas e Modais do CEOP.
 * Lê src/index.html e src/styles.html, injeta o controlador do Showroom
 * com catálogo das 50 telas/modais e mock states ricos.
 */

const fs = require('fs');
const path = require('path');

const srcDir = path.join(__dirname, 'src');
const indexPath = path.join(srcDir, 'index.html');
const stylesPath = path.join(srcDir, 'styles.html');
const outputPath = path.join(__dirname, 'showroom.html');

console.log('Construindo Showroom CEOP...');

let indexHtml = fs.readFileSync(indexPath, 'utf-8');
let stylesHtml = fs.readFileSync(stylesPath, 'utf-8');

// Extrai apenas o conteúdo interno de <style>...</style> de styles.html se presente
let stylesContent = stylesHtml;
if (stylesHtml.includes('<style>') && stylesHtml.includes('</style>')) {
  stylesContent = stylesHtml.replace(/<style>([\s\S]*?)<\/style>/i, '$1');
}

// Substitui <?!= include('styles'); ?> pelo CSS inline real
indexHtml = indexHtml.replace(/<\?!\=\s*include\(['"]styles['"]\);\s*\?>/g, `<style>\n${stylesContent}\n</style>`);

// Remove ou substitui <?!= include('scripts'); ?> por um mock runtime seguro para o showroom
const mockRuntimeAndShowroomScript = `
<style>
/* =========================================================
   SHOWROOM DOCK & TOOLBAR STYLES (Glassmorphic Premium UI)
   ========================================================= */
#ceopShowroomDock {
  position: fixed;
  bottom: 18px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 99999;
  background: rgba(19, 23, 79, 0.94);
  backdrop-filter: blur(14px);
  -webkit-backdrop-filter: blur(14px);
  border: 1px solid rgba(255, 255, 255, 0.22);
  border-radius: 20px;
  box-shadow: 0 16px 40px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(245, 0, 135, 0.35);
  padding: 10px 18px;
  display: flex;
  align-items: center;
  gap: 12px;
  color: #fff;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 13px;
  max-width: 96vw;
  box-sizing: border-box;
  transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
}

#ceopShowroomDock .dock-brand {
  display: flex;
  align-items: center;
  gap: 8px;
  font-weight: 800;
  letter-spacing: 0.3px;
  color: #fff;
  white-space: nowrap;
  border-right: 1px solid rgba(255, 255, 255, 0.2);
  padding-right: 14px;
}

#ceopShowroomDock .dock-badge {
  background: #f50087;
  color: #fff;
  font-size: 10px;
  font-weight: 800;
  padding: 3px 7px;
  border-radius: 999px;
  letter-spacing: 0.5px;
}

#ceopShowroomDock .dock-select {
  background: rgba(255, 255, 255, 0.12);
  border: 1px solid rgba(255, 255, 255, 0.28);
  border-radius: 10px;
  color: #fff;
  padding: 8px 12px;
  font-size: 13px;
  font-weight: 600;
  outline: none;
  cursor: pointer;
  max-width: 340px;
}

#ceopShowroomDock .dock-select option {
  background: #171b68;
  color: #fff;
}

#ceopShowroomDock .dock-btn {
  background: rgba(255, 255, 255, 0.14);
  border: 1px solid rgba(255, 255, 255, 0.22);
  border-radius: 10px;
  color: #fff;
  font-size: 13px;
  font-weight: 700;
  padding: 8px 14px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.15s ease;
  white-space: nowrap;
}

#ceopShowroomDock .dock-btn:hover {
  background: rgba(255, 255, 255, 0.28);
  transform: translateY(-1px);
}

#ceopShowroomDock .dock-btn.primary {
  background: #f50087;
  border-color: #f50087;
}

#ceopShowroomDock .dock-btn.primary:hover {
  background: #ff199b;
  box-shadow: 0 4px 14px rgba(245, 0, 135, 0.5);
}

#ceopShowroomDock .dock-btn.danger {
  background: rgba(239, 68, 68, 0.3);
  border-color: rgba(239, 68, 68, 0.6);
  color: #fca5a5;
}

#ceopShowroomDock .dock-btn.danger:hover {
  background: rgba(239, 68, 68, 0.8);
  color: #fff;
}

#ceopShowroomDock .dock-counter {
  font-weight: 700;
  font-size: 12px;
  color: rgba(255, 255, 255, 0.7);
  white-space: nowrap;
}

/* Info pill flutuante no topo */
#ceopScreenInfoBar {
  position: fixed;
  top: 14px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 99998;
  background: rgba(16, 18, 40, 0.88);
  backdrop-filter: blur(10px);
  border: 1px solid rgba(255, 255, 255, 0.15);
  border-radius: 999px;
  padding: 6px 18px;
  color: #fff;
  font-size: 12px;
  display: flex;
  align-items: center;
  gap: 10px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.3);
  pointer-events: none;
}
#ceopScreenInfoBar strong { color: #f50087; }

/* Viewport emulator frame */
body.showroom-mobile-mode {
  display: flex;
  justify-content: center;
  background: #111428;
}
body.showroom-mobile-mode > main {
  width: 390px !important;
  max-width: 390px !important;
  min-height: 844px;
  border-left: 2px solid rgba(255,255,255,0.1);
  border-right: 2px solid rgba(255,255,255,0.1);
  background: var(--bg, #f4f5f8);
  box-shadow: 0 0 50px rgba(0,0,0,0.7);
}
body.showroom-tablet-mode {
  display: flex;
  justify-content: center;
  background: #111428;
}
body.showroom-tablet-mode > main {
  width: 768px !important;
  max-width: 768px !important;
  min-height: 1024px;
  border-left: 2px solid rgba(255,255,255,0.1);
  border-right: 2px solid rgba(255,255,255,0.1);
  background: var(--bg, #f4f5f8);
  box-shadow: 0 0 50px rgba(0,0,0,0.7);
}

/* Força que o container de fundo não corte os modais no showroom */
.offline-panel {
  transition: opacity 0.2s ease;
}
</style>

<div id="ceopScreenInfoBar">
  <span>Visão Atual:</span>
  <strong id="showroomInfoTitle">Mapa & HUD Operacional</strong>
  <span id="showroomInfoTarget" style="color:#94a3b8; font-family:monospace;">(Base)</span>
</div>

<div id="ceopShowroomDock" role="toolbar" aria-label="Showroom de Telas CEOP">
  <div class="dock-brand">
    <span>CEOP Showroom</span>
    <span class="dock-badge">55 Telas</span>
  </div>

  <button id="dockPrevBtn" class="dock-btn" type="button" title="Tela Anterior (Seta Esquerda)">‹ Anterior</button>

  <select id="dockScreenSelect" class="dock-select" aria-label="Selecione a tela ou modal"></select>

  <button id="dockNextBtn" class="dock-btn primary" type="button" title="Próxima Tela (Seta Direita)">Próxima ›</button>

  <button id="dockResetBtn" class="dock-btn" type="button" title="Fechar painéis e voltar ao mapa">⌖ Mapa</button>

  <div class="dock-counter" id="dockCounter">01 / 55</div>

  <button id="dockDeviceBtn" class="dock-btn" type="button" title="Alternar viewport (Desktop / Mobile / Tablet)">📱 Modo</button>
</div>

<script>
(function () {
  // Catálogo com todas as 50 telas/modais mapeadas
  const SCREENS = [
    // --- GRUPO 1: MAPA & HUD PRINCIPAL ---
    { id: 'view_mapa_base', name: '01. Mapa Principal & HUD do Mall', type: 'base', desc: 'Planta interativa, toolbar, seletor de pisos e marcadores' },
    { id: 'appMenuS22513', name: '02. Menu Lateral do App (Principal)', type: 'drawer', desc: 'Gaveta de navegação lateral com ações rápidas e gestão' },
    { id: 'appMenuCartografiaS268D', name: '03. Menu Lateral (Submenu Cartografia)', type: 'drawer_sub', desc: 'Submenu de ferramentas cartográficas avançadas' },
    { id: 'camadas', name: '04. Menu Rápido de Camadas', type: 'popup', desc: 'Toggle rápido de camadas ativas sobre o mapa' },
    { id: 'centralCamadasS261', name: '05. Central Avançada de Camadas', type: 'drawer', desc: 'Gerenciador com visibilidade, legendas e presets corporativos' },
    { id: 'sinalizacaoMapaCard', name: '06. Card de Sinalização Selecionada', type: 'card', desc: 'Card flutuante com dados, fotos, estado e ações rápidas' },
    { id: 'localCard', name: '07. Card de Local Geométrico Selecionado', type: 'card', desc: 'Card flutuante com coordenadas e botão Confirmar Local' },
    { id: 'legendaOcupacaoMapa', name: '08. Legenda de Ocupação Comercial', type: 'card', desc: 'Legenda flutuante com cores de ocupação das lojas' },

    // --- GRUPO 2: OPERAÇÃO & SINALIZAÇÕES (AS) ---
    { id: 'formPanel', name: '09. Painel de Cadastro/Edição (Ativo / AS)', type: 'drawer', desc: 'Formulário completo de ativo e Autorização de Serviço' },
    { id: 'inspecaoPanel', name: '10. Painel de Inspeção de Campo (S8)', type: 'modal', desc: 'Vistoria presencial, estado físico e checklist' },
    { id: 'historicoPanel', name: '11. Linha do Tempo e Histórico (S8)', type: 'modal', desc: 'Histórico cronológico de inspeções e alterações' },
    { id: 'pendenciasPanel', name: '12. Painel de Pendências e Chamados (S10)', type: 'modal', desc: 'Gestão de anomalias com prioridades e resoluções' },
    { id: 'cicloVidaPanelS18', name: '13. Painel de Ciclo de Vida do Ativo (S18)', type: 'modal', desc: 'Vida útil, lonas, ferragens e desgaste acumulado' },
    { id: 'eventoCicloEditorS18', name: '14. Modal de Evento de Ciclo de Vida (S18)', type: 'modal', desc: 'Submodal de lançamento de troca ou reforma' },
    { id: 'reagendarPanelS191', name: '15. Modal de Reagendamento (S19.1)', type: 'modal', desc: 'Ajuste de vencimento de AS ou manutenção' },
    { id: 'fotoModalS225', name: '16. Galeria de Fotos em Alta Resolução', type: 'modal_gallery', desc: 'Visualizador de fotos em tela cheia com zoom' },

    // --- GRUPO 3: LOJISTAS & ESPAÇOS COMERCIAIS ---
    { id: 'lojistasPanel', name: '17. Central de Gestão de Lojistas', type: 'modal', desc: 'Base mestre de espaços, contratos, LUCs e ocupação' },
    { id: 'lojistaModal', name: '18. Modal Completo de Cadastro de Lojista', type: 'modal_overlay', desc: 'Ficha mestre com abas de contrato, contatos e AS' },

    // --- GRUPO 4: TORRES E NÚCLEOS VERTICAIS ---
    { id: 'torresPanelS2610C', name: '19. Central de Torres e Núcleos Verticais', type: 'modal', desc: 'Governança vertical N0-N3, elevadores e snapshots' },

    // --- GRUPO 5: REFERÊNCIAS & CATÁLOGOS ---
    { id: 'referenciasPanelS2610R3', name: '20. Central de Referências Espaciais', type: 'modal', desc: 'Consulta e filtros de pontos notáveis do mall' },
    { id: 'referenciaNovaS2610R4', name: '21. Painel de Nova Referência', type: 'modal', desc: 'Formulário de cadastro inicial de referência' },
    { id: 'referenciaConfirmarS2610R4', name: '22. Confirmação de Cadastro de Referência', type: 'modal', desc: 'Validação de posição e metadados' },
    { id: 'referenciaReposicionarConfirmarS2610R5', name: '23. Confirmação de Reposicionamento', type: 'modal', desc: 'Atualização de coordenadas espaciais' },
    { id: 'referenciaRemoverS2610R6C', name: '24. Modal de Remoção de Referência', type: 'modal', desc: 'Exclusão com verificação e digitação de confirmação' },
    { id: 'referenciaEditorS2610R3', name: '25. Editor Detalhado de Referência', type: 'modal', desc: 'Edição de propriedades complementares' },
    { id: 'referenciaCatalogoS2610R8', name: '26. Catálogo de Tipos de Referência', type: 'modal', desc: 'Categorias e taxonomias de referências' },
    { id: 'catalogosPanelS269C1', name: '27. Central de Catálogos de Domínio', type: 'modal', desc: 'Taxonomias de materiais, tipos e situações' },
    { id: 'catalogoOpcaoEditorS269C2', name: '28. Editor de Opção de Catálogo', type: 'modal', desc: 'Criação/edição de item de catálogo' },
    { id: 'catalogoHistoricoPanelS269C2', name: '29. Histórico e Auditoria dos Catálogos', type: 'modal', desc: 'Trilha de auditoria das taxonomias' },

    // --- GRUPO 6: ANÁLISE, DASHBOARD & RELATÓRIOS ---
    { id: 'centralGestaoPanel', name: '30. Central Geral de Gestão Operacional', type: 'modal', desc: 'Visão executiva, KPIs de campo e ações em lote' },
    { id: 'dashboardPanel', name: '31. Dashboard Analítico de Indicadores', type: 'modal', desc: 'Gráficos de conformidade e distribuição por piso' },
    { id: 'relatoriosPanel', name: '32. Painel de Relatórios & Google Slides', type: 'modal', desc: 'Exportação em lote e geração de apresentações' },
    { id: 'agendaPanelS19', name: '33. Agenda Operacional e Cronograma', type: 'modal', desc: 'Calendário de trocas e vencimentos de AS' },
    { id: 'alertasPanelS21', name: '34. Central Proativa de Alertas', type: 'modal', desc: 'Monitoramento de AS vencidas e itens danificados' },
    { id: 'planoEditorS20', name: '35. Editor de Planos Operacionais', type: 'modal', desc: 'Configuração de rotinas preventivas' },
    { id: 'regraEditorS22', name: '36. Editor de Regras de Alerta', type: 'modal', desc: 'Disparadores automáticos de notificações' },
    { id: 'escalamentoTesteS222', name: '37. Teste de Escalonamento de Alertas', type: 'modal', desc: 'Simulador de envio de notificações' },

    // --- GRUPO 7: CARTOGRAFIA ADMINISTRATIVA ---
    { id: 'calibracaoPanelS242', name: '38. Calibração dos Níveis e Plantas', type: 'modal', desc: 'Alinhamento de coordenadas com plantas 2025' },
    { id: 'areasNivel0PanelS268D', name: '39. Delimitação do Subsolo / Nível 0', type: 'modal', desc: 'Polígonos de estacionamento, docas e acessos N0' },
    { id: 'areasNivel1PanelS246', name: '40. Áreas Especiais do Nível 1', type: 'modal', desc: 'Polígonos de Hotel, CDM e áreas externas' },
    { id: 'areaVermelhaPanelS244', name: '41. Estacionamento Nível 3 (Área Vermelha)', type: 'modal', desc: 'Delimitação de vagas e setor vermelho' },
    { id: 'cartoPublicacaoPanelS255', name: '42. Publicação Cartográfica (Rascunho/Prod)', type: 'modal', desc: 'Revisão e publicação controlada de mapas' },
    { id: 'restaurarCartografiaPanelS254', name: '43. Restauração de Versões Cartográficas', type: 'modal', desc: 'Rollback seguro de plantas' },
    { id: 'cartoHistoricoDetalhePanelS253', name: '44. Detalhe do Histórico Cartográfico', type: 'modal', desc: 'Auditoria visual de alterações em mapas' },

    // --- GRUPO 8: ADMINISTRAÇÃO, SEGURANÇA & OFFLINE ---
    { id: 'adminPanelS14', name: '45. Central de Administração do Sistema', type: 'modal', desc: 'Painel com abas de usuários, perfis, auditoria e backups' },
    { id: 'usuarioEditorS14', name: '46. Modal Editor de Usuário', type: 'modal', desc: 'Cadastro de operadores e permissões' },
    { id: 'perfilEditorS14', name: '47. Modal Editor de Perfis de Acesso', type: 'modal', desc: 'Matriz de privilégios e papéis' },
    { id: 'restorePanelS16', name: '48. Modal de Restauração de Backups SHA-256', type: 'modal', desc: 'Restauração com chave de integridade' },
    { id: 'filaPanel', name: '49. Fila de Sincronização Offline (Outbox)', type: 'modal', desc: 'Itens locais pendentes de envio ao servidor' },
    { id: 'offlinePanel', name: '50. Diagnóstico do Cache Offline', type: 'modal', desc: 'Barra de progresso, status e reparo local' },

    // --- GRUPO 9: PATRIMÔNIO & RECLAMAÇÕES (CEOP) ---
    { id: 'formPatrimonioModal', name: '51. Solicitação de Ordem de Patrimônio (OS)', type: 'modal_overlay', desc: 'Formulário com 25 serviços, tags de setor e dados' },
    { id: 'gestaoPatrimonioPanel', name: '52. Gestão de Ordens de Patrimônio', type: 'modal', desc: 'Painel de OS com KPIs, filtros e emissão de termos' },
    { id: 'formReclamacaoModal', name: '53. Novo Chamado de Reclamação (SAC)', type: 'modal_overlay', desc: 'Chamados com upload de fotos/áudio/vídeo e LGPD' },
    { id: 'gestaoReclamacoesPanel', name: '54. Central de Gestão de Reclamações', type: 'modal', desc: 'Triagem, averiguação com mídias e resolução' },
    { id: 'termoAssinaturaModal', name: '55. Coleta de Assinatura Digital & Termos', type: 'modal_overlay', desc: 'Assinatura touch em tela e emissão de PDF no Drive' }
  ];

  let currentIndex = 0;
  let deviceModes = ['desktop', 'tablet', 'mobile'];
  let currentDeviceModeIndex = 0;

  // Popula o select do Dock
  const select = document.getElementById('dockScreenSelect');
  SCREENS.forEach((s, idx) => {
    const opt = document.createElement('option');
    opt.value = idx;
    opt.textContent = s.name;
    select.appendChild(opt);
  });

  function fecharTodasAsTelas() {
    // Esconde painéis e modais
    document.querySelectorAll('.offline-panel, .form-panel, .layers, .modal-overlay').forEach(el => {
      el.hidden = true;
      el.setAttribute('aria-hidden', 'true');
      el.classList.remove('open');
    });
    // Esconde cards flutuantes
    ['sinalizacaoMapaCard', 'localCard', 'legendaOcupacaoMapa', 'camadas', 'centralCamadasS261', 'appMenuS22513'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.hidden = true;
    });
    const bdrop = document.getElementById('appMenuBackdropS22513');
    if (bdrop) bdrop.hidden = true;
    const cbdrop = document.getElementById('centralCamadasBackdropS261');
    if (cbdrop) cbdrop.hidden = true;
  }

  function injetarDadosSimuladosSeNecessario(screen) {
    // 1. Simulação para o Card de Sinalização
    const sigCard = document.getElementById('sinalizacaoMapaCard');
    if (sigCard) {
      const proto = document.getElementById('sigCardProtocolo');
      if (proto) proto.textContent = 'CEOP-2026-N1-0842';
      const tit = document.getElementById('sigCardTitulo');
      if (tit) tit.textContent = 'Totem Direcional — Praça de Alimentação';
      const tp = document.getElementById('sigCardTipo');
      if (tp) tp.textContent = 'Tipo: Totem Iluminado Dupla Face';
      const loc = document.getElementById('sigCardLocal');
      if (loc) loc.textContent = 'Localização: Corredor Central (Piso L1, próx. LUC 1024)';
      const resp = document.getElementById('sigCardResponsavel');
      if (resp) resp.textContent = 'Responsável: Equipe Operações Mall';
    }

    // 2. Simulação para Localizador e Card de Ponto Selecionado
    const locResumo = document.getElementById('localResumo');
    if (locResumo) {
      locResumo.innerHTML = '<strong>Ponto Coordenadas:</strong> X: 1420.50 | Y: 850.25<br><small>Referência mais próxima: Escada Rolante Central (12m)</small>';
    }

    // 3. Simulação para o Formulário de Registro
    const fLocalResumo = document.getElementById('formLocalResumo');
    if (fLocalResumo) fLocalResumo.textContent = 'Ponto selecionado em Piso L1 (Setor Laranja)';
    const fCoords = document.getElementById('formCoords');
    if (fCoords) fCoords.textContent = 'Coordenadas relativas: X 1420.5, Y 850.2';
    const preview = document.getElementById('fotoPreview');
    if (preview && preview.children.length === 0) {
      preview.innerHTML = \`
        <div class="photo-card" style="background:#e2e8f0; display:flex; align-items:center; justify-content:center; font-size:11px; color:#475569;">Foto 1 (Exemplo)</div>
        <div class="photo-card" style="background:#e2e8f0; display:flex; align-items:center; justify-content:center; font-size:11px; color:#475569;">Foto 2 (Exemplo)</div>
      \`;
    }

    // 4. Simulação para Lojistas (Tabela preenchida)
    const lojistasTbody = document.querySelector('#tabelaLojistas tbody, #listaLojistas');
    if (lojistasTbody && lojistasTbody.children.length === 0) {
      lojistasTbody.innerHTML = \`
        <tr style="border-bottom:1px solid #e2e8f0;">
          <td style="padding:10px;"><strong>LUC-1024</strong></td>
          <td style="padding:10px;">Magazine Luiza</td>
          <td style="padding:10px;">Nível 1</td>
          <td style="padding:10px;"><span class="sig-badge" style="background:#16a34a;">OPERANDO</span></td>
          <td style="padding:10px;">Âncora</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <td style="padding:10px;"><strong>LUC-1048</strong></td>
          <td style="padding:10px;">Lojas Americanas</td>
          <td style="padding:10px;">Nível 1</td>
          <td style="padding:10px;"><span class="sig-badge" style="background:#ca8a04;">EM OBRAS</span></td>
          <td style="padding:10px;">Âncora</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <td style="padding:10px;"><strong>LUC-2015</strong></td>
          <td style="padding:10px;">Espaço Vago (Antiga Livraria)</td>
          <td style="padding:10px;">Nível 2</td>
          <td style="padding:10px;"><span class="sig-badge" style="background:#dc2626;">VAGO</span></td>
          <td style="padding:10px;">Satélite</td>
        </tr>
      \`;
    }

    // 5. Simulação de Usuários / Admin
    const listaUsers = document.getElementById('listaUsuariosS14');
    if (listaUsers && listaUsers.children.length === 0) {
      listaUsers.innerHTML = \`
        <div style="padding:12px; border:1px solid #e2e8f0; border-radius:10px; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;">
          <div><strong>operador.mall@ceop.com.br</strong><br><small style="color:#64748b;">Perfil: Operador de Campo (Sinalização & AS)</small></div>
          <span class="sig-badge">ATIVO</span>
        </div>
        <div style="padding:12px; border:1px solid #e2e8f0; border-radius:10px; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center;">
          <div><strong>gestor.operacoes@ceop.com.br</strong><br><small style="color:#64748b;">Perfil: Administrador Geral</small></div>
          <span class="sig-badge" style="background:#f50087;">ADMIN</span>
        </div>
      \`;
    }

    // 6. Simulação de Galeria de Fotos
    const fotoImg = document.getElementById('fotoImagemS225');
    if (fotoImg) {
      fotoImg.src = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="600" height="400" viewBox="0 0 600 400"><rect width="600" height="400" fill="%231e293b"/><text x="50%" y="50%" fill="%23f8fafc" font-family="sans-serif" font-size="24" text-anchor="middle" dominant-baseline="middle">Demonstração da Foto da Sinalização</text></svg>';
      const stage = document.getElementById('fotoStageS225');
      if (stage) stage.hidden = false;
      const loading = document.getElementById('fotoLoadingS225');
      if (loading) loading.hidden = true;
      const rodape = document.getElementById('fotoRodapeS225');
      if (rodape) rodape.hidden = false;
      const nome = document.getElementById('fotoNomeS225');
      if (nome) nome.textContent = 'Foto_Totem_Alimentacao_01.jpg';
      const meta = document.getElementById('fotoMetaS225');
      if (meta) meta.textContent = 'Enviada em 16/09/2026 às 10:14 por Operador 01';
      const count = document.getElementById('fotoContadorS225');
      if (count) count.textContent = '1 de 3';
    }
  }

  function mostrarTela(index) {
    if (index < 0) index = SCREENS.length - 1;
    if (index >= SCREENS.length) index = 0;
    currentIndex = index;

    fecharTodasAsTelas();
    const screen = SCREENS[index];

    // Atualiza info bar e dock
    document.getElementById('showroomInfoTitle').textContent = screen.name;
    document.getElementById('showroomInfoTarget').textContent = \`[#\${screen.id}] • \${screen.desc}\`;
    document.getElementById('dockCounter').textContent = \`\${String(index + 1).padStart(2, '0')} / \${SCREENS.length}\`;
    select.value = index;

    if (screen.id === 'view_mapa_base') {
      return; // Apenas mapa base visível
    }

    injetarDadosSimuladosSeNecessario(screen);

    // Abre a tela alvo
    if (screen.id === 'appMenuS22513') {
      const menu = document.getElementById('appMenuS22513');
      const bdrop = document.getElementById('appMenuBackdropS22513');
      const princ = document.getElementById('appMenuPrincipalS268D');
      const carto = document.getElementById('appMenuCartografiaS268D');
      if (menu) menu.hidden = false;
      if (bdrop) bdrop.hidden = false;
      if (princ) princ.hidden = false;
      if (carto) carto.hidden = true;
    } else if (screen.id === 'appMenuCartografiaS268D') {
      const menu = document.getElementById('appMenuS22513');
      const bdrop = document.getElementById('appMenuBackdropS22513');
      const princ = document.getElementById('appMenuPrincipalS268D');
      const carto = document.getElementById('appMenuCartografiaS268D');
      if (menu) menu.hidden = false;
      if (bdrop) bdrop.hidden = false;
      if (princ) princ.hidden = true;
      if (carto) carto.hidden = false;
    } else if (screen.id === 'centralCamadasS261') {
      const drawer = document.getElementById('centralCamadasS261');
      const bdrop = document.getElementById('centralCamadasBackdropS261');
      if (drawer) { drawer.hidden = false; drawer.setAttribute('aria-hidden', 'false'); }
      if (bdrop) bdrop.hidden = false;
    } else if (screen.id === 'camadas') {
      const popup = document.getElementById('camadas');
      if (popup) popup.hidden = false;
    } else if (screen.id === 'formPanel') {
      const f = document.getElementById('formPanel');
      if (f) {
        f.hidden = false;
        f.setAttribute('aria-hidden', 'false');
        f.classList.add('open');
      }
    } else {
      const el = document.getElementById(screen.id);
      if (el) {
        el.hidden = false;
        el.setAttribute('aria-hidden', 'false');
        if (el.classList.contains('form-panel')) el.classList.add('open');
      }
    }
  }

  // Event Listeners
  select.addEventListener('change', (e) => mostrarTela(parseInt(e.target.value, 10)));
  document.getElementById('dockPrevBtn').addEventListener('click', () => mostrarTela(currentIndex - 1));
  document.getElementById('dockNextBtn').addEventListener('click', () => mostrarTela(currentIndex + 1));
  document.getElementById('dockResetBtn').addEventListener('click', () => mostrarTela(0));

  // Alternador de dispositivo
  document.getElementById('dockDeviceBtn').addEventListener('click', () => {
    currentDeviceModeIndex = (currentDeviceModeIndex + 1) % deviceModes.length;
    const mode = deviceModes[currentDeviceModeIndex];
    document.body.classList.remove('showroom-mobile-mode', 'showroom-tablet-mode');
    if (mode === 'mobile') {
      document.body.classList.add('showroom-mobile-mode');
      document.getElementById('dockDeviceBtn').textContent = '📱 Celular (390px)';
    } else if (mode === 'tablet') {
      document.body.classList.add('showroom-tablet-mode');
      document.getElementById('dockDeviceBtn').textContent = '📱 Tablet (768px)';
    } else {
      document.getElementById('dockDeviceBtn').textContent = '🖥️ Desktop';
    }
  });

  // Atalhos de teclado: Setas e Esc
  window.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
    if (e.key === 'ArrowRight') {
      mostrarTela(currentIndex + 1);
    } else if (e.key === 'ArrowLeft') {
      mostrarTela(currentIndex - 1);
    } else if (e.key === 'Escape') {
      mostrarTela(0);
    }
  });

  // Inicia na tela base do mapa
  mostrarTela(0);
})();
</script>
`;

indexHtml = indexHtml.replace(/<\?!\=\s*include\(['"]scripts['"]\);\s*\?>/g, mockRuntimeAndShowroomScript);

fs.writeFileSync(outputPath, indexHtml, 'utf-8');
console.log('Showroom gerado com sucesso em:', outputPath);
console.log('Tamanho final:', (fs.statSync(outputPath).size / 1024).toFixed(1), 'KB');
