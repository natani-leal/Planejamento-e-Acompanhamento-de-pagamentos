// ==========================================
// ARQUIVO: Code.gs (Backend)
// ==========================================

// Configuração
var SPREADSHEET_ID = '15ZflxC82kDMcK8k4CbZQmX_fO9wDUl86tGYlJ3IMiT4';
var ABA_CONSOLIDADO = 'Consolidado';
// NOVO LINK ATUALIZADO:
var LINK_DIVISAO_EXECUTORES = 'https://script.google.com/macros/s/AKfycbzAXBIR5r5C-8FmgQws4YzL1jzEbNKNBFxLQQ6WjndJL5TeEO-vh3fjp2Dle_Efo3WOzA/exec';

var COL = {
  AGENCIA: 0,      // A
  DATA_CONSOLIDADO: 1, // B
  CAMPANHA: 2,     // C
  NF: 5,           // F
  VEICULO: 9,      // J
  TIPO_MIDIA: 11,  // L - Tipo de Mídia
  STATUS_PAG: 12,  // M
  EXECUTOR: 13,    // N
  ATESTO: 18,      // S
  CONTROLE: 19,    // T
  DATA_PAGO: 20,   // U
  VALOR: 27        // AB
};

// Status exatos na planilha
var STATUS = {
  EM_PROCESSO: 'em processo de pagamento',
  ATESTADA: 'atestada',
  INCONFORMIDADE: 'com inconformidade', 
  PAGA: 'paga'
};

// --- PERSISTENCE HELPERS ---
function getStoredData_() {
  var props = PropertiesService.getScriptProperties();
  var data = props.getProperty('DASHBOARD_DATA');
  return data ? JSON.parse(data) : {};
}

function saveStoredData_(data) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('DASHBOARD_DATA', JSON.stringify(data));
}

function getKey_(row) {
  var controle = String(row[COL.CONTROLE] || 'SEM_CONTROLE');
  return row[COL.AGENCIA] + '|' + row[COL.CAMPANHA] + '|' + controle;
}

// --- PUBLIC FUNCTIONS ---

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Controle de Pagamentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ABA_CONSOLIDADO);
}

function getAllData_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  return data.slice(1); // Remove header
}

function formatDateForJSON_(date) {
  if (!date) return null;
  if (date instanceof Date) {
    return date.toISOString();
  }
  return null;
}

function saveStatus(key, newStatus) {
  var stored = getStoredData_();
  if (!stored[key]) stored[key] = {};
  stored[key].statusPag = newStatus;
  saveStoredData_(stored);
  return true;
}

function saveObs(key, newObs) {
  var stored = getStoredData_();
  if (!stored[key]) stored[key] = {};
  stored[key].obs = newObs;
  saveStoredData_(stored);
  return true;
}

function getEmProcesso() {
  var data = getAllData_();
  var stored = getStoredData_();
  var grupos = {};
  
  data.forEach(function(row) {
    var status = String(row[COL.STATUS_PAG] || '').toLowerCase().trim();
    if (status === STATUS.EM_PROCESSO) {
      var key = getKey_(row);
      if (!grupos[key]) {
        var saved = stored[key] || {};
        grupos[key] = {
          agencia: row[COL.AGENCIA],
          campanha: row[COL.CAMPANHA],
          controle: String(row[COL.CONTROLE] || 'SEM_CONTROLE'),
          qtd: 0,
          valor: 0,
          dataAtesto: null,
          obs: saved.obs || '',
          statusPag: saved.statusPag || 'montando'
        };
      }
      grupos[key].qtd++;
      grupos[key].valor += parseFloat(row[COL.VALOR]) || 0;
      var dataAtesto = row[COL.ATESTO];
      if (dataAtesto && (!grupos[key].dataAtesto || dataAtesto < grupos[key].dataAtesto)) {
        grupos[key].dataAtesto = formatDateForJSON_(dataAtesto);
      }
    }
  });
  
  var result = Object.values(grupos);
  result.sort(function(a, b) {
    if (!a.dataAtesto) return 1;
    if (!b.dataAtesto) return -1;
    return new Date(a.dataAtesto) - new Date(b.dataAtesto);
  });
  return result;
}

function getAtestadas() {
  var data = getAllData_();
  var stored = getStoredData_();
  var grupos = {};
  
  data.forEach(function(row) {
    var status = String(row[COL.STATUS_PAG] || '').toLowerCase().trim();
    var isAtestada = status.indexOf('atestada') !== -1;
    var isInconformidade = status.indexOf('inconformidade') !== -1;
    var isEmAnalise = !status || status.indexOf('analise') !== -1;
    
    if (isAtestada || isInconformidade || isEmAnalise) {
      var groupKey = row[COL.AGENCIA] + '|' + row[COL.CAMPANHA];
      if (!grupos[groupKey]) {
        var saved = stored[groupKey] || {};
        grupos[groupKey] = {
          agencia: row[COL.AGENCIA],
          campanha: row[COL.CAMPANHA],
          qtdEmAnalise: 0,
          valorEmAnalise: 0,
          qtdAtestada: 0,
          valorAtestado: 0,
          qtdPendente: 0,
          valorPendente: 0,
          pendentes: [],
          obs: saved.obs || ''
        };
      }
      var valor = parseFloat(row[COL.VALOR]) || 0;
      
      if (isEmAnalise) {
        grupos[groupKey].qtdEmAnalise++;
        grupos[groupKey].valorEmAnalise += valor;
      } else if (isAtestada && !isInconformidade) {
        grupos[groupKey].qtdAtestada++;
        grupos[groupKey].valorAtestado += valor;
      }
      
      if (isInconformidade) {
        grupos[groupKey].qtdPendente++;
        grupos[groupKey].valorPendente += valor;
        grupos[groupKey].pendentes.push({
          nf: row[COL.NF],
          veiculo: row[COL.VEICULO],
          executor: row[COL.EXECUTOR],
          tipoMidia: row[COL.TIPO_MIDIA],
          tipo: 'Inconformidade',
          valor: valor
        });
      }
    }
  });
  return Object.values(grupos);
}

function getUltimasPagas() {
  var data = getAllData_();
  var grupos = {};
  var hoje = new Date();
  var limite = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  
  data.forEach(function(row) {
    var status = String(row[COL.STATUS_PAG] || '').toLowerCase().trim();
    var dataPago = row[COL.DATA_PAGO];
    if (status === STATUS.PAGA && dataPago && dataPago >= limite) {
      var key = getKey_(row);
      if (!grupos[key]) {
        grupos[key] = {
          agencia: row[COL.AGENCIA],
          campanha: row[COL.CAMPANHA],
          controle: String(row[COL.CONTROLE] || 'SEM_CONTROLE'),
          qtd: 0,
          valor: 0,
          dataPago: formatDateForJSON_(dataPago)
        };
      }
      grupos[key].qtd++;
      grupos[key].valor += parseFloat(row[COL.VALOR]) || 0;
      if (dataPago > new Date(grupos[key].dataPago)) {
        grupos[key].dataPago = formatDateForJSON_(dataPago);
      }
    }
  });
  var result = Object.values(grupos);
  result.sort(function(a, b) {
    return new Date(b.dataPago) - new Date(a.dataPago);
  });
  return result;
}

function getDados() {
  return {
    emProcesso: getEmProcesso(),
    atestadas: getAtestadas(),
    pagas: getUltimasPagas()
  };
}
// ==========================================
// DIVISÃO DE NOTAS POR EXECUTOR
// ==========================================

function getDadosExecutores(dataStr) {
  var data = getAllData_();
  var datas = {};
  var executores = {};
  
  // Coletar datas disponíveis
  data.forEach(function(row) {
    var dataNota = row[COL.DATA_CONSOLIDADO]; // Usando coluna B conforme solicitado
    if (dataNota && dataNota instanceof Date) {
      var key = Utilities.formatDate(dataNota, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (key) datas[key] = true;
    }
  });
  
  var datasOrdenadas = Object.keys(datas).sort().reverse(); // Todas as datas, ordenadas da mais recente
  
  // Se não passou filtro, usar data mais recente
  var dataSelecionada = dataStr;
  if (!dataSelecionada && datasOrdenadas.length > 0) {
    dataSelecionada = datasOrdenadas[0];
  }
  
  if (!dataSelecionada) {
    return { datasDisponiveis: [], executores: [] };
  }
  
  // Processar dados da data selecionada
  data.forEach(function(row) {
    var dataNota = row[COL.ATESTO];
    if (!dataNota || !(dataNota instanceof Date)) return;
    
    var dataNotaStr = Utilities.formatDate(dataNota, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dataNotaStr !== dataSelecionada) return;
    
    var executor = String(row[COL.EXECUTOR] || '').trim();
    if (!executor) executor = 'SEM EXECUTOR';
    
    var agencia = String(row[COL.AGENCIA] || '').toUpperCase().trim();
    var statusRaw = String(row[COL.STATUS_PAG] || '').toLowerCase().trim();
    var valor = parseFloat(row[COL.VALOR]) || 0;
    var nf = String(row[COL.NF] || '').trim();
    var emAnalise = !statusRaw || statusRaw.indexOf('analise') !== -1;
    
    if (!executores[executor]) {
      executores[executor] = {
        nome: executor,
        av: 0, calia: 0, ebm: 0, total: 0, emAnalise: 0, valor: 0,
        nfsAv: [], nfsCalia: [], nfsEbm: []
      };
    }
    
    if (agencia.indexOf('AV') !== -1 || agencia === 'AV') {
      executores[executor].av++;
      if (nf) executores[executor].nfsAv.push(nf);
    } else if (agencia.indexOf('CALIA') !== -1) {
      executores[executor].calia++;
      if (nf) executores[executor].nfsCalia.push(nf);
    } else if (agencia.indexOf('EBM') !== -1) {
      executores[executor].ebm++;
      if (nf) executores[executor].nfsEbm.push(nf);
    }
    
    executores[executor].total++;
    executores[executor].valor += valor;
    if (emAnalise) executores[executor].emAnalise++;
  });
  
  var result = Object.values(executores);
  result.sort(function(a, b) { return b.total - a.total; });
  
  return {
    datasDisponiveis: datasOrdenadas,
    executores: result,
    dataSelecionada: dataSelecionada
  };
}

// ======================================================================
// INTEGRAÇÃO: GERADOR DE TABELA DE CONTROLE DE PAGAMENTO
// ======================================================================

const SHEET_NAME_EMPENHOS = 'SaldoEmpenhos'; 
const SHEET_NAME_CERTIDOES = 'Certidões';

function filtrarControleDePagamento(valorBusca) {
  try {
    const data = getAllData_();
    const out = [];
    const buscaStr = String(valorBusca || "").trim();

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const valorT_raw = r[COL.CONTROLE];
      let valorT = typeof valorT_raw === 'number' ? valorT_raw.toFixed(0) : String(valorT_raw || "");
      valorT = valorT.trim();
      
      if (!r[0] && !r[2] && !r[5]) continue;

      if (buscaStr === "" || valorT.includes(buscaStr)) {
        out.push({
          agencia_principal: String(r[COL.AGENCIA] || ""), 
          valor_nf_agencia: (r[6] === "" || r[6] == null) ? 0 : Number(r[6]), // Mantendo índices fixos se não mapeados
          veiculo_forn: String(r[COL.VEICULO] || ""), 
          cnpj: String(r[10] || ""), 
          tipo_midia: String(r[COL.TIPO_MIDIA] || ""), 
          valor_nf: (r[COL.VALOR] === "" || r[COL.VALOR] == null) ? 0 : Number(r[COL.VALOR]), 
          nf: String(r[COL.NF] || ""), 
          glosa: (r[14] === "" || r[14] == null) ? 0 : Number(r[14]),
        });
      }
    }
    return out;
  } catch (e) {
    throw new Error("Erro ao filtrar controle: " + e.message);
  }
}

function getSaldoByPI(valorBusca) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const buscaStr = String(valorBusca || "").trim().toUpperCase();
  const defaultReturn = { saldo: 0, agencia: null, campanha: null };
  if (!buscaStr) return defaultReturn;

  let agencia = null;
  let campanha = null;
  let saldoAtual = 0;

  try {
    const data = getAllData_();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const valorT_raw = r[COL.CONTROLE];
      let valorT = typeof valorT_raw === 'number' ? valorT_raw.toFixed(0) : String(valorT_raw || "");
      valorT = valorT.trim().toUpperCase(); 

      if (valorT.includes(buscaStr)) {
        agencia = String(r[COL.AGENCIA] || "").trim();
        campanha = String(r[COL.CAMPANHA] || "").trim();
        break; 
      }
    }

    if (!agencia || !campanha) return defaultReturn;

    const shEmpenhos = ss.getSheetByName(SHEET_NAME_EMPENHOS);
    if (!shEmpenhos) return { ...defaultReturn, agencia, campanha };
    
    const rowsEmpenhos = shEmpenhos.getDataRange().getValues();
    const buscaAgenciaUpper = agencia.toUpperCase();
    const buscaCampanhaUpper = campanha.toUpperCase();
    
    for (let i = 1; i < rowsEmpenhos.length; i++) {
      const r = rowsEmpenhos[i];
      const valorAgenciaEmpenho = String(r[5] || "").trim().toUpperCase();
      const valorCampanhaEmpenho = String(r[4] || "").trim().toUpperCase();

      if (valorAgenciaEmpenho.includes(buscaAgenciaUpper) && valorCampanhaEmpenho.includes(buscaCampanhaUpper)) {
        saldoAtual = (r[7] === "" || r[7] == null) ? 0 : Number(r[7]);
        break; 
      }
    }

    return { saldo: saldoAtual, agencia: agencia, campanha: campanha };
  } catch (e) {
    return defaultReturn; 
  }
}

function getCertidoesByAgencia(agenciaBusca) {
    const defaultReturn = { rfb: null, sefaz_df: null, fgts: null, tst: null, link: null, agencia_cert: null };
    const buscaAgencia = String(agenciaBusca || "").trim().toUpperCase(); 
    if (!buscaAgencia) return defaultReturn;
    
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sh = ss.getSheetByName(SHEET_NAME_CERTIDOES);
        if (!sh) return defaultReturn;
        
        const rows = sh.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            const agenciaCertidao = String(r[0] || "").trim().toUpperCase();
            if (agenciaCertidao.includes(buscaAgencia) || buscaAgencia.includes(agenciaCertidao)) {
                return {
                    agencia_cert: String(r[0] || ""),
                    rfb: formatSheetDate_(r[3]),
                    sefaz_df: formatSheetDate_(r[4]),
                    fgts: formatSheetDate_(r[5]),
                    tst: formatSheetDate_(r[6]),
                    link: String(r[7] || ""),
                };
            }
        }
        return defaultReturn; 
    } catch (e) {
        return defaultReturn;
    }
}

function formatSheetDate_(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    const d = dateValue.getDate().toString().padStart(2, '0');
    const m = (dateValue.getMonth() + 1).toString().padStart(2, '0');
    const y = dateValue.getFullYear().toString().slice(-2);
    return `${d}/${m}/${y}`;
  }
  return String(dateValue || 'N/A');
}
