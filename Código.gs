// ==========================================
// ARQUIVO: Code.gs (Backend)
// ==========================================

// Configuração
var SPREADSHEET_ID = '15ZflxC82kDMcK8k4CbZQmX_fO9wDUl86tGYlJ3IMiT4';
var ABA_CONSOLIDADO = 'Consolidado';
// NOVO LINK ATUALIZADO:
var LINK_DIVISAO_EXECUTORES = 'https://script.google.com/macros/s/AKfycbzAXBIR5r5C-8FmgQws4YzL1jzEbNKNBFxLQQ6WjndJL5TeEO-vh3fjp2Dle_Efo3WOzA/exec';

// Colunas (0-indexed)
// Confirmed mapping: L is 11 (A=0, ..., L=11)
var COL = {
  AGENCIA: 0,      // A
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
    
    if (isAtestada || isInconformidade) {
      var groupKey = row[COL.AGENCIA] + '|' + row[COL.CAMPANHA];
      if (!grupos[groupKey]) {
        var saved = stored[groupKey] || {};
        grupos[groupKey] = {
          agencia: row[COL.AGENCIA],
          campanha: row[COL.CAMPANHA],
          qtdAtestada: 0,
          valorAtestado: 0,
          qtdPendente: 0,
          valorPendente: 0,
          pendentes: [],
          obs: saved.obs || ''
        };
      }
      var valor = parseFloat(row[COL.VALOR]) || 0;
      if (isAtestada && !isInconformidade) {
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
    var dataNota = row[COL.ATESTO]; // Usando data de atesto
    if (dataNota && dataNota instanceof Date) {
      var key = Utilities.formatDate(dataNota, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (key) datas[key] = true;
    }
  });
  
  var datasOrdenadas = Object.keys(datas).sort().reverse().slice(0, 60);
  
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
