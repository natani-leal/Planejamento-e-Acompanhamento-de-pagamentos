/**
 * Code.gs - PROJETO FOCO: CONTROLE DE SALDO ORÇAMENTÁRIO
 * * Este código extrai dados da planilha "Consolidado", "SaldoEmpenhos" e "Certidões"
 * para alimentar o App Web.
 */

const URL_PLANILHA = 'https://docs.google.com/spreadsheets/d/15ZflxC82kDMcK8k4CbZQmX_fO9wDUl86tGYlJ3IMiT4/edit';
const SHEET_NAME = 'Consolidado';
const SHEET_NAME_EMPENHOS = 'SaldoEmpenhos'; // Nome da aba de Saldo/Empenhos
const SHEET_NAME_CERTIDOES = 'Certidões'; // Nome da aba de Certidões

// Função utilitária para formatar a data como DD/MM/YY no Apps Script
function formatSheetDate(dateValue) {
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    const d = dateValue.getDate().toString().padStart(2, '0');
    const m = (dateValue.getMonth() + 1).toString().padStart(2, '0');
    const y = dateValue.getFullYear().toString().slice(-2);
    return `${d}/${m}/${y}`;
  }
  return String(dateValue || 'N/A');
}

// Função principal de entrada do App Web
function doGet() {
  return HtmlService.createHtmlOutputFromFile('controleDepagamento')
    .setTitle('Controle de Saldo Orçamentário')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------------------
// Função auxiliar para abrir a planilha
// ----------------------------------------------------------------------
function getSpreadsheet() {
  // Extrai o ID da URL
  const match = URL_PLANILHA.match(/\/d\/([a-zA-Z0-9_-]+)/);
  const spreadsheetId = match ? match[1] : URL_PLANILHA;
  return SpreadsheetApp.openById(spreadsheetId);
}


// ======================================================================
// 1. FUNÇÃO DE BUSCA DA ABA CONSOLIDADO (USADA PELA BUSCA SEI)
// ======================================================================

function filtrarControleDePagamento(valorBusca) {
  try {
    const ss = getSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    
    if (!sh) {
       throw new Error(`A aba com o nome "${SHEET_NAME}" não foi encontrada.`);
    }
    
    // Lê a faixa de dados
    const rows = sh.getDataRange().getValues();
    const out = [];

    // CRÍTICO: Remoção do .toUpperCase() e tratamento de string para busca
    const buscaStr = String(valorBusca || "").trim();

    // Itera pelas linhas (começa da segunda linha, índice 1)
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      
      // Coluna T (índice 19) - Critério de filtro (Nº SEI/PI)
      const valorT_raw = r[19];
      let valorT = "";

      // Trata números como texto (Nº SEI/PI)
      if (typeof valorT_raw === 'number') {
          valorT = valorT_raw.toFixed(0); 
      } else {
          valorT = String(valorT_raw || "");
      }
      
      valorT = valorT.trim();
      
      // Linhas vazias são ignoradas
      if (!r[0] && !r[1] && !r[4]) continue;

      // CRITÉRIO DE FILTRO: Busca por Inclusão (Valor T e Busca Str limpas)
      if (buscaStr === "" || valorT.includes(buscaStr)) {
        
        out.push({
          // Coluna A (0) - Agência Principal (Chave para Certidão)
          agencia_principal: String(r[0] || ""), 
          
          // Coluna G (6) - Valor Nota Agência
          valor_nf_agencia: (r[6] === "" || r[6] == null) ? 0 : Number(r[6]),
          
          // Coluna J (9) - Veículo/Forncededor (Mantido para a Tabela de NFs)
          veiculo_forn: String(r[9] || ""), 
          
          // Coluna K (10) - CNPJ 
          cnpj: String(r[10] || ""), 
          
          // Coluna L (11) - Tipo de Mídia 
          tipo_midia: String(r[11] || ""), 
          
          // Coluna AB (27) - Valor Realmente Pago (Montante Total)
          valor_nf: (r[27] === "" || r[27] == null) ? 0 : Number(r[27]), 
          
          // Coluna F (5) - NF
          nf: String(r[5] || ""), 

          // Coluna O (14) - Glosa
          glosa: (r[14] === "" || r[14] == null) ? 0 : Number(r[14]),
        });
      }
    }

    Logger.log(`[Consolidado] Total de linhas encontradas para o SEI/PI (${buscaStr}): ${out.length}`);
    return out;
  } catch (e) {
    Logger.log("Erro ao processar planilha (Consolidado): " + e.message);
    throw new Error("Erro ao acessar a aba 'Consolidado' ou nome da aba incorreto. Detalhes: " + e.message);
  }
}


// ======================================================================
// 2. FUNÇÃO PARA BUSCAR O SALDO CRUZANDO DADOS COM O PI
// ======================================================================

/**
 * Busca Agência (Col. A) e Campanha (Col. C) na Consolidado usando o PI (Col. T),
 * e então usa esses dados para buscar o Saldo Atual (Col. H) em SaldoEmpenhos.
 * * Retorna um objeto contendo o saldo, agência e campanha.
 */
function getSaldoByPI(valorBusca) {
  const ss = getSpreadsheet();
  // Mantido em maiúsculas para cruzar com a aba SaldoEmpenhos
  const buscaStr = String(valorBusca || "").trim().toUpperCase();

  // Retorno padrão se a busca for vazia
  const defaultReturn = { saldo: 0, agencia: null, campanha: null };
  if (!buscaStr) return defaultReturn;

  let agencia = null;
  let campanha = null;
  let saldoAtual = 0;

  // 1. BUSCAR AGÊNCIA (A) E CAMPANHA (C) NA ABA CONSOLIDADO
  try {
    const shConsolidado = ss.getSheetByName(SHEET_NAME);
    if (!shConsolidado) throw new Error(`A aba "${SHEET_NAME}" não foi encontrada.`);
    
    const rowsConsolidado = shConsolidado.getDataRange().getValues();

    for (let i = 1; i < rowsConsolidado.length; i++) {
      const r = rowsConsolidado[i];
      // Coluna T (índice 19) - Critério de filtro (Nº SEI/PI)
      const valorT_raw = r[19];
      let valorT = typeof valorT_raw === 'number' ? valorT_raw.toFixed(0) : String(valorT_raw || "");
      // Limpa e compara
      valorT = valorT.trim().toUpperCase(); 

      // Busca por Inclusão (Para encontrar o primeiro match e obter Agência/Campanha)
      if (valorT.includes(buscaStr)) {
        // Coluna A (índice 0) - Agência
        agencia = String(r[0] || "").trim();
        // Coluna C (índice 2) - Campanha
        campanha = String(r[2] || "").trim();
        break; 
      }
    }

    if (!agencia || !campanha) {
        return defaultReturn;
    }

    // 2. BUSCAR SALDO ATUAL (H) NA ABA SALDOEMPENHOS
    const shEmpenhos = ss.getSheetByName(SHEET_NAME_EMPENHOS);
    if (!shEmpenhos) throw new Error(`A aba "${SHEET_NAME_EMPENHOS}" não foi encontrada.`);
    
    const rowsEmpenhos = shEmpenhos.getDataRange().getValues();
    
    // Converte e limpa para comparação
    const buscaAgenciaUpper = agencia.toUpperCase();
    const buscaCampanhaUpper = campanha.toUpperCase();
    
    for (let i = 1; i < rowsEmpenhos.length; i++) {
      const r = rowsEmpenhos[i];
      
      // Coluna F (5) - Agência na SaldoEmpenhos
      const valorAgenciaEmpenho = String(r[5] || "").trim().toUpperCase();
      
      // Coluna E (4) - Campanha na SaldoEmpenhos
      const valorCampanhaEmpenho = String(r[4] || "").trim().toUpperCase();

      // CRITÉRIO DE CRUZAMENTO: Busca por Inclusão
      if (valorAgenciaEmpenho.includes(buscaAgenciaUpper) && valorCampanhaEmpenho.includes(buscaCampanhaUpper)) {
        
        // Coluna H (7) - Saldo Atual
        const valorBrutoH = (r[7] === "" || r[7] == null) ? 0 : Number(r[7]);
        
        // SOLUÇÃO EPSILON: Trata erro de precisão de ponto flutuante
        const EPSILON = 0.000001; 

        if (Math.abs(valorBrutoH) < EPSILON) {
            saldoAtual = 0;
        } else {
            // Mantém a precisão total
            saldoAtual = valorBrutoH;
        }
        
        break; 
      }
    }

    return { 
        saldo: saldoAtual, 
        agencia: agencia, 
        campanha: campanha 
    };
    
  } catch (e) {
    Logger.log("Erro ao buscar Saldo por PI: " + e.message);
    return defaultReturn; 
  }
}


// ======================================================================
// 3. FUNÇÃO PARA BUSCAR CERTIDÕES POR AGÊNCIA
// ======================================================================

function getCertidoesByAgencia(agenciaBusca) {
    const defaultReturn = { 
        rfb: null, 
        sefaz_df: null, 
        fgts: null, 
        tst: null, 
        link: null, 
        agencia_cert: null 
    };
    
    const buscaAgencia = String(agenciaBusca || "").trim().toUpperCase(); 
    if (!buscaAgencia) return defaultReturn;
    
    try {
        const ss = getSpreadsheet();
        const sh = ss.getSheetByName(SHEET_NAME_CERTIDOES);

        if (!sh) {
            Logger.log(`A aba com o nome "${SHEET_NAME_CERTIDOES}" não foi encontrada.`);
            return defaultReturn;
        }
        
        const rows = sh.getDataRange().getValues();

        for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            
            // Coluna A (0) - Agência na Certidões
            const agenciaCertidao = String(r[0] || "").trim().toUpperCase();
            
            if (agenciaCertidao.includes(buscaAgencia) || buscaAgencia.includes(agenciaCertidao)) {
                
                return {
                    agencia_cert: String(r[0] || ""),
                    // CRÍTICO: Formatação de data aplicada aqui
                    rfb: formatSheetDate(r[3]),
                    sefaz_df: formatSheetDate(r[4]),
                    fgts: formatSheetDate(r[5]),
                    tst: formatSheetDate(r[6]),
                    link: String(r[7] || ""),
                };
            }
        }

        return defaultReturn; 
        
    } catch (e) {
        Logger.log("Erro ao buscar Certidões por Agência: " + e.message);
        return defaultReturn;
    }
}
