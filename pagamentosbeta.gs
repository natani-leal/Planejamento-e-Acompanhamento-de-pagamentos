<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <title>Controle de Saldo Orçamentário</title>
  <style>
    :root {
      --color-primary: #007bff;
      --color-secondary: #0056b3;
      --color-maroon: #800020;
    }

    body { 
      font-family: Calibri, sans-serif; 
      padding: 20px;
      background-color: white;
      display: flex;
      flex-direction: column;
      align-items: flex-start; 
    }

    body, input, th, td, p, span, a, div {
      font-family: Calibri, sans-serif !important;
      font-size: 1em;
      color: #333;
    }

    .container-central {
      width: 100%;
      max-width: 1200px;
      background-color: white;
      padding: 0;
    }

    .header-title {
      text-align: left;
      font-weight: bold;
      margin: 0 0 20px 0;
      font-size: 1.1em;
      color: #444;
      border-bottom: 2px solid #eee;
      padding-bottom: 10px;
      width: 100%;
    }

    /* BUSCA */
    .search-container {
      width: 100%;
      margin-bottom: 20px;
    }

    .search-box-row-main {
      display: flex;
      align-items: center;
      gap: 15px;
      flex-wrap: wrap;
    }

    .search-input-group input {
      padding: 10px 15px;
      width: 380px;
      border: 1px solid #ccc;
      border-radius: 6px;
      transition: border-color .3s, box-shadow .3s;
    }

    .search-input-group input:focus {
      border-color: var(--color-primary);
      box-shadow: 0 0 0 3px rgba(0, 123, 255, .25);
      outline: none;
    }

    .copy-button-small {
      background: none;
      border: none;
      cursor: pointer;
      display: flex;
      padding: 0;
    }

    .copy-button-small svg {
      width: 18px;
      height: 18px;
      fill: var(--color-primary);
    }

    .copy-button-small:hover svg {
      fill: var(--color-secondary);
    }

    /* STATUS SALDO */
    #statusSaldo {
      margin-top: 10px;
    }

    .info-saldo-text {
      font-weight: bold;
      color: var(--color-secondary);
    }

    /* CERTIDÕES */
    #certidoes-area {
      width: 100%;
      margin-top: 15px;
    }

    .certidao-item {
      display: inline-block;
      margin-right: 12px;
      padding: 3px 6px;
      background: #f5f5f5;
      border-radius: 4px;
    }

    .certidao-nome {
      font-weight: bold;
      margin-right: 4px;
    }

    .certidao-link-icon {
      margin-left: 10px;
      font-size: 1.1em;
      text-decoration: none;
      color: var(--color-primary);
      font-weight: bold;
      cursor: pointer;
    }

    .certidao-link-icon:hover {
      color: var(--color-secondary);
    }

    .warning-producao {
      margin-top: 8px;
      padding-left: 5px;
      border-left: 3px solid var(--color-maroon);
      font-size: .95em;
    }

    /* ATESTO */
    #atesto-info-container {
      margin-top: 25px;
      margin-bottom: 25px;
    }

    .atesto-values {
      display: inline-flex;
      gap: 6px;
      align-items: center;
      font-size: 1em;
      white-space: nowrap;
    }

    /* POPUP COPIAR */
    #copy-popup {
      background-color: var(--color-secondary);
      color: white;
      font-weight: bold;
      position: fixed;
      bottom: 22px;
      right: 22px;
      padding: 10px 20px;
      border-radius: 5px;
      font-size: .9em;
      display: none;
    }

    /* TABELA */
    .result-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 30px;
      background-color: white;
    }

    .result-table th,
    .result-table td {
      border: 1px solid #000;
      padding: 10px;
      text-align: center;
      vertical-align: middle;
      background-color: white !important;
      white-space: nowrap;
    }

    /* COLUNAS */
    .valor-cell { text-align: center !important; }

    .glosa-cell,
    .glosa-total {
      color: var(--color-maroon);
      font-weight: bold;
    }

    .total-negrito {
      font-weight: bold !important;
    }

    .negative-balance {
      color: var(--color-maroon) !important;
      font-weight: bold;
    }

    /* CARREGANDO */
    #loading {
      color: #444;
      font-weight: bold;
      margin-top: 10px;
    }
  </style>
</head>
<body>

<div class="container-central">
  <div class="search-container">
    <div class="search-box-row-main">
      <div class="search-input-group">
        <input type="text" id="valorBusca"
          oninput="debounceSearch();"
          placeholder="Digite o nº SEI do Controle de pagamento"
          inputmode="numeric">
      </div>

      <!-- BOTÃO COPIAR ATESTO -->
      <button class="copy-button-small"
        onclick="
          let texto = document.getElementById('nfList')?.textContent || '';
          texto = texto.replace('Texto para atesto:', '').trim();
          copiarParaClipboard(texto);
        "
        title="Copiar texto para atesto">
        <svg viewBox='0 0 24 24'>
          <path d='M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2
            v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z'/>
        </svg>
      </button>
    </div>

    <div id="statusSaldo"></div>
    <div id="certidoes-area"></div>
  </div>

  <!-- ATESTO GERADO AUTOMATICAMENTE -->
  <div id="atesto-info-container"></div>

  <p id="loading" style="display:none;">Carregando dados...</p>

  <!-- CONTAINER DA TABELA -->
  <div id="table-container"></div>
</div>

<!-- POPUP -->
<div id="copy-popup">Copiado!</div>

<script>
/* VARIÁVEIS GLOBAIS */
let searchTimeout = null;
window.totalValorPago = 0;
window.saldoEmpenhoAtual = 0;

/* ============================
   COPIAR / POPUP
============================ */
function showCopyPopup() {
  const popup = document.getElementById("copy-popup");
  popup.style.display = "block";
  popup.style.opacity = "1";

  setTimeout(() => {
    popup.style.opacity = "0";
    setTimeout(() => (popup.style.display = "none"), 300);
  }, 1200);
}

function copiarParaClipboard(texto) {
  navigator.clipboard.writeText(texto)
    .then(showCopyPopup)
    .catch(() => alert("Erro ao copiar automaticamente."));
}

/* ============================
   FORMATAR VALOR
============================ */
function formatBRL(v) {
  const n = Number(v);
  if (isNaN(n)) return "-";
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

/* ============================
   DEBOUNCE DA BUSCA
============================ */
function debounceSearch() {
  clearTimeout(searchTimeout);

  searchTimeout = setTimeout(() => {
    const valor = document.getElementById("valorBusca").value.trim();

    if (!valor) {
      document.getElementById("table-container").innerHTML = "";
      document.getElementById("certidoes-area").innerHTML = "";
      document.getElementById("atesto-info-container").innerHTML = "";
      document.getElementById("statusSaldo").innerHTML = "";
      return;
    }

    buscarDados(valor);
    buscarSaldo(valor);
  }, 300);
}

/* ============================
   BUSCAR SALDO
============================ */
function buscarSaldo(valor) {
  google.script.run
    .withSuccessHandler(updateSaldo)
    .withFailureHandler(handleError)
    .getSaldoByPI(valor);
}

function updateSaldo(result) {
  const div = document.getElementById("statusSaldo");

  if (result?.agencia && result?.campanha) {
    window.saldoEmpenhoAtual = Number(result.saldo || 0);

    div.innerHTML = `
      Agência: <span class="info-saldo-text">${result.agencia}</span>,
      Campanha: <span class="info-saldo-text">${result.campanha}</span>,
      Saldo Atual: <span class="info-saldo-text">${formatBRL(window.saldoEmpenhoAtual)}</span>
    `;
  } else {
    div.innerHTML = `<span style="color:var(--color-maroon);font-weight:bold;">Agência/Campanha não encontradas.</span>`;
    window.saldoEmpenhoAtual = 0;
  }

  calcularSaldoFinal();
}

/* ============================
   CALCULAR SALDO FINAL
============================ */
function calcularSaldoFinal() {
  const cell = document.getElementById("saldoFinalCell");
  if (!cell) return;

  const saldoFinal = (window.saldoEmpenhoAtual || 0) - (window.totalValorPago || 0);

  cell.textContent = formatBRL(saldoFinal);

  if (saldoFinal < 0) {
    cell.classList.add("negative-balance");
  } else {
    cell.classList.remove("negative-balance");
  }
}

/* ============================
   ERRO
============================ */
function handleError(error) {
  document.getElementById("loading").style.display = "none";
  document.getElementById("table-container").innerHTML =
    `<p style="color:var(--color-maroon);font-weight:bold;">ERRO: ${error.message || error}</p>`;
}

/* ============================
   BUSCAR DADOS DA TABELA
============================ */
function buscarDados(valor) {
  document.getElementById("loading").style.display = "block";
  document.getElementById("table-container").innerHTML = "";

  google.script.run
    .withSuccessHandler(exibirTabela)
    .withFailureHandler(handleError)
    .filtrarControleDePagamento(valor);
}

/* ============================
   EXIBIR TABELA PRINCIPAL
============================ */
function exibirTabela(data) {
  document.getElementById("loading").style.display = "none";
  const container = document.getElementById("table-container");

  if (!data || data.length === 0) {
    container.innerHTML =
      `<p style="color:var(--color-maroon);">Nenhuma Nota Fiscal encontrada.</p>`;
    document.getElementById("certidoes-area").innerHTML = "";
    document.getElementById("atesto-info-container").innerHTML = "";
    return;
  }

  /* --- BUSCA DE CERTIDÕES --- */
  const agencia = data[0].agencia_principal;
  const tiposMidia = data.map(r => r.tipo_midia);
  const veiculos = data.map(r => r.veiculo_forn);

  google.script.run
    .withSuccessHandler(d => exibirCertidoes(d, tiposMidia))
    .withFailureHandler(handleError)
    .getCertidoesByAgencia(agencia);

  /* --- PROCESSAMENTO DE VALORES --- */
  let totalNF = 0;
  let totalGlosa = 0;
  let totalLiquido = 0;

  let nfList = [];

  data.forEach(r => {
    const nf = Number(r.valor_nf_agencia) || 0;
    const gl = Number(r.glosa) || 0;
    const liq = Number(r.valor_nf) || 0;

    totalNF += nf;
    totalGlosa += gl;
    totalLiquido += liq;

    if (r.nf && r.nf !== "-") nfList.push(r.nf);
  });

  window.totalValorPago = totalLiquido;

  nfList = [...new Set(nfList)].sort((a, b) => Number(a) - Number(b));

  const nfTexto =
    nfList.length > 1
      ? nfList.slice(0, -1).join(", ") + " e " + nfList[nfList.length - 1]
      : nfList[0];

  const totalFormatado = formatBRL(totalLiquido);

  /* --- TEXTO PARA ATESTO --- */
  document.getElementById("atesto-info-container").innerHTML = `
    <div id="atesto-content-container">
      <div class="atesto-values" id="nfList">
        <span style="font-weight:bold;color:var(--color-secondary);">Texto para atesto:</span>
        <span class="atesto-nf-numbers">${nfTexto}</span>,
        no montante total de <span style="font-weight:bold;">${totalFormatado}</span>

        <button class="copy-button-small"
          onclick="
            let full = document.getElementById('nfList').textContent.trim();
            let clean = full.replace('Texto para atesto:', '').trim();
            copiarParaClipboard(clean);
          "
          title="Copiar texto do atesto">
          <svg viewBox='0 0 24 24'>
            <path d='M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2
                     v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z'/>
          </svg>
        </button>
      </div>
    </div>
  `;

  /* --- TABELA --- */
  const temGlosa = totalGlosa > 0;

  let html = `
    <table id="relatorio-tabela" class="result-table">
      <thead>
        <tr>
          <td colspan="${temGlosa ? 9 : 7}"
            style="text-align:left; font-weight:bold; padding:10px;">
            4. CONTROLE DE SALDO ORÇAMENTÁRIO
          </td>
        </tr>

        <tr>
          <th>Ord.</th>
          <th>NF</th>
          <th>Veículo/Fornecedor</th>
          <th>CNPJ</th>
          <th>Tipo de Mídia</th>
          <th>Valor da NF</th>
          ${temGlosa ? "<th>Glosa</th><th>Valor após Glosa</th>" : ""}
          <th>Saldo do Empenho</th>
        </tr>
      </thead>
      <tbody>
  `;

  /* --- LINHAS DA TABELA --- */
  data.forEach((r, i) => {
    const nf = Number(r.valor_nf_agencia) || 0;
    const gl = Number(r.glosa) || 0;
    const liq = Number(r.valor_nf) || 0;

    html += `
      <tr>
        <td>${i + 1}</td>
        <td>${r.nf || "-"}</td>
        <td>${r.veiculo_forn || "-"}</td>
        <td>${r.cnpj || "-"}</td>
        <td>${r.tipo_midia || "-"}</td>
        <td class="valor-cell">${formatBRL(nf)}</td>
    `;

    if (temGlosa) {
      html += `
        <td class="valor-cell glosa-cell">
          ${gl > 0 ? formatBRL(gl) : "-"}
        </td>
        <td class="valor-cell ${gl > 0 ? "total-negrito" : ""}">
          ${formatBRL(liq)}
        </td>
      `;
    }

    html += `<td class="valor-cell">-</td></tr>`;
  });

  /* --- TOTAL --- */
  html += `
    <tr class="total-row">
      <td colspan="5" style="text-align:right; font-weight:bold; padding-right:15px;">
        TOTAL
      </td>

      <td class="valor-cell ${!temGlosa ? "total-negrito" : ""}">
        ${formatBRL(totalNF)}
      </td>
  `;

  if (temGlosa) {
    html += `
      <td class="valor-cell glosa-total">${formatBRL(totalGlosa)}</td>
      <td class="valor-cell total-negrito">${formatBRL(totalLiquido)}</td>
    `;
  }

  html += `
      <td id="saldoFinalCell" class="valor-cell">-</td>
    </tr>
    </tbody>
    </table>
  `;

  container.innerHTML = html;

  calcularSaldoFinal();
}

/* ============================
   CERTIDÕES
============================ */
function exibirCertidoes(data, tiposMidia) {
  const container = document.getElementById("certidoes-area");
  container.innerHTML = "";

  if (!data || !data.agencia_cert) {
    container.innerHTML =
      `<p style="color:var(--color-maroon);font-weight:bold;">Certidões: Informação não encontrada.</p>`;
    return;
  }

  const lista = [
    { nome: "RFB", data: data.rfb },
    { nome: "SEFAZ DF", data: data.sefaz_df },
    { nome: "FGTS", data: data.fgts },
    { nome: "TST", data: data.tst }
  ];

  let html = `
    <strong>Certidões:</strong>
    ${lista
      .map(c => `<span class="certidao-nome">${c.nome}</span> ${c.data}`)
      .join(", ")}
  `;

  if (data.link) {
    html += `
      <a href="${data.link}" target="_blank" class="certidao-link-icon">🔗</a>
    `;
  }

  if (tiposMidia.some(t => String(t).toUpperCase().includes("PRODUÇÃO"))) {
    html += `
      <p class="warning-producao">
        lembre-se de verificar validade das certidões das NFs de produção
      </p>
    `;
  }

  container.innerHTML = html;
}
</script>

</body>
</html>
