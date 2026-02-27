let charts = {};

document.addEventListener("DOMContentLoaded", () => {
  const btnLoad = document.getElementById("btnLoad"); 

  renderRecentes();

  btnLoad.addEventListener("click", async () => {
    try{
      const response = await window.electronAPI.openExcel();
      if (!response) return;
  
      if(!response.data || response.data.length === 0){
        mostrarErro("A planilha está vazia ou não foi possível ler os dados");
        return;
      }
  
      const primeiraLinha = normalizarItem(response.data[0]);
      if(!primeiraLinha["Placa"] && !primeiraLinha["Estabelecimento"]){
        mostrarErro("Formato de planilha não reconhecido. Verifique se o arquivo está correto");
        return;
      }

      const { data, fileName } = response;
      document.getElementById("fileName").innerText = fileName;
      processarDados(data);
    } catch(error){
      console.error("Erro ao carregar o Excel:", error);
      mostrarErro("Ocorreu um erro inesperado ao carregar a planilha");
    }
  });

  renderRecentes();

});

function processarDados(dadosRaw) {
  let dados = dadosRaw[0] && normalizarItem(dadosRaw[0])["Placa"]?.toString().trim() === "Placa" ? dadosRaw.slice(1) : dadosRaw;

  if(dados[0] && normalizarItem(dados[0])["Placa"]?.toString().trim() === "Placa"){
    dados = dados.slice(1);
  };

  let totalGeral       = 0;
  let totalLitros      = 0;
  let combustiveis     = {};
  let estabelecimentos = {};
  let veiculos         = {};
  let autorizado       = 0;
  let efetivado        = 0;

  const limparNum = (val) => {
    if (!val) return 0;
    // Se já é número (xlsx converte automaticamente), retorna direto
    if (typeof val === "number") return val;
    // Remove motorista que vem junto ex: "1.158,00  MARCOS"
    const str = val.toString().split(/\s{2,}/)[0].trim();
    // Remove pontos de milhar, troca vírgula por ponto decimal
    return parseFloat(str.replace(/\./g, "").replace(",", ".")) || 0;
  };

  dados.forEach((itemRaw) => {
    const item = normalizarItem(itemRaw);

    // Pula linhas sem placa ou linha de totais
    if (!item["Placa"]) return;
    if (item["Placa"].toString().trim().toLowerCase() === "placa") return;
    if (item["Placa"].toString().toLowerCase().includes("total")) return;

    // Limpa nome do estabelecimento
    const estBruto = item["Estabelecimento"]?.toString() || "NÃO INFORMADO";
    const est = estBruto.replace(/Razão Social:/gi, "").trim().toUpperCase();

    const placaRaw = item["Placa"]?.toString().trim().split(/\s+/) || [];
    const placa = placaRaw[0] || "SEM PLACA";
    const comb = item["Combustível"]?.toString().trim() || placaRaw.slice(1).join(" ") || "OUTRO";

    const litros         = limparNum(item["Efetivado"]);
    const vlrUnitario    = limparNum(item["Vlr Unitário"]);
    const vlrAutorizado  = limparNum(item["Autorizado"]);
    const vlrTotalDireto = limparNum(item["Vlr Total"]);

    // Usa Vlr Total direto, senão calcula litros × preço unitário
    const vlrFinal = vlrTotalDireto || (litros * vlrUnitario);

    totalGeral  += vlrFinal;
    totalLitros += litros;
    autorizado  += vlrAutorizado;
    efetivado   += litros;

    if (litros > 0) {
      combustiveis[comb] = (combustiveis[comb] || 0) + litros;
    }
    if (vlrFinal > 0) {
      estabelecimentos[est] = (estabelecimentos[est] || 0) + vlrFinal;
      veiculos[placa]       = (veiculos[placa]       || 0) + vlrFinal;
    }
  });

  console.log("Total de linhas:", dados.length);
  console.log("Primeira linha:", JSON.stringify(normalizarItem(dados[0])));
  console.log("Segunda linha:", JSON.stringify(normalizarItem(dados[1])));

  // Atualiza os cards
  const media = totalLitros > 0 ? totalGeral / totalLitros : 0;

  animarContador("totalGeral", totalGeral, "moeda");
  animarContador("totalLitros", totalLitros, "litros");
  const totalLinhasValidas = dados.filter((itemRaw) => {
    const item = normalizarItem(itemRaw);
    return item["Placa"] && !item["Placa"].toString().toLowerCase().includes("total") && item["Placa"].toString().toLowerCase() !== "placa";
  }).length;
  animarContador("totalRegistros", totalLinhasValidas, "numeros")
  animarContador("mediaUnit", media, "moeda");

  Chart.defaults.color       = "#e2e8f0";
  Chart.defaults.borderColor = "rgba(255, 255, 255, 0.1)";
  Chart.defaults.font.family = "'Inter'";

  renderPizza("chartCombustivel", combustiveis);
  renderBarras("chartEstabelecimento", estabelecimentos, "Gasto por Estabelecimento (R$)");
  renderBarras("chartVeiculos", veiculos, "Gasto por Veículo (R$)", true, 10);
  renderComparativo("chartComparativo", autorizado, efetivado);

  renderTabela(dados);

  const btnExportar = document.getElementById("btnExportar");
  btnExportar.disabled = false;
  btnExportar.classList.add("bg-blue-700", "hover:bg-blue-600", "cursor-pointer");
  btnExportar.classList.remove("bg-slate-700");

  const nomeArquivo = document.getElementById("fileName").innerText;
  window.electronAPI.setTitle(`MV Penha - ${nomeArquivo}`);

  document.querySelectorAll(".chartPlaceholder").forEach(el => el.classList.add("hidden"));
}

async function renderRecentes() {
  const lista = document.getElementById("listaRecentes");
  if(!lista) return;

  const recentes = await window.electronAPI.getHistorico();

  if(recentes.length === 0){
    lista.innerHTML = `<p class="text-slate-500 text-xs italic">Nenhum arquivo recente</p>`;
    return;
  }

  lista.innerHTML = recentes.map(r => `
    <button onclick="abrirRecente('${r.path.replace(/\\/g, "\\\\")}')"
      class="w-full text-left px-3 py-2 rounded-xl hover:bg-slate-700 transition-colors group">
      <p class="text-xs font-medium text-slate-300 group-hover:text-white truncate">${r.name}</p>
      <p class="text-[10px] text-slate-500 mt-0.5">${r.data}</p>
    </button>
  `).join("");
}

async function abrirRecente(filePath){
  const response = await window.electronAPI.openExcelPath(filePath);
  if(!response){
    mostrarErro("Arquivo não encontrado. Ele pode ter sido movido ou excluído.");
    return;
  }
  document.getElementById("fileName").innerText = response.fileName;
  processarDados(response.data);
}
