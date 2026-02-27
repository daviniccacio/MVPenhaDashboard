function renderTabela(dados) {
  const corpo = document.getElementById("tabelaCorpo");
  corpo.innerHTML = "";

  const limparNum = (val) => {
    if (!val) return 0;
    if (typeof val === "number") return val;
    const str = val.toString().split(/\s{2,}/)[0].trim();
    return parseFloat(str.replace(/\./g, "").replace(",", ".")) || 0;
  };

  // Converte número serial do Excel para data
  const excelParaData = (val) => {
    if (!val) return "-";
    if (typeof val === "number") {
      // Excel conta dias desde 01/01/1900
      const data = new Date((val - 25569) * 86400 * 1000);
      return data.toLocaleDateString("pt-BR");
    }
    if (val instanceof Date) return val.toLocaleDateString("pt-BR");
    return val.toString().slice(0, 10);
  };

  dados.forEach((itemRaw) => {
    const item = normalizarItem(itemRaw);
    if (!item["Placa"]) return;
    if (item["Placa"].toString().toLowerCase().includes("total")) return;
    if (item["Placa"].toString().trim().toLowerCase() === "placa") return;

    const est    = item["Estabelecimento"]?.toString().replace(/Razão Social:/gi, "").trim() || "-";
    const placaRaw = (item["Placa"]?.toString().trim() || "").split(/\s+/);
    const placa = placaRaw[0] || "SEM PLACA";
    const comb = item["Combustível"]?.toString().trim() || placaRaw.slice(1).join(" ") || "OUTRO";
    const litrosAutorizados = limparNum(item["Autorizado"]);
    const litrosEfetivado = limparNum(item["Efetivado"]);
    const vlrTotalDireto = limparNum(item["Vlr Total"]);
    const vlrUnitario    = limparNum(item["Vlr Unitário"]);
    const valor  = vlrTotalDireto || (litrosEfetivado * vlrUnitario);
    const data   = excelParaData(item["Data/Hora"]);
    const motoristaRaw = item["Motorista"]?.toString().trim();
    const motorista = motoristaRaw || item["Vlr Total"]?.toString().split(/\s{2,}/)[1]?.trim() || "-";
    

    corpo.innerHTML += `
      <tr class="border-b border-slate-700/50 hover:bg-slate-700/30 transition-colors">
        <td class="px-4 py-3 text-slate-300">${data}</td>
        <td class="px-4 py-3 font-mono font-bold text-blue-400">${placa}</td>
        <td class="px-4 py-3 text-slate-300">${est}</td>
        <td class="px-4 py-3 text-slate-300">${comb}</td>
        <td class="px-4 py-3 text-slate-300">${litrosAutorizados.toLocaleString("pt-BR")} L</td>
        <td class="px-4 py-3 text-slate-300">${litrosEfetivado.toLocaleString("pt-BR")} L</td>
        <td class="px-4 py-3 text-emerald-400 font-semibold">${vlrUnitario.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</td>
        <td class="px-4 py-3 text-emerald-400 font-semibold">${valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" })}</td>
        <td class="px-4 py-3 text-slate-300">${motorista}</td>
      </tr>
    `;
  });
}

function filtrarTabela(){
  const busca = document.getElementById("campoBusca").value.toLowerCase().trim();
  const linhas = document.querySelectorAll("#tabelaCorpo tr");

  linhas.forEach((tr) => {
    const texto = tr.innerText.toLowerCase();
    tr.style.display = texto.includes(busca) ? "" : "none";
  });
}

let ordemAsc = true;
let colunaAtual = -1;

function ordenarTabela(colIndex) {
  const tbody = document.getElementById("tabelaCorpo");
  const linhas = Array.from(tbody.querySelectorAll("tr:not([style*='display: none'])"));
  // Inclui também as ocultas pelo filtro
  const todasLinhas = Array.from(tbody.querySelectorAll("tr"));

  // Alterna direção se mesma coluna, senão reseta
  if (colunaAtual === colIndex) {
    ordemAsc = !ordemAsc;
  } else {
    ordemAsc = true;
    colunaAtual = colIndex;
  }

  todasLinhas.sort((a, b) => {
    const tdA = a.querySelectorAll("td")[colIndex];
    const tdB = b.querySelectorAll("td")[colIndex];
    const valA = tdA?.innerText.trim() || "";
    const valB = tdB?.innerText.trim() || "";

    // Tenta converter pra número — remove R$, L, pontos de milhar
    const limpar = (v) => parseFloat(v.replace(/[R$\sL.]/g, "").replace(",", "."));
    const numA = limpar(valA);
    const numB = limpar(valB);

    if (!isNaN(numA) && !isNaN(numB)) {
      return ordemAsc ? numA - numB : numB - numA;
    }

    // Tenta ordenar por data dd/mm/aaaa
    const toDate = (v) => {
      const [d, m, y] = v.split("/");
      return y && m && d ? new Date(`${y}-${m}-${d}`) : null;
    };
    const dateA = toDate(valA);
    const dateB = toDate(valB);
    if (dateA && dateB) {
      return ordemAsc ? dateA - dateB : dateB - dateA;
    }

    // Texto normal
    return ordemAsc
      ? valA.localeCompare(valB, "pt-BR")
      : valB.localeCompare(valA, "pt-BR");
  });

  // Reinserve ordenado
  todasLinhas.forEach((linha) => tbody.appendChild(linha));

  document.querySelectorAll("thead th").forEach((th, i) => {
  const label = th.textContent.replace(/\s*[↕↑↓]\s*$/, "").trim();
  if (i === colunaAtual) {
    th.textContent = `${label} ${ordemAsc ? "↑" : "↓"}`;
  } else {
    th.textContent = `${label} ↕`;
  }
});

}