function normalizarItem(itemRaw) {
  const mapa = {
    "combustivel":          "Combustível",
    "combustível":          "Combustível",
    "vir total motorista":  "Vlr Total",
    "vlr total motorista":  "Vlr Total",
    "vlr total":            "Vlr Total",
    "vir combust.":         "Vlr Unitário",
    "vlr unitário":         "Vlr Unitário",
    "vlr unitario":         "Vlr Unitário",
    "hora":                 "Data/Hora",
  };

  return Object.fromEntries(
    Object.entries(itemRaw).map(([k, v]) => {
      const chaveNorm = mapa[k.trim().toLowerCase()] || k.trim();
      return [chaveNorm, v];
    })
  );
}

function animarContador(elementId, valorFinal, formato = "numero"){
  const el = document.getElementById(elementId);
  const duracao = 1500;
  const steps = 60;
  const intervalo = duracao/steps;
  let step = 0;

  const timer = setInterval(() => {
    step++;
    const progresso = step / steps;
    const ease = 1 - Math.pow(1 - progresso, 4);
    const valorAtual = valorFinal * ease;

    if(formato === "moeda"){
      el.innerText = valorAtual.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
    } else if(formato === "litros"){
      el.innerText = `${Math.round(valorAtual).toLocaleString("pt-BR")} L`;
    } else {
      el.innerText = Math.round(valorAtual);
    }

    if(step >= steps) clearInterval(timer);
  }, intervalo);
}