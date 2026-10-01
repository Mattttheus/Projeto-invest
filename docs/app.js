/* Gestão de Investimentos — página web gerada por gestao.py (--web) ou demo/gerar_demo.py.
 *
 * dados.json traz a base (ativos, lançamentos, proventos, metas, premissas) e a análise quantitativa já
 * calculada em Python. Carteira, Meta, Painel e Projeção são recalculados aqui com as mesmas fórmulas da
 * planilha, para que metas, premissas e lançamentos possam mudar na própria página.
 * modo "local" (WAMP): gravações vão para api.php -> gestao.py --web-registrar (mesmas regras do Cadastro).
 * modo "demo" (GitHub Pages): nada é gravado no servidor; as alterações ficam no localStorage do navegador.
 */
"use strict";

const ABAS = [["painel", "Painel"], ["carteira", "Carteira"], ["meta", "Meta"], ["quant", "Quant"], ["modelos", "Modelos"],
  ["projecao", "Projeção"], ["lancar", "✚ Lançar"], ["lancamentos", "Lançamentos"], ["proventos", "Proventos"],
  ["premissas", "Premissas"], ["verificar", "Verificar"]];
const TIPOS_PROVENTO = ["DIVIDENDO", "JCP", "RENDIMENTO", "AMORTIZAÇÃO"];
const CHAVE_DEMO = "invest-demo-v1";
const DIA = 864e5;

let D = null;                 // dados.json como veio do servidor
let E = null;                 // estado editável: lancamentos, proventos, metas, cfg
let aba = "painel";
let filtro = "Todos";
const graficos = {};
const projParams = {};        // parâmetros da Projeção alterados na página

// ------------------------------------------------------------------ utilidades

const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const nf = (c) => new Intl.NumberFormat("pt-BR", { minimumFractionDigits: c, maximumFractionDigits: c });
const NF = new Proxy({}, { get: (o, c) => o[c] || (o[c] = nf(Number(c))) });   // NF[2].format(x): casas decimais sob demanda
const ok = v => v !== null && v !== undefined && Number.isFinite(v);
const brl = (v, c = 2) => ok(v) ? "R$ " + NF[c].format(v) : "—";
const pct = (v, c = 1) => ok(v) ? NF[c].format(v * 100) + "%" : "—";
const num = (v, c = 0) => ok(v) ? NF[c].format(v) : "—";
const qtd = v => ok(v) ? (Number.isInteger(v) ? NF[0].format(v) : NF[4].format(v)) : "—";
const sinal = v => !ok(v) || v === 0 ? "" : (v > 0 ? "pos" : "neg");
const seta = (v, f = pct) => ok(v) ? (v > 0 ? "▲ " : v < 0 ? "▼ " : "") + f(Math.abs(v)) : "—";
const dataBR = s => s ? s.slice(8, 10) + "/" + s.slice(5, 7) + "/" + s.slice(0, 4) : "—";
const isoParaBR = s => s ? s.split("-").reverse().join("/") : "";
const hoje = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
const iso = d => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
const soma = (a, f = x => x) => a.reduce((s, x) => s + (f(x) || 0), 0);
const cor = n => getComputedStyle(document.documentElement).getPropertyValue("--" + n).trim();
const lerNumero = v => {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  let s = String(v).replace("R$", "").trim();
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
};

function toast(msg, ms = 3500) {
  const t = $("#aviso-toast");
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toast.t);
  toast.t = setTimeout(() => (t.hidden = true), ms);
}

const local = {
  ler() { try { return JSON.parse(localStorage.getItem(CHAVE_DEMO) || "null"); } catch { return null; } },
  gravar(v) { try { localStorage.setItem(CHAVE_DEMO, JSON.stringify(v)); } catch { /* sem armazenamento: vale só nesta aba */ } },
  apagar() { try { localStorage.removeItem(CHAVE_DEMO); } catch { /* idem */ } },
};

// ------------------------------------------------------------------ carga

async function carregar() {
  const r = await fetch("dados.json?" + Date.now(), { cache: "no-store" });
  if (!r.ok) throw new Error("dados.json não encontrado (" + r.status + ")");
  D = await r.json();
  const base = { lancamentos: D.lancamentos, proventos: D.proventos, metas: D.metas, cfg: D.cfg };
  E = structuredClone(base);
  if (D.modo === "demo") {
    const salvo = local.ler();
    if (salvo && salvo.gerado_em === D.gerado_em) Object.assign(E, salvo.estado);
  }
}

function salvarDemo() {
  local.gravar({ gerado_em: D.gerado_em, estado: E });
}

async function api(corpo) {
  const r = await fetch("api.php", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(corpo) });
  let j;
  try { j = await r.json(); } catch { throw new Error("resposta inválida do servidor (" + r.status + ")"); }
  if (!r.ok && !j.erros) throw new Error(j.erro || "falha no servidor (" + r.status + ")");
  return j;
}

// ------------------------------------------------------------------ cálculos (mesmas fórmulas da planilha)

function metaDe(tk) {
  const m = E.metas.find(x => x.ticker === tk);
  return { meta: m ? m.meta_mensal : E.cfg.meta_mensal_por_ativo, prazo: m ? m.prazo_meses : (E.cfg.prazo_meta_meses || 60),
    personalizada: !!(m && m.personalizada) };
}

/** Aba Carteira + Meta: uma linha por ativo do cadastro. */
function carteira() {
  const h = hoje(), umAno = new Date(h - 365 * DIA);
  const aporte = E.cfg.aporte_mensal || 0;
  const linhas = D.ativos.map(a => {
    const L = E.lancamentos.filter(x => x.ticker === a.ticker);
    const P = E.proventos.filter(x => x.ticker === a.ticker);
    const compras = L.filter(x => x.operacao === "COMPRA");
    const qCompra = soma(compras, x => x.quantidade);
    const q = qCompra - soma(L.filter(x => x.operacao === "VENDA"), x => x.quantidade);
    const pm = qCompra > 0 ? soma(compras, x => x.quantidade * x.preco + (x.custos || 0)) / qCompra : 0;
    const investido = q * pm, cot = a.preco || 0, pos = q * cot, lucro = pos - investido;
    const provRec = soma(P, x => x.valor_cota * x.quantidade);
    const prov12 = soma(P.filter(x => x.data && new Date(x.data + "T00:00") >= umAno), x => x.valor_cota * x.quantidade);
    const p12cota = a.provento_anual_cota || 0, pmes = p12cota / 12;
    const { meta, prazo, personalizada } = metaDe(a.ticker);
    const renda = q * pmes;
    const cotasMeta = pmes > 0 && meta > 0 ? Math.ceil(meta / pmes - 1e-9) : null;
    const falt = cotasMeta === null ? null : Math.max(0, cotasMeta - q);
    const valorFalt = falt === null ? null : falt * cot;
    const rendMes = cot > 0 ? pmes / cot : 0;
    const aportePrazo = valorFalt === null || !(prazo > 0) ? null : valorFalt / prazo;
    const aporteReinv = aportePrazo === null ? null : (rendMes > 0 ? valorFalt * rendMes / ((1 + rendMes) ** prazo - 1) : aportePrazo);
    const prev = new Date(h); prev.setMonth(prev.getMonth() + (prazo || 0));
    return {
      ticker: a.ticker, tipo: a.tipo, segmento: a.segmento, qtd: q, pm, investido, cot, pos, lucro,
      valorizacao: investido > 0 ? lucro / investido : 0, provRec, prov12, retorno: lucro + provRec,
      retornoPct: investido > 0 ? (lucro + provRec) / investido : 0, yoc: pm > 0 ? p12cota / pm : 0, dy: cot > 0 ? p12cota / cot : 0,
      pesoIdeal: a.peso_ideal || 0, p12cota, pmes, renda, meta, prazo, personalizada, rendMes,
      pctMeta: meta > 0 ? Math.min(1, renda / meta) : 0, cotasMeta, falt, valorFalt,
      valorTotal: cotasMeta === null ? null : cotasMeta * cot,
      aporteMeta: (falt || 0) * cot, mesesMeta: aporte > 0 ? Math.ceil(((falt || 0) * cot) / aporte) : 0,
      cotasMes: falt === null || !(prazo > 0) ? null : Math.ceil(falt / prazo), aportePrazo, aporteReinv, previsao: prev,
    };
  });
  const total = soma(linhas, l => l.pos);
  for (const l of linhas) {
    l.peso = total > 0 ? l.pos / total : 0;
    l.desvio = l.peso - l.pesoIdeal;
    l.aportePeso = Math.max(0, l.pesoIdeal * total - l.pos);
  }
  return linhas;
}

function totais(linhas) {
  const t = { pos: soma(linhas, l => l.pos), investido: soma(linhas, l => l.investido), lucro: soma(linhas, l => l.lucro),
    provRec: soma(linhas, l => l.provRec), prov12: soma(linhas, l => l.prov12), renda: soma(linhas, l => l.renda),
    meta: soma(linhas, l => l.meta), aporteMeta: soma(linhas, l => l.aporteMeta), falt: soma(linhas, l => l.falt),
    aportePrazo: soma(linhas, l => l.aportePrazo), aporteReinv: soma(linhas, l => l.aporteReinv),
    valorTotal: soma(linhas, l => l.valorTotal), valorFalt: soma(linhas, l => l.valorFalt), qtd: soma(linhas, l => l.qtd),
    cotasMes: soma(linhas, l => l.cotasMes) };
  t.valorizacao = t.investido > 0 ? t.lucro / t.investido : 0;
  t.retorno = t.lucro + t.provRec;
  t.retornoPct = t.investido > 0 ? t.retorno / t.investido : 0;
  t.naMeta = linhas.filter(l => l.pctMeta >= 1).length;
  t.pctMetaMedia = linhas.length ? soma(linhas, l => l.pctMeta) / linhas.length : 0;
  return t;
}

function alocacao(linhas) {
  const pesoTipo = tipo => (E.cfg.peso_ideal_renda_variavel || 0) * soma(D.ativos.filter(a => a.tipo === tipo), a => a.peso_ideal)
    / (soma(D.ativos, a => a.peso_ideal) || 1);
  const cls = [
    { classe: "Ações", valor: soma(linhas.filter(l => l.tipo === "AÇÃO"), l => l.pos), ideal: pesoTipo("AÇÃO") },
    { classe: "FIIs", valor: soma(linhas.filter(l => l.tipo === "FII"), l => l.pos), ideal: pesoTipo("FII") },
    ...E.cfg.outras_classes.map(o => ({ classe: o.classe, valor: o.valor, ideal: o.peso_ideal })),
  ];
  const tot = soma(cls, c => c.valor);
  for (const c of cls) {
    c.peso = tot > 0 ? c.valor / tot : 0;
    c.desvio = c.peso - c.ideal;
    c.aporte = Math.max(0, c.ideal * tot - c.valor);
  }
  return { classes: cls, total: tot };
}

// ------------------------------------------------------------------ componentes

function kpi(rot, val, det = "", clsVal = "", clsDet = "") {
  return `<div class="card kpi"><div class="rot">${esc(rot)}</div><div class="val ${clsVal}">${val}</div><div class="det ${clsDet}">${det}</div></div>`;
}

function secao(t, dica = "") {
  return `<div class="secao"><h2>${esc(t)}</h2>${dica ? `<span>${esc(dica)}</span>` : ""}</div>`;
}

function grafico(id, alto = false) {
  return `<div class="grafico${alto ? " alto" : ""}"><canvas id="${id}" aria-label="gráfico"></canvas></div>`;
}

/**
 * Tabela com ordenação por clique. colunas: [{t: título, v: linha => valor, f: formata, txt, cls: (v, linha) => classe, html}]
 */
function tabela(id, colunas, linhas, opts = {}) {
  const ord = tabela.ordem[id];
  let ls = [...linhas];
  if (ord) {
    const c = colunas[ord.i];
    ls.sort((a, b) => {
      const x = c.v(a), y = c.v(b);
      const r = typeof x === "string" || typeof y === "string" ? String(x ?? "").localeCompare(String(y ?? ""), "pt-BR")
        : (ok(x) ? x : -Infinity) - (ok(y) ? y : -Infinity);
      return ord.asc ? r : -r;
    });
  }
  const cab = colunas.map((c, i) => `<th class="${c.txt ? "txt" : ""}" data-tab="${id}" data-col="${i}" title="Ordenar">${esc(c.t)}${ord && ord.i === i ? (ord.asc ? " ▲" : " ▼") : ""}</th>`).join("");
  const corpo = ls.map(l => `<tr class="${opts.classeLinha ? opts.classeLinha(l) : ""}">` + colunas.map(c => {
    const v = c.v(l);
    const conteudo = c.html ? c.html(l) : esc(c.f ? c.f(v) : (v ?? "—"));
    return `<td class="${c.txt ? "txt" : ""} ${c.cls ? c.cls(v, l) : ""}">${conteudo}</td>`;
  }).join("") + "</tr>").join("");
  const total = opts.total ? `<tr class="total">${opts.total.map((x, i) => `<td class="${colunas[i]?.txt ? "txt" : ""}">${x ?? ""}</td>`).join("")}</tr>` : "";
  return `<div class="tabela-wrap"><table><thead><tr>${cab}</tr></thead><tbody>${corpo || `<tr><td class="txt" colspan="${colunas.length}">Nada registrado ainda.</td></tr>`}${total}</tbody></table></div>`;
}
tabela.ordem = {};

function desenhar(id, config) {
  if (graficos[id]) graficos[id].destroy();
  const el = document.getElementById(id);
  if (!el) return;
  if (typeof Chart === "undefined") {
    el.parentElement.innerHTML = '<div class="sem-grafico">Gráfico indisponível (Chart.js não carregou).</div>';
    return;
  }
  Chart.defaults.font.family = '"Segoe UI", system-ui, sans-serif';
  Chart.defaults.color = cor("suave");
  Chart.defaults.borderColor = cor("borda");
  config.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false }, config.options || {});
  graficos[id] = new Chart(el, config);
}

const tickBRL = { callback: v => "R$ " + NF[0].format(v) };
const tickPct = { callback: v => NF[0].format(v * 100) + "%" };

function barras(id, rotulos, valores, { horizontal = false, fmt = brl, corBarra = cor("azul"), pctEixo = false } = {}) {
  desenhar(id, {
    type: "bar",
    data: { labels: rotulos, datasets: [{ data: valores, backgroundColor: corBarra, borderRadius: 3, maxBarThickness: 28 }] },
    options: {
      indexAxis: horizontal ? "y" : "x", plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => fmt(c.raw) } } },
      scales: { [horizontal ? "x" : "y"]: { ticks: pctEixo ? tickPct : tickBRL, beginAtZero: true, max: pctEixo ? 1 : undefined },
        [horizontal ? "y" : "x"]: { grid: { display: false } } },
    },
  });
}

// ------------------------------------------------------------------ abas

const RENDER = {};

RENDER.painel = () => {
  const todas = carteira();
  const ls = todas.filter(l => filtro === "Todos" || l.tipo === filtro);
  const comPos = ls.filter(l => l.qtd > 0);
  const t = totais(ls), tt = totais(todas);
  const aloc = alocacao(todas);
  const maior = [...comPos].sort((a, b) => b.pos - a.pos)[0];
  const melhor = [...comPos].sort((a, b) => b.valorizacao - a.valorizacao)[0];
  const metaT = tt.meta;
  const vazio = !todas.some(l => l.qtd > 0);
  const fil = ["Todos", "AÇÃO", "FII"].map(f => `<button class="botao secundario ${filtro === f ? "ativo" : ""}" data-filtro="${f}">${f === "Todos" ? "Todos" : f === "AÇÃO" ? "Ações" : "FIIs"}</button>`).join("");
  return `
  ${vazio ? `<div class="msg erro">✚ Nenhuma ação ou FII na carteira ainda — registre sua primeira compra em <button class="link" data-ir="lancar">✚ Lançar</button>.</div>` : ""}
  <div class="lateral">
    <aside class="card filtros">
      <h3>Filtros</h3>${fil}
      <h3 style="margin-top:18px">Atalhos</h3>
      <button class="botao secundario" data-ir="carteira">Carteira</button>
      <button class="botao secundario" data-ir="meta">Meta por ativo</button>
      <button class="botao secundario" data-ir="projecao">Projeção</button>
      <button class="botao secundario" data-ir="verificar">Verificar dados</button>
      <h3 style="margin-top:18px">Progresso da meta</h3>
      <div class="barra"><i style="width:${(metaT > 0 ? Math.min(1, tt.renda / metaT) : 0) * 100}%"></i></div>
      <div class="det" style="font-size:12px;color:var(--suave);margin-top:6px">${brl(tt.renda)} de ${brl(metaT, 0)} por mês</div>
    </aside>
    <div>
      <div class="grade g6">
        ${kpi("PATRIMÔNIO TOTAL", brl(aloc.total, 0), pct(aloc.total > 0 ? tt.pos / aloc.total : 0, 0) + " em renda variável")}
        ${kpi("RENDA VARIÁVEL", brl(t.pos, 0), num(t.qtd) + " cotas no filtro")}
        ${kpi("LUCRO / PREJUÍZO", brl(t.lucro, 0), seta(t.valorizacao) + " sobre o investido", sinal(t.lucro), sinal(t.valorizacao))}
        ${kpi("PROVENTOS RECEBIDOS", brl(t.provRec, 0), "Renda mensal " + brl(t.renda))}
        ${kpi("MAIOR POSIÇÃO", maior ? brl(maior.pos, 0) : "—", maior ? maior.ticker : "")}
        ${kpi("MELHOR ATIVO", melhor ? melhor.ticker : "—", melhor ? seta(melhor.valorizacao) + " de valorização" : "", "", melhor ? sinal(melhor.valorizacao) : "")}
      </div>
      ${secao("Carteira e proventos")}
      <div class="grade g3">
        <div class="card"><h3>Proventos por mês</h3>${grafico("g-prov-mes")}</div>
        <div class="card"><h3>Posição por tipo</h3>${grafico("g-tipo")}</div>
        <div class="card"><h3>Posição por ativo</h3>${grafico("g-pos")}</div>
      </div>
      ${secao("Alocação e meta de renda")}
      <div class="grade g3">
        <div class="card"><h3>Patrimônio por classe</h3>${grafico("g-classe")}</div>
        <div class="card"><h3>Progresso da meta por ativo</h3>${grafico("g-meta")}</div>
        <div class="card"><h3>Proventos por ativo</h3>${grafico("g-prov-ativo")}</div>
      </div>
      ${secao("Meta de renda", "meta e prazo de cada ativo: aba Meta (padrão nas Premissas)")}
      <div class="grade g6">
        ${kpi("RENDA MENSAL", brl(tt.renda, 0), brl(tt.renda * 12, 0) + " por ano")}
        ${kpi("META MENSAL", brl(metaT, 0), "Aporte no prazo " + brl(tt.aportePrazo, 0) + "/mês")}
        ${kpi("PROGRESSO DA META", pct(tt.pctMetaMedia), pct(metaT > 0 ? tt.renda / metaT : 0) + " da renda total")}
        ${kpi("ATIVOS NA META", `${tt.naMeta} de ${todas.length}`, num(tt.falt) + " cotas faltantes")}
        ${kpi("APORTE P/ META", brl(tt.aporteMeta, 0), "Aporte de " + brl(E.cfg.aporte_mensal, 0) + "/mês")}
        ${kpi("TEMPO ATÉ A META", E.cfg.aporte_mensal > 0 ? num(Math.ceil(tt.aporteMeta / E.cfg.aporte_mensal)) + " meses" : "—",
          E.cfg.aporte_mensal > 0 ? NF[1].format(tt.aporteMeta / E.cfg.aporte_mensal / 12) + " anos" : "")}
      </div>
      ${secao("Destaques", "rankings no filtro escolhido")}
      <div class="grade g3">
        ${destaque("Maiores posições", ["Posição", "Participação"], [...comPos].sort((a, b) => b.pos - a.pos), l => [brl(l.pos, 0), pct(t.pos > 0 ? l.pos / t.pos : 0)])}
        ${destaque("Maiores altas", ["Valorização", "Lucro"], [...comPos].sort((a, b) => b.valorizacao - a.valorizacao), l => [seta(l.valorizacao), brl(l.lucro, 0)], true)}
        ${destaque("Maiores quedas", ["Valorização", "Lucro"], [...comPos].sort((a, b) => a.valorizacao - b.valorizacao), l => [seta(l.valorizacao), brl(l.lucro, 0)], true)}
      </div>
      ${secao("Alocação por classe", "valores de outras classes e pesos ideais: aba Premissas")}
      ${tabela("aloc", [
        { t: "Classe", v: c => c.classe, txt: true }, { t: "Valor", v: c => c.valor, f: v => brl(v, 0) },
        { t: "Peso atual", v: c => c.peso, f: pct }, { t: "Peso ideal", v: c => c.ideal, f: pct },
        { t: "Desvio", v: c => c.desvio, f: v => seta(v), cls: sinal }, { t: "Aporte p/ ideal", v: c => c.aporte, f: v => brl(v, 0) },
      ], aloc.classes, { total: ["TOTAL", brl(aloc.total, 0), pct(1), pct(soma(aloc.classes, c => c.ideal)), "", brl(soma(aloc.classes, c => c.aporte), 0)] })}
    </div>
  </div>`;
};

function destaque(titulo, cab, ls, vals, colorir = false) {
  const linhas = ls.slice(0, 5).map(l => {
    const [a, b] = vals(l);
    const c = colorir ? sinal(l.valorizacao) : "";
    return `<tr><td class="txt">${esc(l.ticker)}</td><td class="${c}">${a}</td><td class="${colorir ? sinal(l.lucro) : ""}">${b}</td></tr>`;
  }).join("") || `<tr><td class="txt" colspan="3">Sem posições no filtro.</td></tr>`;
  return `<div class="card"><h3>${esc(titulo)}</h3><table><thead><tr><th class="txt">Ativo</th><th>${cab[0]}</th><th>${cab[1]}</th></tr></thead><tbody>${linhas}</tbody></table></div>`;
}

RENDER.painel.depois = () => {
  const todas = carteira();
  const ls = todas.filter(l => filtro === "Todos" || l.tipo === filtro);
  const tks = new Set(ls.map(l => l.ticker));
  // proventos dos últimos 12 meses, por mês
  const h = hoje(), meses = [];
  for (let k = 11; k >= 0; k--) meses.push(new Date(h.getFullYear(), h.getMonth() - k, 1));
  const porMes = meses.map((m, i) => {
    const fim = i < 11 ? meses[i + 1] : new Date(h.getFullYear(), h.getMonth() + 1, 1);
    return soma(E.proventos.filter(p => tks.has(p.ticker) && p.data && new Date(p.data + "T00:00") >= m && new Date(p.data + "T00:00") < fim),
      p => p.valor_cota * p.quantidade);
  });
  barras("g-prov-mes", meses.map(m => m.toLocaleDateString("pt-BR", { month: "short", year: "2-digit" })), porMes);
  barras("g-tipo", ["Ações", "FIIs"], [soma(todas.filter(l => l.tipo === "AÇÃO"), l => l.pos), soma(todas.filter(l => l.tipo === "FII"), l => l.pos)],
    { horizontal: true, corBarra: cor("laranja") });
  const ordem = (k) => {
    const base = ls.some(l => l.qtd > 0) ? ls.filter(l => l.qtd > 0) : ls;
    return [...base].sort((a, b) => b[k] - a[k]);
  };
  const pos = ordem("pos");
  barras("g-pos", pos.map(l => l.ticker), pos.map(l => l.pos), { horizontal: true });
  const aloc = alocacao(todas).classes.filter(c => c.valor > 0);
  desenhar("g-classe", {
    type: "doughnut",
    data: { labels: aloc.map(c => c.classe), datasets: [{ data: aloc.map(c => c.valor), borderWidth: 2, borderColor: cor("card"),
      backgroundColor: [cor("azul"), cor("laranja"), cor("teal"), "#8b5cf6", "#94a3b8", "#ec4899"] }] },
    options: { cutout: "62%", plugins: { legend: { position: "right" }, tooltip: { callbacks: { label: c => c.label + ": " + brl(c.raw, 0) } } } },
  });
  const meta = [...ls].sort((a, b) => b.pctMeta - a.pctMeta);
  barras("g-meta", meta.map(l => l.ticker), meta.map(l => l.pctMeta), { horizontal: true, fmt: pct, corBarra: cor("laranja"), pctEixo: true });
  const prov = ordem("provRec");
  barras("g-prov-ativo", prov.map(l => l.ticker), prov.map(l => l.provRec), { horizontal: true });
};

RENDER.carteira = () => {
  const ls = carteira(), t = totais(ls);
  const C = [
    { t: "Ticker", v: l => l.ticker, txt: true }, { t: "Tipo", v: l => l.tipo, txt: true }, { t: "Qtd atual", v: l => l.qtd, f: qtd },
    { t: "Preço médio", v: l => l.pm, f: brl }, { t: "Valor investido", v: l => l.investido, f: brl }, { t: "Cotação", v: l => l.cot, f: brl },
    { t: "Posição atual", v: l => l.pos, f: brl }, { t: "Lucro / prejuízo", v: l => l.lucro, f: brl, cls: sinal },
    { t: "Valorização", v: l => l.valorizacao, f: pct, cls: sinal }, { t: "Proventos recebidos", v: l => l.provRec, f: brl },
    { t: "Proventos 12m", v: l => l.prov12, f: brl }, { t: "Retorno total", v: l => l.retorno, f: brl, cls: sinal },
    { t: "Retorno total %", v: l => l.retornoPct, f: pct, cls: sinal }, { t: "Yield on cost", v: l => l.yoc, f: pct },
    { t: "DY atual", v: l => l.dy, f: pct }, { t: "Peso atual", v: l => l.peso, f: pct }, { t: "Peso ideal", v: l => l.pesoIdeal, f: pct },
    { t: "Desvio do peso", v: l => l.desvio, f: v => seta(v), cls: sinal }, { t: "Aporte p/ peso ideal", v: l => l.aportePeso, f: brl },
    { t: "Provento mensal/cota", v: l => l.pmes, f: v => brl(v, 4) }, { t: "Renda mensal", v: l => l.renda, f: brl },
    { t: "% da meta", v: l => l.pctMeta, html: l => pct(l.pctMeta) + `<span class="mini"><i style="width:${l.pctMeta * 100}%"></i></span>` },
    { t: "Cotas p/ meta", v: l => l.cotasMeta, f: qtd }, { t: "Cotas faltantes", v: l => l.falt, f: qtd },
    { t: "Aporte p/ meta", v: l => l.aporteMeta, f: v => brl(v, 0) }, { t: "Meses p/ meta", v: l => l.mesesMeta, f: num },
    { t: "Segmento", v: l => l.segmento, txt: true },
  ];
  const tot = ["TOTAL", "", "", "", brl(t.investido), "", brl(t.pos), brl(t.lucro), pct(t.valorizacao), brl(t.provRec), brl(t.prov12),
    brl(t.retorno), pct(t.retornoPct), "", "", pct(1), pct(soma(ls, l => l.pesoIdeal)), "", brl(soma(ls, l => l.aportePeso)), "",
    brl(t.renda), pct(t.pctMetaMedia), `${t.naMeta} de ${ls.length} na meta`, num(t.falt), brl(t.aporteMeta, 0),
    E.cfg.aporte_mensal > 0 ? num(Math.ceil(t.aporteMeta / E.cfg.aporte_mensal)) : "", ""];
  return `<h2>Carteira</h2><p class="sub">Cada ação/FII em uma linha só: posição, proventos, peso e meta de renda. Linhas em cinza: ativos que você ainda não tem. Clique no cabeçalho para ordenar.</p>
    ${tabela("carteira", C, ls, { total: tot, classeLinha: l => l.qtd > 0 ? "" : "cinza" })}
    <ul class="notas"><li>Tipo, segmento, cotação, provento 12m e peso ideal vêm de dados/ativos.csv (atualizados pelo Yahoo Finance).</li>
    <li>Meta de renda: cotas p/ meta = meta mensal ÷ provento mensal por cota; aporte = cotas faltantes × cotação.</li></ul>`;
};

RENDER.meta = () => {
  const ls = carteira(), t = totais(ls);
  const C = [
    { t: "Ticker", v: l => l.ticker, txt: true }, { t: "Tipo", v: l => l.tipo, txt: true }, { t: "Segmento", v: l => l.segmento, txt: true },
    { t: "Valor da cota", v: l => l.cot, f: brl }, { t: "Provento mensal/cota", v: l => l.pmes, f: v => brl(v, 4) },
    { t: "Rendimento mensal", v: l => l.rendMes, f: v => pct(v, 2) },
    { t: "Meta mensal (R$)", v: l => l.meta, html: l => `<input inputmode="decimal" class="${l.personalizada ? "personalizada" : ""}" data-meta="${esc(l.ticker)}" data-campo="meta" value="${NF[2].format(l.meta)}" aria-label="Meta mensal ${esc(l.ticker)}">` },
    { t: "Prazo (meses)", v: l => l.prazo, html: l => `<input inputmode="numeric" style="width:70px" class="${l.personalizada ? "personalizada" : ""}" data-meta="${esc(l.ticker)}" data-campo="prazo" value="${l.prazo}" aria-label="Prazo ${esc(l.ticker)}">` },
    { t: "Cotas necessárias", v: l => l.cotasMeta, f: v => v === null ? "sem proventos" : qtd(v) },
    { t: "Valor total a investir", v: l => l.valorTotal, f: v => brl(v, 0) }, { t: "Cotas que você tem", v: l => l.qtd, f: qtd },
    { t: "Preço médio", v: l => l.pm, f: brl }, { t: "Investido hoje", v: l => l.investido, f: v => brl(v, 0) },
    { t: "Renda mensal atual", v: l => l.renda, f: brl },
    { t: "% da meta", v: l => l.pctMeta, html: l => pct(l.pctMeta) + `<span class="mini"><i style="width:${l.pctMeta * 100}%"></i></span>` },
    { t: "Cotas faltantes", v: l => l.falt, f: qtd }, { t: "Valor faltante", v: l => l.valorFalt, f: v => brl(v, 0) },
    { t: "Cotas por mês", v: l => l.cotasMes, f: qtd }, { t: "Aporte mensal no prazo", v: l => l.aportePrazo, f: v => brl(v, 0) },
    { t: "Aporte reinvestindo proventos", v: l => l.aporteReinv, f: v => brl(v, 0) },
    { t: "Data prevista", v: l => l.previsao.getTime(), f: v => new Date(v).toLocaleDateString("pt-BR", { month: "2-digit", year: "numeric" }) },
  ];
  const ap = E.cfg.aporte_mensal || 0, sobra = ap - t.aporteReinv;
  const sub = tipo => ls.filter(l => l.tipo === tipo);
  return `<h2>Meta de renda por ativo</h2>
    <p class="sub">Amarelo = você decide quanto cada ativo deve pagar por mês e em quantos meses (negrito azul = personalizado; apague para voltar ao padrão de ${brl(E.cfg.meta_mensal_por_ativo, 0)} em ${E.cfg.prazo_meta_meses || 60} meses).
    ${D.modo === "local" ? "As alterações são salvas em dados/metas.csv." : "Na demonstração, ficam só neste navegador."}</p>
    ${tabela("meta", C, ls, { total: ["TOTAL", "", "", "", "", "", brl(t.meta), "", "", brl(t.valorTotal, 0), "", "", brl(t.investido, 0), brl(t.renda),
      pct(t.meta > 0 ? t.renda / t.meta : 0), num(t.falt), brl(t.valorFalt, 0), num(t.cotasMes), brl(t.aportePrazo, 0), brl(t.aporteReinv, 0), ""] })}
    <div class="grade g3" style="margin-top:14px">
      <div class="card param">
        <span>Aporte mensal para cumprir todos os prazos</span><strong style="text-align:right">${brl(t.aportePrazo, 0)}</strong>
        <span>… reinvestindo os proventos</span><strong style="text-align:right">${brl(t.aporteReinv, 0)}</strong>
        <span>Aporte mensal disponível (Premissas)</span><strong style="text-align:right">${brl(ap, 0)}</strong>
        <span>Sobra (+) ou falta (−) por mês, reinvestindo</span><strong style="text-align:right" class="${sinal(sobra)}">${brl(sobra, 0)}</strong>
      </div>
      <div class="card param">
        <span>Ações — renda desejada</span><strong style="text-align:right">${brl(soma(sub("AÇÃO"), l => l.meta), 0)}</strong>
        <span>Ações — a investir</span><strong style="text-align:right">${brl(soma(sub("AÇÃO"), l => l.valorTotal), 0)}</strong>
        <span>FIIs — renda desejada</span><strong style="text-align:right">${brl(soma(sub("FII"), l => l.meta), 0)}</strong>
        <span>FIIs — a investir</span><strong style="text-align:right">${brl(soma(sub("FII"), l => l.valorTotal), 0)}</strong>
      </div>
    </div>
    <ul class="notas">
      <li>Cotas necessárias = meta mensal ÷ provento mensal por cota (proventos dos últimos 12 meses ÷ 12), arredondado para cima.</li>
      <li>Aporte reinvestindo proventos: cada cota comprada passa a pagar e o provento compra mais cotas (juros compostos à taxa do rendimento mensal).</li>
      <li>Cálculo com a cotação e o provento de hoje; mudanças de preço e de proventos alteram o resultado.</li>
    </ul>`;
};

function aoAlterarMeta(input) {
  const tk = input.dataset.meta, campo = input.dataset.campo;
  const padrao = { meta: E.cfg.meta_mensal_por_ativo, prazo: E.cfg.prazo_meta_meses || 60 };
  const atual = metaDe(tk);
  const bruto = input.value.trim();
  const v = bruto === "" ? null : lerNumero(bruto);
  if (v !== null && (!Number.isFinite(v) || v < 0 || (campo === "prazo" && v < 1))) {
    toast(campo === "prazo" ? "Prazo deve ser de pelo menos 1 mês." : "Meta mensal deve ser um valor positivo.");
    input.value = campo === "meta" ? NF[2].format(atual.meta) : atual.prazo;
    return;
  }
  const meta = campo === "meta" ? v : (atual.personalizada && atual.meta !== padrao.meta ? atual.meta : null);
  const prazo = campo === "prazo" ? (v === null ? null : Math.round(v)) : (atual.personalizada && atual.prazo !== padrao.prazo ? atual.prazo : null);
  let m = E.metas.find(x => x.ticker === tk);
  if (!m) E.metas.push(m = { ticker: tk });
  m.meta_mensal = meta ?? padrao.meta;
  m.prazo_meses = prazo ?? padrao.prazo;
  m.personalizada = meta !== null || prazo !== null;
  m._meta = meta; m._prazo = prazo;
  if (D.modo === "demo") salvarDemo();
  else salvarMetasServidor();
  render();
}

const salvarMetasServidor = (() => {
  let t;
  return () => {
    clearTimeout(t);
    t = setTimeout(async () => {
      const padrao = { meta: E.cfg.meta_mensal_por_ativo, prazo: E.cfg.prazo_meta_meses || 60 };
      const metas = D.ativos.map(a => {
        const m = E.metas.find(x => x.ticker === a.ticker);
        if (!m || !m.personalizada) return [a.ticker, "", ""];
        const meta = "_meta" in m ? m._meta : (m.meta_mensal !== padrao.meta ? m.meta_mensal : null);
        const prazo = "_prazo" in m ? m._prazo : (m.prazo_meses !== padrao.prazo ? m.prazo_meses : null);
        return [a.ticker, meta ?? "", prazo ?? ""];
      });
      try {
        const r = await api({ acao: "registrar", metas });
        toast(r.ok ? "Metas salvas em dados/metas.csv" : "Erro: " + r.erros.join("; "));
      } catch (e) { toast("Não foi possível salvar: " + e.message, 6000); }
    }, 600);
  };
})();

RENDER.quant = () => {
  const A = D.analise;
  if (A.vazio) return `<h2>Quant</h2><div class="msg erro">Sem histórico de cotações: rode com internet para gerar a análise quantitativa.</div>`;
  const pa = A.por_ativo, c = A.carteira;
  const ls = D.ativos.filter(a => pa[a.ticker]).map(a => ({ ticker: a.ticker, tipo: a.tipo, ...pa[a.ticker] }));
  const C = [
    { t: "Ticker", v: l => l.ticker, txt: true }, { t: "Qualidade", v: l => l.qualidade, txt: true },
    { t: "Retorno 12m", v: l => l.ret12_total, f: pct, cls: sinal }, { t: "Retorno anual (3a)", v: l => l.ret_anual, f: pct, cls: sinal },
    { t: "Volatilidade", v: l => l.vol, f: pct }, { t: "Vol. EWMA", v: l => l.vol_ewma, f: pct },
    { t: "Sharpe", v: l => l.sharpe, f: v => num(v, 2), cls: sinal }, { t: "Sortino", v: l => l.sortino, f: v => num(v, 2), cls: sinal },
    { t: "Queda máxima", v: l => l.max_queda, f: pct, cls: () => "neg" }, { t: "VaR 95% (sem.)", v: l => l.var95, f: pct },
    { t: "CVaR 95%", v: l => l.cvar95, f: pct }, { t: "Beta", v: l => l.beta, f: v => num(v, 2) },
    { t: "Correl. Ibov", v: l => l.corr_ibov, f: v => num(v, 2) }, { t: "CAPM", v: l => l.capm, f: pct },
    { t: "Alfa de Jensen", v: l => l.alfa, f: pct, cls: sinal }, { t: "Assimetria", v: l => l.assimetria, f: v => num(v, 2) },
    { t: "Curtose", v: l => l.curtose, f: v => num(v, 2) }, { t: "Jarque-Bera p", v: l => l.jb_p, f: v => num(v, 3) },
    { t: "Autocorr.", v: l => l.autocorr, f: v => num(v, 2) }, { t: "Taxa inst. f′", v: l => l.velocidade, f: pct, cls: sinal },
    { t: "Curvatura f″", v: l => l.aceleracao, f: v => num(v, 2), cls: sinal }, { t: "Leitura", v: l => l.leitura, txt: true },
    { t: "Tendência 3a", v: l => l.tendencia, f: pct, cls: sinal }, { t: "R²", v: l => l.r2, f: v => num(v, 2) },
    { t: "vs tendência", v: l => l.vs_tendencia, f: pct, cls: sinal }, { t: "Faixa 12m (mín.)", v: l => l.proj12_min, f: brl },
    { t: "Centro 12m", v: l => l.proj12, f: brl }, { t: "Faixa 12m (máx.)", v: l => l.proj12_max, f: brl },
    { t: "Proventos 12m", v: l => l.prov12, f: brl }, { t: "Cresc. proventos", v: l => l.cresc_prov, f: pct, cls: sinal },
    { t: "Peso na análise", v: l => l.peso_analise, f: pct }, { t: "Contrib. ao risco", v: l => l.contrib_risco, f: pct },
  ];
  const opcoes = ls.map(l => `<option ${l.ticker === RENDER.quant.sel ? "selected" : ""}>${esc(l.ticker)}</option>`).join("");
  return `<h2>Análise quantitativa</h2>
    <p class="sub">Histórico semanal de 5 anos (Yahoo Finance) até ${dataBR(A.data_base)}. Pesos: ${esc(c.base)}. Janela de risco: 3 anos; "12m" = 52 semanas.</p>
    <div class="grade g6">
      ${kpi("VOLATILIDADE", pct(c.vol), "diversificação " + num(c.diversificacao, 2) + "×")}
      ${kpi("RETORNO ANUAL", pct(c.ret_anual), "12m: " + pct(c.ret12) + (ok(c.ibov12) ? " · Ibov " + pct(c.ibov12) : ""), sinal(c.ret_anual))}
      ${kpi("SHARPE", num(c.sharpe, 2), "CDI " + pct(A.cfg.cdi), sinal(c.sharpe))}
      ${kpi("BETA", num(c.beta, 2), "sensibilidade ao Ibovespa")}
      ${kpi("DIVIDEND YIELD", pct(c.dy), "ponderado pela carteira")}
      ${kpi("QUEDA MÁXIMA", pct(c.max_queda), c.tir !== null && c.tir !== undefined ? "TIR dos lançamentos " + pct(c.tir) : "3 anos", "neg")}
    </div>
    ${secao("Métricas por ativo", "clique no cabeçalho para ordenar")}
    ${tabela("quant", C, ls)}
    <div class="grade g2" style="margin-top:14px">
      <div class="card"><h3>Preço semanal e faixa de 12 meses <select id="sel-quant" class="busca" style="width:auto;margin:0 0 0 8px">${opcoes}</select></h3>${grafico("g-preco", true)}</div>
      <div class="card"><h3>Matriz de correlação (retornos semanais, 3 anos)</h3><div style="overflow:auto">${correlacao(A.correlacao)}</div></div>
    </div>
    <ul class="notas">
      <li>Derivadas analíticas de um polinômio de grau 3 ajustado ao log-preço de 12 meses: f′ = taxa instantânea anual, f″ = curvatura.</li>
      <li>Faixa 12m: regressão log-linear de 3 anos ± 1 desvio dos resíduos. VaR/CVaR: perda semanal no pior 5%.</li>
    </ul>`;
};

function correlacao(m) {
  const cel = v => {
    if (!ok(v)) return `<td>—</td>`;
    const a = Math.abs(v), base = v >= 0 ? "37,99,235" : "220,38,38";
    return `<td style="background:rgba(${base},${(a * .75).toFixed(2)});color:${a > .55 ? "#fff" : "inherit"}">${NF[2].format(v)}</td>`;
  };
  return `<table><thead><tr><th class="txt"></th>${m.tickers.map(t => `<th>${esc(t)}</th>`).join("")}</tr></thead><tbody>` +
    m.matriz.map((r, i) => `<tr><td class="txt">${esc(m.tickers[i])}</td>${r.map(cel).join("")}</tr>`).join("") + "</tbody></table>";
}

RENDER.quant.depois = () => {
  if (D.analise.vazio) return;
  const sel = $("#sel-quant");
  if (!sel) return;
  RENDER.quant.sel = sel.value;
  const serie = (D.historico[sel.value] || []).slice(-156);
  const m = D.analise.por_ativo[sel.value] || {};
  const rot = serie.map(p => dataBR(p[0]));
  const ds = [{ label: "Fechamento", data: serie.map(p => p[1]), borderColor: cor("azul"), pointRadius: 0, borderWidth: 1.6, tension: .1 }];
  if (ok(m.proj12)) {
    const prox = new Date(serie.length ? serie[serie.length - 1][0] + "T00:00" : hoje());
    prox.setFullYear(prox.getFullYear() + 1);
    rot.push(prox.toLocaleDateString("pt-BR"));
    const extra = (v) => [...serie.map((_, i) => i === serie.length - 1 ? serie[i][1] : null), v];
    ds[0].data.push(null);
    ds.push({ label: "Centro 12m", data: extra(m.proj12), borderColor: cor("laranja"), borderDash: [5, 4], pointRadius: 0, spanGaps: true },
      { label: "Mín.", data: extra(m.proj12_min), borderColor: cor("cinza-serie"), borderDash: [2, 3], pointRadius: 0, spanGaps: true },
      { label: "Máx.", data: extra(m.proj12_max), borderColor: cor("cinza-serie"), borderDash: [2, 3], pointRadius: 0, spanGaps: true });
  }
  desenhar("g-preco", { type: "line", data: { labels: rot, datasets: ds },
    options: { interaction: { mode: "index", intersect: false }, plugins: { tooltip: { callbacks: { label: c => c.dataset.label + ": " + brl(c.raw) } } },
      scales: { x: { ticks: { maxTicksLimit: 8 } }, y: { ticks: { callback: v => "R$ " + NF[0].format(v) } } } } });
};

RENDER.modelos = () => {
  const A = D.analise, M = A.modelos;
  if (A.vazio || !M) return `<h2>Modelos</h2><div class="msg erro">Modelos precisam do histórico de pelo menos 2 ativos.</div>`;
  const e = M.economia, f = M.fatores, mc = M.monte_carlo;
  const nomes = Object.keys(M.carteiras);
  const tks = f.cargas.tickers;
  const linhasPesos = tks.map(t => ({ ticker: t, ...Object.fromEntries(nomes.map(n => [n, M.carteiras[n][t]])), mu: M.mu[t] }));
  const C = [{ t: "Ticker", v: l => l.ticker, txt: true }, { t: "Retorno esperado μ", v: l => l.mu, f: pct },
    ...nomes.map(n => ({ t: n, v: l => l[n], f: v => v > 0.0005 ? pct(v) : "—" }))];
  const fim = mc.anos[mc.anos.length - 1] || {};
  return `<h2>Modelos</h2><p class="sub">Economia (Fisher), álgebra linear (fatores de risco e Markowitz) e processo estocástico (Monte Carlo) sobre a carteira.</p>
    <div class="grade g6">
      ${kpi("INFLAÇÃO", pct(e.inflacao), "premissa anual")}
      ${kpi("CDI REAL", pct(e.cdi_real), "Fisher: (1+i)/(1+π)−1")}
      ${kpi("CARTEIRA REAL", pct(e.carteira_real), "retorno anual descontada a inflação", sinal(e.carteira_real))}
      ${kpi("DY REAL", pct(e.dy_real), "dividend yield real", sinal(e.dy_real))}
      ${kpi("PRÊMIO DE RISCO", pct(e.premio_risco), "retorno − CDI", sinal(e.premio_risco))}
      ${kpi("CONDIÇÃO Σ", num(f.condicao, 0), "λmáx / λmín da covariância")}
    </div>
    ${secao("Carteiras de Markowitz", "sem venda a descoberto; μ = metade histórico, metade CAPM")}
    <div class="grade g2">
      ${tabela("mk", C, linhasPesos)}
      <div class="card"><h3>Risco × retorno</h3>${grafico("g-fronteira", true)}</div>
    </div>
    <div style="margin-top:14px">${tabela("mk-met", [{ t: "Carteira", v: l => l.n, txt: true }, { t: "Retorno", v: l => l.ret, f: pct, cls: sinal },
      { t: "Volatilidade", v: l => l.vol, f: pct }, { t: "Sharpe", v: l => l.sharpe, f: v => num(v, 2), cls: sinal }],
      nomes.map(n => ({ n, ...M.metricas[n] })))}</div>
    ${secao("Fatores de risco", "decomposição espectral Σ = VΛVᵀ")}
    <div class="grade g2">
      <div class="card"><h3>Variância explicada por fator</h3>${grafico("g-fatores")}</div>
      ${tabela("cargas", [{ t: "Ticker", v: l => l.t, txt: true }, ...f.cargas.colunas.map((c, j) => ({ t: c, v: l => l.v[j], f: v => num(v, 3), cls: sinal }))],
        tks.map((t, i) => ({ t, v: f.cargas.valores[i] })))}
    </div>
    ${secao("Monte Carlo", `${num(mc.n)} simulações · movimento browniano geométrico mensal · μ ${pct(mc.mu)} · σ ${pct(mc.sigma)}`)}
    <div class="grade g6">
      ${kpi("MEDIANA NO FIM", brl(fim.p50, 0), `ano ${fim.ano ?? "—"}`)}
      ${kpi("FAIXA 5%–95%", brl(fim.p5, 0), "até " + brl(fim.p95, 0))}
      ${kpi("VENCER O CDI", pct(fim.prob_cdi, 0), "probabilidade no fim do horizonte")}
      ${kpi("PERDA NOMINAL", pct(fim.prob_perda, 0), "terminar abaixo do aportado")}
      ${kpi("ATINGIR A META", pct(mc.prob_meta, 0), mc.ano_meta_mediano ? "mediana: " + NF[1].format(mc.ano_meta_mediano) + " anos" : "no horizonte")}
      ${kpi("TOTAL APORTADO", brl(fim.aportado, 0), "patrimônio inicial + aportes")}
    </div>
    <div class="card" style="margin-top:14px">${grafico("g-mc", true)}</div>
    <ul class="notas"><li>Monte Carlo calculado em Python com a carteira e o aporte da última geração${D.modo === "demo" ? " (não muda com as edições da demonstração)" : ""}.</li></ul>`;
};

RENDER.modelos.depois = () => {
  const M = D.analise.modelos;
  if (!M) return;
  const nomes = Object.keys(M.metricas);
  const pal = [cor("cinza-serie"), cor("azul"), cor("teal"), cor("laranja")];
  const tk = M.fatores.cargas.tickers;
  const pa = D.analise.por_ativo;
  desenhar("g-fronteira", { type: "scatter", data: { datasets: [
    { label: "Ativos", data: tk.map(t => ({ x: pa[t]?.vol, y: M.mu[t], t })), backgroundColor: cor("cinza-serie"), pointRadius: 4 },
    ...nomes.map((n, i) => ({ label: n, data: [{ x: M.metricas[n].vol, y: M.metricas[n].ret, t: n }], backgroundColor: pal[i % 4], pointRadius: 8, pointStyle: "rectRot" })),
  ] }, options: { plugins: { tooltip: { callbacks: { label: c => `${c.raw.t}: retorno ${pct(c.raw.y)} · vol ${pct(c.raw.x)}` } } },
    scales: { x: { title: { display: true, text: "Volatilidade" }, ticks: tickPct }, y: { title: { display: true, text: "Retorno esperado" }, ticks: tickPct } } } });
  const ex = M.fatores.explicado;
  barras("g-fatores", ex.map((_, i) => "Fator " + (i + 1)), ex, { fmt: pct, corBarra: cor("teal"), pctEixo: true });
  const A = M.monte_carlo.anos;
  const linha = (k, label, c, extra = {}) => ({ label, data: A.map(a => a[k]), borderColor: c, backgroundColor: c, pointRadius: 0, borderWidth: 1.5, ...extra });
  const faixa = "rgba(37,99,235,.12)";
  desenhar("g-mc", { type: "line", data: { labels: A.map(a => "Ano " + a.ano), datasets: [
    linha("p95", "95%", "transparent", { fill: "+4", backgroundColor: faixa }), linha("p75", "75%", "transparent", { fill: "+2", backgroundColor: faixa }),
    linha("p50", "Mediana", cor("azul"), { borderWidth: 2.5 }), linha("p25", "25%", "transparent"), linha("p5", "5%", "transparent"),
    linha("cdi", "CDI", cor("laranja"), { borderDash: [5, 4] }), linha("aportado", "Aportado", cor("suave"), { borderDash: [2, 3] }),
  ] }, options: { interaction: { mode: "index", intersect: false },
    plugins: { legend: { labels: { filter: i => !["95%", "75%", "25%", "5%"].includes(i.text) } }, tooltip: { callbacks: { label: c => c.dataset.label + ": " + brl(c.raw, 0) } } },
    scales: { y: { ticks: tickBRL } } } });
};

function parametrosProjecao() {
  const ls = carteira(), t = totais(ls);
  const p = (D.analise.projecao || { dy: 0.06, cdi: D.analise.cfg.cdi, reinvestir: D.analise.cfg.reinvestir, g_base: 0, g_pess: -0.05, g_otim: 0.05,
    horizonte: D.analise.cfg.horizonte });
  const base = { inicial: t.pos, aporte: E.cfg.aporte_mensal || 0, dy: p.dy, reinv: p.reinvestir ? 1 : 0, g_pess: p.g_pess, g_base: p.g_base,
    g_otim: p.g_otim, cdi: p.cdi, meta: t.meta, horizonte: p.horizonte };
  return Object.assign(base, projParams);
}

function projetar(P) {
  const meses = Math.round(P.horizonte * 12), out = [];
  let s = [P.inicial, P.inicial, P.inicial], cdi = P.inicial, aportado = P.inicial;
  const g = [P.g_pess, P.g_base, P.g_otim];
  for (let m = 0; m <= meses; m++) {
    if (m > 0) {
      s = s.map((v, i) => v * (1 + g[i]) ** (1 / 12) + P.aporte + P.reinv * v * P.dy / 12);
      cdi = cdi * (1 + P.cdi) ** (1 / 12) + P.aporte;
      aportado += P.aporte;
    }
    out.push({ m, pess: s[0], base: s[1], otim: s[2], cdi, renda: s[1] * P.dy / 12, aportado });
  }
  return out;
}

RENDER.projecao = () => {
  const P = parametrosProjecao(), M = projetar(P);
  const campos = [["inicial", "Patrimônio inicial (ações + FIIs)", "brl", "Carteira (posição atual)"], ["aporte", "Aporte mensal", "brl", "Premissas"],
    ["dy", "Dividend yield anual", "pct", "Proventos 12m ÷ cotação, ponderado"], ["reinv", "Reinvestir proventos (1 = sim)", "int", "config.json → analise"],
    ["g_pess", "Valorização anual — pessimista", "pct", "Base − 1 desvio-padrão no horizonte"], ["g_base", "Valorização anual — base", "pct", "Tendência 3a × R² (−10%…+20%)"],
    ["g_otim", "Valorização anual — otimista", "pct", "Base + 1 desvio-padrão no horizonte"], ["cdi", "CDI anual (comparação)", "pct", "config.json"],
    ["meta", "Meta de renda mensal total", "brl", "Soma das metas (aba Meta)"], ["horizonte", "Horizonte (anos)", "int", "config.json"]];
  const val = (k, t) => t === "pct" ? NF[2].format(P[k] * 100) : t === "brl" ? NF[2].format(P[k]) : String(P[k]);
  const anos = M.filter(x => x.m > 0 && x.m % 12 === 0);
  const ate = M.findIndex(x => x.renda >= P.meta);
  return `<h2>Projeção futura</h2><p class="sub">Evolução mês a mês em 3 cenários + o mesmo aporte no CDI. Altere os parâmetros em amarelo (percentuais em %).
    <button class="link" id="proj-reset">Restaurar parâmetros</button></p>
    <div class="grade g2">
      <div class="card param">${campos.map(([k, r, t, o]) => `<span>${esc(r)}</span><input data-proj="${k}" data-tipo="${t}" value="${val(k, t)}" inputmode="decimal" aria-label="${esc(r)}"><small>${esc(o)}</small>`).join("")}</div>
      <div>
        <div class="grade g6" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))">
          ${kpi("BASE NO FIM", brl(M[M.length - 1].base, 0), "pessimista " + brl(M[M.length - 1].pess, 0))}
          ${kpi("CDI NO FIM", brl(M[M.length - 1].cdi, 0), "base − CDI " + brl(M[M.length - 1].base - M[M.length - 1].cdi, 0), "", sinal(M[M.length - 1].base - M[M.length - 1].cdi))}
          ${kpi("RENDA MENSAL NO FIM", brl(M[M.length - 1].renda, 0), "cenário base")}
          ${kpi("MESES ATÉ A META", P.meta <= 0 ? "—" : ate < 0 ? `além de ${P.horizonte} anos` : num(ate), ate > 0 ? NF[1].format(ate / 12) + " anos" : "")}
        </div>
        <div class="card" style="margin-top:14px">${grafico("g-proj", true)}</div>
      </div>
    </div>
    ${secao("Resumo por ano")}
    ${tabela("proj", [{ t: "Ano", v: x => x.m / 12, f: num }, { t: "Pessimista", v: x => x.pess, f: v => brl(v, 0) }, { t: "Base", v: x => x.base, f: v => brl(v, 0) },
      { t: "Otimista", v: x => x.otim, f: v => brl(v, 0) }, { t: "CDI", v: x => x.cdi, f: v => brl(v, 0) }, { t: "Renda mensal base", v: x => x.renda, f: v => brl(v, 0) },
      { t: "Total aportado", v: x => x.aportado, f: v => brl(v, 0) }, { t: "Base − CDI", v: x => x.base - x.cdi, f: v => brl(v, 0), cls: sinal },
      { t: "% da meta de renda", v: x => P.meta > 0 ? x.renda / P.meta : 0, html: x => { const v = P.meta > 0 ? Math.min(1, x.renda / P.meta) : 0; return pct(P.meta > 0 ? x.renda / P.meta : 0) + `<span class="mini"><i style="width:${v * 100}%"></i></span>`; } }], anos)}`;
};

RENDER.projecao.depois = () => {
  const M = projetar(parametrosProjecao());
  const l = (k, label, c, extra = {}) => ({ label, data: M.map(x => x[k]), borderColor: c, backgroundColor: c, pointRadius: 0, borderWidth: 1.8, ...extra });
  desenhar("g-proj", { type: "line", data: { labels: M.map(x => x.m % 12 === 0 ? "Ano " + x.m / 12 : ""), datasets: [
    l("otim", "Otimista", cor("teal")), l("base", "Base", cor("azul"), { borderWidth: 2.6 }), l("pess", "Pessimista", cor("vermelho")),
    l("cdi", "CDI", cor("laranja"), { borderDash: [5, 4] }), l("aportado", "Aportado", cor("suave"), { borderDash: [2, 3] }),
  ] }, options: { interaction: { mode: "index", intersect: false }, plugins: { tooltip: { callbacks: { title: c => "Mês " + c[0].dataIndex, label: c => c.dataset.label + ": " + brl(c.raw, 0) } } },
    scales: { x: { ticks: { autoSkip: false, maxRotation: 0, callback: (v, i) => M[i].m % 12 === 0 && M[i].m > 0 ? "Ano " + M[i].m / 12 : null } }, y: { ticks: tickBRL } } } });
};

// ---- Lançar: mesmas regras de dominio/lancamentos.py
function validarLancamento(d, cotas) {
  const tk = (d.Ticker || "").trim().toUpperCase(), op = (d["Operação"] || "COMPRA").toUpperCase();
  const q = lerNumero(d["Quantidade de cotas"]), custos = lerNumero(d["Custos (R$)"]) || 0;
  if (!tk) throw new Error("informe o ticker");
  if (!["COMPRA", "VENDA"].includes(op)) throw new Error("operação deve ser COMPRA ou VENDA");
  if (!q || !(q > 0)) throw new Error("informe a quantidade de cotas");
  const ativo = D.ativos.find(a => a.ticker === tk);
  if (!ativo) {
    if (D.modo === "demo") throw new Error("na demonstração, use um ticker do cadastro");
    if (!["AÇÃO", "FII"].includes((d["Tipo do ativo"] || "").toUpperCase())) throw new Error("ticker novo: informe o Tipo do ativo (AÇÃO ou FII)");
  }
  const preco = lerNumero(d["Preço por cota (R$)"]) || (ativo ? ativo.preco : null);
  if (!preco || !(preco > 0)) throw new Error("sem cotação para este ticker — informe o preço por cota");
  if (op === "VENDA" && q > (cotas[tk] || 0)) throw new Error(`venda maior que as cotas que você tem (${qtd(cotas[tk] || 0)})`);
  return { tk, op, q, preco, custos };
}

function validarProvento(d, cotas) {
  const tk = (d.Ticker || "").trim().toUpperCase(), tipo = (d["Tipo de provento"] || "DIVIDENDO").toUpperCase();
  const valor = lerNumero(d["Valor por cota (R$)"]), q = lerNumero(d["Qtd de cotas"]) || cotas[tk] || 0;
  if (!tk || !D.ativos.some(a => a.ticker === tk)) throw new Error("ticker não cadastrado — registre a compra primeiro");
  if (!TIPOS_PROVENTO.includes(tipo)) throw new Error("tipo de provento inválido");
  if (!valor || !(valor > 0)) throw new Error("informe o valor por cota");
  if (!(q > 0)) throw new Error("você não tem cotas deste ativo — informe a quantidade");
  return { tk, tipo, valor, q };
}

const cotasAtuais = () => Object.fromEntries(carteira().map(l => [l.ticker, l.qtd]));

RENDER.lancar = () => {
  const opcoes = D.ativos.map(a => `<option value="${esc(a.ticker)}">${esc(a.tipo)} · ${esc(a.segmento)}</option>`).join("");
  const ult = [...E.lancamentos].sort((a, b) => (b.data || "").localeCompare(a.data || "")).slice(0, 10);
  return `<h2>✚ Lançar</h2><p class="sub">Escolha o ticker e informe a quantidade de cotas. Preço vazio = cotação atual; data vazia = hoje.
    ${D.modo === "local" ? "Grava em dados/lancamentos.csv e dados/proventos.csv com as mesmas regras do Cadastro da planilha." : "Na demonstração, os registros ficam só neste navegador."}</p>
    <datalist id="tickers">${opcoes}</datalist>
    <div class="grade g2">
      <form class="card" id="f-lanc" autocomplete="off">
        <h3>Compra / venda de cotas</h3>
        <div class="formulario">
          <label>Data <input type="date" name="Data"></label>
          <label>Ticker <input name="Ticker" list="tickers" required placeholder="ex.: ITSA4" style="text-transform:uppercase"></label>
          <label>Operação <select name="Operação"><option>COMPRA</option><option>VENDA</option></select></label>
          <label>Quantidade de cotas <input name="Quantidade de cotas" inputmode="decimal" required></label>
          <label>Preço por cota (R$) <input name="Preço por cota (R$)" inputmode="decimal" placeholder="vazio = cotação"></label>
          <label>Custos / taxas (R$) <input name="Custos (R$)" inputmode="decimal" placeholder="0,00"></label>
          <label>Ativo <span class="auto" data-auto="ativo">—</span></label>
          <label>Cotação atual <span class="auto" data-auto="cot">—</span></label>
          <label>Cotas que você tem <span class="auto" data-auto="tem">—</span></label>
          <label>Total do lançamento <span class="auto" data-auto="total">—</span></label>
          ${D.modo === "local" ? `<label>Tipo (só ticker novo) <select name="Tipo do ativo"><option value="">—</option><option>AÇÃO</option><option>FII</option></select></label>
          <label>Segmento (só ticker novo) <input name="Segmento"></label>` : ""}
          <div class="largo"><button class="botao" type="submit">✔ Registrar lançamento</button> <button class="botao secundario" type="reset">Limpar</button></div>
        </div>
        <div class="msg" hidden></div>
      </form>
      <form class="card" id="f-prov" autocomplete="off">
        <h3>Provento recebido</h3>
        <div class="formulario">
          <label>Data pagamento <input type="date" name="Data pagamento"></label>
          <label>Ticker <input name="Ticker" list="tickers" required style="text-transform:uppercase"></label>
          <label>Tipo de provento <select name="Tipo de provento">${TIPOS_PROVENTO.map(t => `<option>${t}</option>`).join("")}</select></label>
          <label>Valor por cota (R$) <input name="Valor por cota (R$)" inputmode="decimal" required></label>
          <label>Qtd de cotas <input name="Qtd de cotas" inputmode="decimal" placeholder="vazio = suas cotas"></label>
          <label>Total recebido <span class="auto" data-auto="total">—</span></label>
          <div class="largo"><button class="botao" type="submit">✔ Registrar provento</button> <button class="botao secundario" type="reset">Limpar</button></div>
        </div>
        <div class="msg" hidden></div>
      </form>
    </div>
    ${secao("Últimos lançamentos")}
    ${tabelaLancamentos("ult", ult)}`;
};

RENDER.lancar.depois = () => {
  const fl = $("#f-lanc"), fp = $("#f-prov");
  const atualizarAuto = () => {
    const cot = cotasAtuais();
    const d = Object.fromEntries(new FormData(fl));
    const tk = (d.Ticker || "").toUpperCase(), a = D.ativos.find(x => x.ticker === tk);
    const q = lerNumero(d["Quantidade de cotas"]), p = lerNumero(d["Preço por cota (R$)"]) || (a ? a.preco : null);
    $('[data-auto="ativo"]', fl).textContent = a ? `${a.tipo} · ${a.segmento}` : tk ? "ticker novo" : "—";
    $('[data-auto="cot"]', fl).textContent = a ? brl(a.preco) : "—";
    $('[data-auto="tem"]', fl).textContent = tk ? qtd(cot[tk] || 0) : "—";
    $('[data-auto="total"]', fl).textContent = q && p ? brl(q * p + (lerNumero(d["Custos (R$)"]) || 0)) : "—";
    const dp = Object.fromEntries(new FormData(fp));
    const tp = (dp.Ticker || "").toUpperCase();
    const qp = lerNumero(dp["Qtd de cotas"]) || cot[tp] || 0, vp = lerNumero(dp["Valor por cota (R$)"]);
    $('[data-auto="total"]', fp).textContent = vp && qp ? brl(vp * qp) : "—";
  };
  fl.addEventListener("input", atualizarAuto);
  fp.addEventListener("input", atualizarAuto);
  fl.addEventListener("reset", () => setTimeout(atualizarAuto));
  fp.addEventListener("reset", () => setTimeout(atualizarAuto));
  fl.addEventListener("submit", ev => registrar(ev, fl, "lancamentos"));
  fp.addEventListener("submit", ev => registrar(ev, fp, "proventos"));
};

async function registrar(ev, form, tipo) {
  ev.preventDefault();
  const msg = $(".msg", form), botao = $('button[type="submit"]', form);
  const d = Object.fromEntries(new FormData(form));
  for (const k of ["Data", "Data pagamento"]) if (k in d) d[k] = isoParaBR(d[k]);
  d.Ticker = (d.Ticker || "").trim().toUpperCase();
  const mostrar = (texto, okk) => { msg.hidden = false; msg.className = "msg " + (okk ? "ok" : "erro"); msg.textContent = texto; };
  try {
    const cot = cotasAtuais();
    const v = tipo === "lancamentos" ? validarLancamento(d, cot) : validarProvento(d, cot);
    if (D.modo === "demo") {
      const data = d.Data || d["Data pagamento"] ? (d.Data || d["Data pagamento"]).split("/").reverse().join("-") : iso(hoje());
      if (tipo === "lancamentos") E.lancamentos.push({ data, ticker: v.tk, operacao: v.op, quantidade: v.q, preco: v.preco, custos: v.custos });
      else E.proventos.push({ data, ticker: v.tk, tipo: v.tipo, valor_cota: v.valor, quantidade: v.q });
      salvarDemo();
      render();
      toast(`✔ ${tipo === "lancamentos" ? v.op.toLowerCase() : "provento"} de ${v.tk} registrado (só neste navegador)`);
      return;
    }
    botao.disabled = true;
    const r = await api({ acao: "registrar", [tipo]: [d] });
    if (!r.ok) return mostrar(r.erros.join(" · "), false);
    await carregar();
    render();
    toast("✔ " + r.resumo);
  } catch (e) {
    mostrar(e.message, false);
  } finally {
    botao.disabled = false;
  }
}

function tabelaLancamentos(id, ls) {
  return tabela(id, [{ t: "Data", v: l => l.data || "", f: dataBR, txt: true }, { t: "Ticker", v: l => l.ticker, txt: true },
    { t: "Operação", v: l => l.operacao, txt: true, cls: v => v === "VENDA" ? "neg" : "" }, { t: "Quantidade", v: l => l.quantidade, f: qtd },
    { t: "Preço", v: l => l.preco, f: brl }, { t: "Custos", v: l => l.custos, f: brl },
    { t: "Total", v: l => l.quantidade * l.preco + (l.custos || 0), f: brl }], ls);
}

function filtrar(ls, termo) {
  const t = (termo || "").trim().toUpperCase();
  return t ? ls.filter(l => Object.values(l).some(v => String(v ?? "").toUpperCase().includes(t))) : ls;
}

RENDER.lancamentos = () => {
  const ls = filtrar([...E.lancamentos].sort((a, b) => (b.data || "").localeCompare(a.data || "")), RENDER.lancamentos.busca);
  const compras = soma(ls.filter(l => l.operacao === "COMPRA"), l => l.quantidade * l.preco + (l.custos || 0));
  const vendas = soma(ls.filter(l => l.operacao === "VENDA"), l => l.quantidade * l.preco);
  return `<h2>Lançamentos de compra e venda</h2><p class="sub">${ls.length} registro(s) · compras ${brl(compras, 0)} · vendas ${brl(vendas, 0)}</p>
    <input class="busca" data-busca="lancamentos" placeholder="Filtrar (ticker, data, operação)…" value="${esc(RENDER.lancamentos.busca || "")}">
    ${tabelaLancamentos("lanc", ls)}`;
};

RENDER.proventos = () => {
  const ls = filtrar([...E.proventos].sort((a, b) => (b.data || "").localeCompare(a.data || "")), RENDER.proventos.busca);
  return `<h2>Proventos recebidos</h2><p class="sub">${ls.length} registro(s) · total ${brl(soma(ls, p => p.valor_cota * p.quantidade))}</p>
    <input class="busca" data-busca="proventos" placeholder="Filtrar (ticker, tipo, data)…" value="${esc(RENDER.proventos.busca || "")}">
    ${tabela("prov", [{ t: "Data pagamento", v: l => l.data || "", f: dataBR, txt: true }, { t: "Ticker", v: l => l.ticker, txt: true },
      { t: "Tipo", v: l => l.tipo, txt: true }, { t: "Valor por cota", v: l => l.valor_cota, f: v => brl(v, 4) },
      { t: "Qtd de cotas", v: l => l.quantidade, f: qtd }, { t: "Total", v: l => l.valor_cota * l.quantidade, f: brl }], ls)}`;
};

RENDER.premissas = () => {
  const c = E.cfg, demo = D.modo === "demo";
  const campo = (k, rot, v, uso) => `<span>${esc(rot)}</span>${demo ? `<input data-cfg="${k}" value="${v}" inputmode="decimal" aria-label="${esc(rot)}">` : `<strong style="text-align:right">${v}</strong>`}<small>${esc(uso)}</small>`;
  const a = D.analise.cfg || {};
  return `<h2>Premissas</h2><p class="sub">${demo ? "Na demonstração, altere à vontade: Painel, Carteira, Meta e Projeção recalculam na hora." : "Origem: dados/config.json — edite lá e rode atualizar.bat."}</p>
    <div class="grade g2">
      <div class="card param">
        ${campo("meta_mensal_por_ativo", "Meta mensal padrão por ativo (R$)", NF[2].format(c.meta_mensal_por_ativo), "Renda que cada ativo deve pagar por mês")}
        ${campo("prazo_meta_meses", "Prazo padrão da meta (meses)", c.prazo_meta_meses || 60, "Em quanto tempo cada ativo deve atingir a meta")}
        ${campo("aporte_mensal", "Aporte mensal (R$)", NF[2].format(c.aporte_mensal), "Quanto você investe por mês")}
        ${campo("peso_ideal_renda_variavel", "Peso ideal da renda variável (%)", NF[2].format(c.peso_ideal_renda_variavel * 100), "Ações + FIIs no patrimônio total")}
      </div>
      <div class="card param">
        <span>CDI anual</span><strong style="text-align:right">${pct(a.cdi)}</strong>
        <span>Inflação anual</span><strong style="text-align:right">${pct(a.inflacao)}</strong>
        <span>Horizonte</span><strong style="text-align:right">${a.horizonte} anos</strong>
        <span>Reinvestir proventos</span><strong style="text-align:right">${a.reinvestir ? "sim" : "não"}</strong>
        <span>λ do EWMA</span><strong style="text-align:right">${a.lambda}</strong>
        <span>Simulações Monte Carlo</span><strong style="text-align:right">${num(a.simulacoes)}</strong>
      </div>
    </div>
    ${secao("Outras classes")}
    ${tabela("classes", [{ t: "Classe", v: o => o.classe, txt: true },
      { t: "Valor", v: o => o.valor, html: (o) => demo ? `<input data-classe="${esc(o.classe)}" value="${NF[2].format(o.valor)}" inputmode="decimal" aria-label="Valor ${esc(o.classe)}">` : brl(o.valor, 0) },
      { t: "Peso ideal no patrimônio", v: o => o.peso_ideal, f: pct }], c.outras_classes)}`;
};

RENDER.verificar = () => {
  const av = D.avisos;
  const n = g => av.filter(a => a[0] === g).length;
  return `<h2>Verificar</h2><p class="sub">Inconsistências encontradas nos dados e na análise quantitativa: ${n("ERRO")} erro(s), ${n("AVISO")} aviso(s), ${n("INFO")} informação(ões).</p>
    ${av.length ? `<ul class="avisos">${av.map(([g, t, m]) => `<li class="${esc(g)}"><span class="tag">${esc(g)}</span><strong>${esc(t)}</strong> — ${esc(m)}</li>`).join("")}</ul>`
      : `<div class="msg ok">Nenhuma inconsistência encontrada.</div>`}`;
};

// ------------------------------------------------------------------ navegação e eventos

function menu() {
  const avisos = D.avisos.filter(a => a[0] !== "INFO").length;
  $("#menu").innerHTML = ABAS.map(([id, rot]) => `<button data-aba="${id}" class="${aba === id ? "ativo" : ""}">${rot}${id === "verificar" && avisos ? `<span class="badge">${avisos}</span>` : ""}</button>`).join("");
}

function render() {
  menu();
  const r = RENDER[aba];
  $("#conteudo").innerHTML = r();
  if (r.depois) r.depois();
}

function ir(id) {
  if (!RENDER[id]) id = "painel";
  aba = id;
  if (location.hash !== "#" + id) history.replaceState(null, "", "#" + id);
  render();
  window.scrollTo({ top: 0 });
}

function eventos() {
  document.addEventListener("click", ev => {
    const el = ev.target.closest("[data-aba],[data-ir],[data-filtro],th[data-tab]");
    if (!el) return;
    if (el.dataset.aba || el.dataset.ir) return ir(el.dataset.aba || el.dataset.ir);
    if (el.dataset.filtro) { filtro = el.dataset.filtro; return render(); }
    const id = el.dataset.tab, i = Number(el.dataset.col), o = tabela.ordem[id];
    tabela.ordem[id] = o && o.i === i ? { i, asc: !o.asc } : { i, asc: false };
    render();
  });
  document.addEventListener("change", ev => {
    const el = ev.target;
    if (el.dataset.meta) return aoAlterarMeta(el);
    if (el.id === "sel-quant") { RENDER.quant.sel = el.value; return RENDER.quant.depois(); }
    if (el.dataset.proj) {
      const v = lerNumero(el.value);
      if (!Number.isFinite(v)) return toast("Valor inválido");
      projParams[el.dataset.proj] = el.dataset.tipo === "pct" ? v / 100 : v;
      return render();
    }
    if (el.dataset.cfg) {
      const v = lerNumero(el.value);
      if (!Number.isFinite(v) || v < 0) return toast("Valor inválido");
      E.cfg[el.dataset.cfg] = el.dataset.cfg === "peso_ideal_renda_variavel" ? v / 100 : v;
      delete projParams.aporte;
      salvarDemo();
      return render();
    }
    if (el.dataset.classe) {
      const v = lerNumero(el.value), o = E.cfg.outras_classes.find(x => x.classe === el.dataset.classe);
      if (!Number.isFinite(v) || v < 0 || !o) return toast("Valor inválido");
      o.valor = v;
      salvarDemo();
      return render();
    }
  });
  document.addEventListener("input", ev => {
    const el = ev.target;
    if (!el.dataset.busca) return;
    RENDER[el.dataset.busca].busca = el.value;
    const pos = el.selectionStart;
    render();
    const novo = $(`[data-busca="${el.dataset.busca}"]`);
    novo.focus();
    novo.setSelectionRange(pos, pos);
  });
  document.addEventListener("click", ev => {
    if (ev.target.id === "proj-reset") { for (const k in projParams) delete projParams[k]; render(); }
  });
  window.addEventListener("hashchange", () => ir(location.hash.slice(1)));
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change", render);
  $("#novo").addEventListener("click", () => ir("lancar"));
  $("#restaurar").addEventListener("click", async () => {
    local.apagar();
    for (const k in projParams) delete projParams[k];
    await carregar();
    render();
    toast("Dados da demonstração restaurados");
  });
  $("#atualizar").addEventListener("click", async () => {
    const b = $("#atualizar");
    b.disabled = true;
    b.textContent = "⟳ Atualizando…";
    try {
      const r = await api({ acao: "atualizar" });
      if (!r.ok) throw new Error((r.erros || []).join("; ") || r.erro || "falha");
      await carregar();
      cabecalho();
      render();
      toast("Cotações, proventos e análise atualizados");
    } catch (e) {
      toast("Não foi possível atualizar: " + e.message, 7000);
    } finally {
      b.disabled = false;
      b.textContent = "⟳ Atualizar cotações";
    }
  });
}

function cabecalho() {
  const demo = D.modo === "demo";
  const n = E.lancamentos.length;
  $("#situacao").textContent = `${demo ? "Demonstração" : "Rede local"} · dados de ${D.gerado_em.split(" ")[0].split("-").reverse().join("/")} · ${D.ativos.length} ativos · ${n} lançamentos`;
  $("#faixa-demo").hidden = !demo;
  $("#atualizar").hidden = demo;
  const b = $("#baixar");
  if (D.planilha) { b.hidden = false; b.href = encodeURI(D.planilha); b.setAttribute("download", ""); }
  $("#rodape").textContent = `Gerado em ${D.gerado_em.split(" ")[0].split("-").reverse().join("/")} ${D.gerado_em.split(" ")[1] || ""}`;
}

(async function iniciar() {
  try {
    await carregar();
  } catch (e) {
    $("#conteudo").innerHTML = `<div class="msg erro">Não foi possível carregar os dados: ${esc(e.message)}. Rode <code>python gestao.py --web</code>.</div>`;
    $("#situacao").textContent = "sem dados";
    return;
  }
  cabecalho();
  eventos();
  ir(location.hash.slice(1) || "painel");
})();
