/**
 * InvestERP — Apps Script da planilha de investimentos (gerado por gestao.py).
 * Extensões › Apps Script › cole este código › salve › execute "configurar" uma vez e autorize.
 */
const MENU = [["Painel", "Painel"], ["Carteira", "Carteira"], ["Análises", "Análises"], ["Meta", "Meta"], ["Quant", "Quant"],
              ["Modelos", "Modelos"], ["Projeção", "Projeção"], ["Cadastro", "✚ Lançar"],
              ["Lançamentos", "Lançamentos"], ["Proventos", "Proventos"], ["Premissas", "Premissas"],
              ["Verificar", "Verificar"]];
const CAMPOS_LC = ["lc_data", "lc_ticker", "lc_op", "lc_qtd", "lc_preco", "lc_custos", "lc_tipo", "lc_seg"];
const CAMPOS_PV = ["pv_data", "pv_ticker", "pv_tipo", "pv_valor", "pv_qtd"];
const TIPOS_PROVENTO = ["DIVIDENDO", "JCP", "RENDIMENTO", "AMORTIZAÇÃO"];

function onOpen() {
  SpreadsheetApp.getUi().createMenu("InvestERP")
    .addItem("✚ Registrar lançamento", "registrarLancamento")
    .addItem("◈ Registrar provento", "registrarProvento")
    .addSeparator()
    .addItem("Limpar lançamento", "limparLancamento")
    .addItem("Limpar provento", "limparProvento")
    .addSeparator()
    .addItem("Refazer barra de navegação", "configurar")
    .addToUi();
  configurar();
}

function configurar() {
  montarNavegacao();
  garantirCaixas();
}

/** Linha 2 de cada aba: menu com links para as outras abas (o módulo aberto fica em negrito). */
function montarNavegacao() {
  const ss = SpreadsheetApp.getActive();
  const gids = {};
  ss.getSheets().forEach(s => gids[s.getName()] = s.getSheetId());
  const nomes = MENU.map(m => m[0]);
  ss.getSheets().forEach(sh => {
    if (nomes.indexOf(sh.getName()) < 0) return;
    let texto = "   ";
    const pos = [];
    MENU.forEach(([aba, rotulo]) => {
      if (!(aba in gids)) return;
      pos.push([aba, texto.length, texto.length + rotulo.length]);
      texto += rotulo + "     ";
    });
    const b = SpreadsheetApp.newRichTextValue().setText(texto);
    pos.forEach(([aba, i, f]) => {
      const ativo = aba === sh.getName();
      b.setLinkUrl(i, f, "#gid=" + gids[aba]);
      b.setTextStyle(i, f, SpreadsheetApp.newTextStyle().setUnderline(false).setBold(ativo)
        .setForegroundColor(ativo ? "#14213D" : "#64748B").setFontSize(10).build());
    });
    sh.getRange("A2").setRichTextValue(b.build());
  });
}

/** Caixas de seleção "Registrar" no Cadastro (marcar = registrar). */
function garantirCaixas() {
  const ss = SpreadsheetApp.getActive();
  ["lc_registrar", "pv_registrar"].forEach(n => {
    const r = ss.getRangeByName(n);
    if (r && !r.getDataValidation()) { r.insertCheckboxes(); r.setValue(false); }
  });
}

function onEdit(e) {
  const ss = e.source, r = e.range;
  [["lc_registrar", registrarLancamento], ["pv_registrar", registrarProvento]].forEach(([nome, acao]) => {
    const alvo = ss.getRangeByName(nome);
    if (alvo && mesma(r, alvo) && r.getValue() === true) { r.setValue(false); acao(); }
  });
}

function mesma(a, b) {
  return a.getSheet().getSheetId() === b.getSheet().getSheetId() && a.getRow() === b.getRow() && a.getColumn() === b.getColumn();
}

function campo(n) { return SpreadsheetApp.getActive().getRangeByName(n); }
function valor(n) { return campo(n).getValue(); }
function texto(n) { return String(valor(n)).trim().toUpperCase(); }
function numero(n) {
  const x = valor(n);
  if (x === "" || x === null) return 0;
  if (typeof x === "number") return x;
  const v = parseFloat(String(x).replace("R$", "").replace(/\./g, "").replace(",", ".").trim());
  return isNaN(v) ? 0 : v;
}
function data(n) { const d = valor(n); return d instanceof Date ? d : new Date(); }
function mensagem(n, t, ok) { campo(n).setValue(t).setFontColor(ok ? "#16A34A" : "#DC2626"); }
function limpar(nomes) { nomes.forEach(n => campo(n).clearContent()); }
function reais(v) { return "R$ " + v.toFixed(2).replace(".", ","); }

/** Próxima linha livre de uma aba de registros (dados a partir da linha 5, ticker na coluna B). */
function proximaLinha(sh) {
  const vals = sh.getRange("B5:B" + Math.max(sh.getLastRow(), 5)).getValues();
  let ultima = 4;
  vals.forEach((l, i) => { if (String(l[0]).trim() !== "") ultima = 5 + i; });
  return ultima + 1;
}

/** Linha do ticker na aba Ativos, ou a primeira vaga livre (-vaga) para ticker novo. */
function linhaAtivo(tk) {
  const at = SpreadsheetApp.getActive().getSheetByName("Ativos");
  const lista = at.getRange("A5:A" + at.getLastRow()).getValues().map(l => String(l[0]).trim().toUpperCase());
  const fim = lista.indexOf("TOTAL") >= 0 ? lista.indexOf("TOTAL") : lista.length;
  const achou = lista.slice(0, fim).indexOf(tk);
  if (achou >= 0) return 5 + achou;
  const vaga = lista.slice(0, fim).indexOf("");
  return vaga >= 0 ? -(5 + vaga) : 0;
}

function registrarLancamento() {
  const ss = SpreadsheetApp.getActive();
  const tk = texto("lc_ticker"), op = texto("lc_op") || "COMPRA";
  const qtd = numero("lc_qtd"), custos = numero("lc_custos"), tem = numero("lc_tem");
  let preco = numero("lc_preco");
  if (preco <= 0) preco = numero("lc_cot");            // valor unitário já estabelecido no cadastro do ativo
  if (!tk) return mensagem("lc_status", "Informe o ticker.", false);
  if (op !== "COMPRA" && op !== "VENDA") return mensagem("lc_status", "Operação deve ser COMPRA ou VENDA.", false);
  if (qtd <= 0) return mensagem("lc_status", "Informe a quantidade de cotas.", false);
  const linha = linhaAtivo(tk);
  if (linha === 0) return mensagem("lc_status", "Sem vagas para ticker novo: gere a planilha de novo no computador.", false);
  if (linha < 0) {                                      // ticker novo: ocupa uma vaga do catálogo
    const tipo = texto("lc_tipo");
    if (tipo !== "AÇÃO" && tipo !== "FII") return mensagem("lc_status", "Ticker novo: informe o Tipo do ativo (AÇÃO ou FII).", false);
    if (preco <= 0) return mensagem("lc_status", "Ticker novo: informe o preço por cota.", false);
    const at = ss.getSheetByName("Ativos"), r = -linha;
    at.getRange(r, 1, 1, 3).setValues([[tk, tipo, String(valor("lc_seg")).trim() || "-"]]);
    at.getRange(r, 5, 1, 3).setValues([[0, 0, "manual"]]);
    at.getRange(r, 8).setValue(preco);
  }
  if (preco <= 0) return mensagem("lc_status", "Sem cotação para este ticker: informe o preço por cota.", false);
  if (op === "VENDA" && qtd > tem) return mensagem("lc_status", "Venda maior que as cotas que você tem (" + tem + ").", false);

  const sh = ss.getSheetByName("Lançamentos"), r = proximaLinha(sh);
  sh.getRange(r, 1, 1, 6).setValues([[data("lc_data"), tk, op, qtd, preco, custos]]);
  sh.getRange(r, 7).setFormula("=D" + r + "*E" + r + "+F" + r);
  sh.getRange(r, 1).setNumberFormat("dd/mm/yyyy");
  limpar(CAMPOS_LC);
  mensagem("lc_status", "✔  Registrado: " + op + " de " + qtd + " " + tk + " a " + reais(preco) +
           (linha < 0 ? "  •  ticker novo incluído na carteira" : ""), true);
}

function registrarProvento() {
  const ss = SpreadsheetApp.getActive();
  const tk = texto("pv_ticker"), tipo = texto("pv_tipo") || "DIVIDENDO", v = numero("pv_valor");
  let qtd = numero("pv_qtd");
  if (qtd <= 0) qtd = numero("pv_tem");
  if (!tk || linhaAtivo(tk) <= 0) return mensagem("pv_status", "Ticker não cadastrado: registre a compra primeiro.", false);
  if (TIPOS_PROVENTO.indexOf(tipo) < 0) return mensagem("pv_status", "Tipo de provento inválido.", false);
  if (v <= 0) return mensagem("pv_status", "Informe o valor por cota.", false);
  if (qtd <= 0) return mensagem("pv_status", "Você não tem cotas deste ativo: informe a quantidade.", false);
  const sh = ss.getSheetByName("Proventos"), r = proximaLinha(sh);
  sh.getRange(r, 1, 1, 5).setValues([[data("pv_data"), tk, tipo, v, qtd]]);
  sh.getRange(r, 6).setFormula("=D" + r + "*E" + r);
  sh.getRange(r, 1).setNumberFormat("dd/mm/yyyy");
  limpar(CAMPOS_PV);
  mensagem("pv_status", "✔  Registrado: " + tipo + " de " + tk + " — " + reais(v * qtd), true);
}

function limparLancamento() { limpar(CAMPOS_LC); campo("lc_status").clearContent(); }
function limparProvento() { limpar(CAMPOS_PV); campo("pv_status").clearContent(); }
