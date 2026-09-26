"""Versão Google Planilhas: planilha .xlsx sem VBA + código Apps Script + passo a passo, prontos para o Drive.

No Google: cotações ao vivo com GOOGLEFINANCE (reserva: último preço do Yahoo), análises com QUERY,
registro de compras/proventos pelo Apps Script (menu InvestERP ou caixa "Registrar") e barra de navegação
com links. A planilha no Drive passa a ser a base dos lançamentos feitos lá.
"""
import shutil
from pathlib import Path

from .caminhos import BASE
from .excel import construir

PASTA = BASE / "google"
PLANILHA = "InvestERP - Google Planilhas.xlsx"
SCRIPT = "InvestERP.gs"
GUIA = "COMO_USAR_NO_GOOGLE.md"
DRIVE_PADRAO = Path(r"G:\Meu Drive\InvestERP")

APPS_SCRIPT = r"""/**
 * InvestERP — Apps Script da planilha de investimentos (gerado por gestao.py).
 * Extensões › Apps Script › cole este código › salve › execute "configurar" uma vez e autorize.
 */
const MENU = [["Painel", "Painel"], ["Carteira", "Carteira"], ["Análises", "Análises"], ["Cadastro", "✚ Lançar"],
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

/** Linha 1 de cada aba: menu com links para as outras abas (o módulo aberto fica em azul). */
function montarNavegacao() {
  const ss = SpreadsheetApp.getActive();
  const gids = {};
  ss.getSheets().forEach(s => gids[s.getName()] = s.getSheetId());
  const nomes = MENU.map(m => m[0]);
  ss.getSheets().forEach(sh => {
    if (nomes.indexOf(sh.getName()) < 0) return;
    const marca = "◆ InvestERP        ";
    let texto = marca;
    const pos = [];
    MENU.forEach(([aba, rotulo]) => {
      if (!(aba in gids)) return;
      pos.push([aba, texto.length, texto.length + rotulo.length]);
      texto += rotulo + "     ";
    });
    const b = SpreadsheetApp.newRichTextValue().setText(texto);
    b.setTextStyle(0, marca.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("#FFFFFF").setFontSize(11).build());
    pos.forEach(([aba, i, f]) => {
      const ativo = aba === sh.getName();
      b.setLinkUrl(i, f, "#gid=" + gids[aba]);
      b.setTextStyle(i, f, SpreadsheetApp.newTextStyle().setUnderline(false).setBold(ativo)
        .setForegroundColor(ativo ? "#60A5FA" : "#CBD5E1").setFontSize(10).build());
    });
    sh.getRange("B1").setRichTextValue(b.build());
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
"""

GUIA_TEXTO = """# InvestERP no Google Planilhas

## 1. Converter a planilha (uma vez)
1. Abra **drive.google.com** › pasta **InvestERP**.
2. Clique com o botão direito em **{planilha}** › **Abrir com** › **Planilhas Google**.
3. Na planilha aberta: **Arquivo › Salvar como Planilhas Google**. Use essa cópia daqui em diante
   (pode apagar o .xlsx do Drive depois).

## 2. Ativar o registro de compras (uma vez)
1. Na planilha: **Extensões › Apps Script**.
2. Apague o conteúdo do editor, cole todo o arquivo **{script}** e clique em **Salvar**.
3. Escolha a função **configurar** e clique em **Executar** › autorize com sua conta Google.
4. Volte para a planilha e recarregue a página: aparece o menu **InvestERP** e a barra de navegação
   da linha 1 vira links.

## 3. Uso
- **✚ Lançar**: escolha o ticker, informe a quantidade e **marque a caixa ao lado de "Registrar"**
  (ou menu **InvestERP › Registrar lançamento**). Preço vazio = cotação atual.
- Ticker novo: preencha também Tipo (AÇÃO/FII) e segmento — ele entra numa das {vagas} vagas do catálogo.
- Proventos: formulário ao lado, mesma forma (quantidade vazia = suas cotas).
- **Cotações** são atualizadas sozinhas pelo GOOGLEFINANCE. **Proventos 12m por cota** vêm da última
  geração no computador (aba Ativos, oculta: menu Ver › Páginas ocultas).
- **Análises** usa QUERY (equivalente às tabelas dinâmicas do Excel).

## Observações
- A planilha do Google passa a ser a base do que for lançado nela; a versão Excel do computador
  (pasta dados/) não é sincronizada automaticamente.
- Não compartilhe a planilha publicamente: ela contém seus dados financeiros.
"""


def gerar(ativos, lanc, prov, cfg, avisos, drive: Path | None = None):
    """Gera o pacote em google/ e, se drive for informado, copia para lá. Retorna mensagens."""
    PASTA.mkdir(exist_ok=True)
    construir(ativos, lanc, prov, cfg, avisos, google=True).save(PASTA / PLANILHA)
    (PASTA / SCRIPT).write_text(APPS_SCRIPT.lstrip(), encoding="utf-8")
    from .excel import VAGAS_GOOGLE
    (PASTA / GUIA).write_text(GUIA_TEXTO.format(planilha=PLANILHA, script=SCRIPT, vagas=VAGAS_GOOGLE), encoding="utf-8")
    msgs = [f"pacote Google em {PASTA}"]
    if drive:
        drive.mkdir(parents=True, exist_ok=True)
        for nome in (PLANILHA, SCRIPT, GUIA):
            shutil.copy2(PASTA / nome, drive / nome)
        msgs.append(f"copiado para o Google Drive: {drive}")
    return msgs
