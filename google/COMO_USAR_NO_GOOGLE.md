# InvestERP no Google Planilhas

## 1. Converter a planilha (uma vez)
1. Abra **drive.google.com** › pasta **InvestERP**.
2. Clique com o botão direito em **InvestERP - Google Planilhas.xlsx** › **Abrir com** › **Planilhas Google**.
3. Na planilha aberta: **Arquivo › Salvar como Planilhas Google**. Use essa cópia daqui em diante
   (pode apagar o .xlsx do Drive depois).

## 2. Ativar o registro de compras (uma vez)
1. Na planilha: **Extensões › Apps Script**.
2. Apague o conteúdo do editor, cole todo o arquivo **InvestERP.gs** e clique em **Salvar**.
3. Escolha a função **configurar** e clique em **Executar** › autorize com sua conta Google.
4. Volte para a planilha e recarregue a página: aparece o menu **InvestERP** e a barra de navegação
   da linha 2 vira links.

## 3. Uso
- **✚ Lançar**: escolha o ticker, informe a quantidade e **marque a caixa ao lado de "Registrar"**
  (ou menu **InvestERP › Registrar lançamento**). Preço vazio = cotação atual.
- Ticker novo: preencha também Tipo (AÇÃO/FII) e segmento — ele entra numa das 15 vagas do catálogo.
- Proventos: formulário ao lado, mesma forma (quantidade vazia = suas cotas).
- **Cotações** são atualizadas sozinhas pelo GOOGLEFINANCE. **Proventos 12m por cota** vêm da última
  geração no computador (aba Ativos, oculta: menu Ver › Páginas ocultas).
- **Análises** usa QUERY (equivalente às tabelas dinâmicas do Excel).

## Observações
- A planilha do Google passa a ser a base do que for lançado nela; a versão Excel do computador
  (pasta dados/) não é sincronizada automaticamente.
- Não compartilhe a planilha publicamente: ela contém seus dados financeiros.
