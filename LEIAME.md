# Gestão de Investimentos

Uma base só (`dados/`) gera a planilha Excel com painel. Cada informação é registrada uma única vez.

## Estrutura
```
dados/          ENTRADA — única fonte (edite só aqui)
  ativos.csv        cadastro: ticker, tipo, segmento, peso ideal (cotação e proventos 12m são automáticos)
  lancamentos.csv   compras e vendas
  proventos.csv     dividendos, JCP e rendimentos recebidos
  config.json       meta mensal por ativo, aporte mensal, peso da renda variável, outras classes
saida/          planilha gerada (Gestao de investimentos.xlsm, com macros)
google/         versão Google Planilhas (planilha + InvestERP.gs + passo a passo)
invest/         código
  dados.py          leitura e checagens
  mercado.py        cotações e proventos (Yahoo Finance)
  excel/            abas, painel, tabelas dinâmicas, macros (VBA) e estilo da planilha
  cadastro.py       formulário de lançamentos e importação
  google.py         versão Google Planilhas (GOOGLEFINANCE, QUERY, Apps Script)
arquivo/        planilha antiga e dados simulados antigos (backup)
gestao.py       comando único
atualizar.bat   duplo clique = atualizar tudo
```

## Registrar compras, vendas e proventos (aba Cadastro)
1. Abra `saida/Gestao de investimentos.xlsm` e clique em **Habilitar Conteúdo** (macros) na primeira vez.
2. No Painel, clique em **✚ Novo lançamento** (ou vá na aba Cadastro).
3. Escolha o **ticker** e informe a **quantidade de cotas**. Tipo, segmento, cotação atual e cotas que você tem
   aparecem sozinhos; **preço vazio = cotação atual**. Ticker novo: informe Tipo (AÇÃO/FII) e segmento.
4. Clique em **✔ Registrar lançamento**: grava na hora em `dados/lancamentos.csv` e na aba Lançamentos,
   atualiza Carteira, Painel e tabelas dinâmicas e limpa o formulário. Erros aparecem em vermelho no formulário.
5. Proventos: mesmo processo no formulário ao lado (quantidade vazia = suas cotas).
6. **⟳ Atualizar cotações** (Painel): salva, fecha, busca cotações e reabre a planilha.

Sem macros habilitadas, o que ficar preenchido no formulário é importado ao rodar `atualizar.bat`.
Renda fixa, internacional e cripto (ex.: CDB Nubank) ficam em `dados/config.json` → `outras_classes`.

## Uso
- **Duplo clique em `atualizar.bat`**: instala/confere as dependências, busca cotações e proventos, confere os dados e gera a planilha.
- Sem internet: `python gestao.py --offline`.
- Tabelas dinâmicas precisam do Excel instalado (Windows); sem ele use `--sem-dinamicas`.
- Feche a planilha no Excel antes de atualizar.

## Planilha
Menu no topo: Painel · Carteira · Análises · ✚ Lançar · Lançamentos · Proventos · Premissas · Verificar.
- **Carteira** é a lista única de ações/FIIs: posição, proventos, peso e meta de renda na mesma linha
  (ativos que você ainda não tem aparecem em cinza). A aba Ativos fica oculta, só como base.
- **Painel (tela inicial estilo sistema ERP):** menu lateral com todos os módulos, indicadores da carteira e da meta de renda, 8 gráficos (posição, lucro, ações x FIIs, renda x meta, aporte necessário, proventos por mês, proventos por ativo, peso atual x ideal) e alocação por classe. Botões no topo levam às demais abas; cada aba tem "◀ Painel" para voltar.
- Azul em fundo amarelo = entrada (vem de `dados/`); verde = vínculo com outra aba; preto = fórmula.
- Todas as tabelas têm filtro no cabeçalho (ex.: filtrar só FIIs).
- **Análises:** tabelas dinâmicas (carteira por tipo, proventos por ativo, lançamentos) e gráfico dinâmico com segmentação por Tipo; atualizam ao abrir o arquivo.
- Paleta executiva: azul = dado principal, cinza = referência; verde/vermelho só para ganho/perda.
- A aba **Verificar** lista inconsistências encontradas nos dados.

## Privacidade
- `.htaccess` impede o Apache/WAMP de servir esta pasta (ela fica dentro de `www`).
- Só os tickers são enviados ao Yahoo Finance.

## Google Planilhas / Google Drive
- `python gestao.py --google --drive` gera `google/` e copia para `G:\Meu Drive\InvestERP` (Google Drive para computador).
- No Drive: abra o .xlsx com Planilhas Google, salve como Planilhas Google e cole `InvestERP.gs` em
  Extensões › Apps Script (detalhes em `google/COMO_USAR_NO_GOOGLE.md`).
- No Google: cotações ao vivo (GOOGLEFINANCE), análises com QUERY e registro pelo menu InvestERP ou caixa "Registrar".
- A planilha do Google vira base própria: o que for lançado lá não volta sozinho para `dados/`.
