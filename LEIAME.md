# Gestão de Investimentos

Uma base só (`dados/`) gera a planilha Excel com painel. Cada informação é registrada uma única vez.

## Estrutura
```
dados/          ENTRADA — única fonte (edite só aqui)
  ativos.csv        cadastro: ticker, tipo, segmento, peso ideal (cotação e proventos 12m são automáticos)
  lancamentos.csv   compras e vendas
  proventos.csv     dividendos, JCP e rendimentos recebidos
  config.json       meta mensal por ativo, aporte mensal, peso da renda variável, outras classes, analise
  historico.csv     AUTOMÁTICO: cotações e proventos semanais de 5 anos + Ibovespa (base da análise quantitativa)
saida/          planilha gerada (Gestao de investimentos.xlsm, com macros)
google/         versão Google Planilhas (planilha + InvestERP.gs + passo a passo)
invest/         código em camadas (Clean Architecture: cada camada só depende das de dentro)
  config.py         caminhos e leitura de dados/config.json
  dominio/          regras puras — sem arquivos, internet ou Excel (testáveis isoladamente)
    carteira.py       posições, fluxos de caixa, TIR
    lancamentos.py    regras para aceitar compra/venda/provento digitados no Cadastro
    validacao.py      conferência da base (vira a aba Verificar)
    analise.py        risco/retorno, qualidade dos dados, regressão, carteira real, parâmetros da projeção
    modelos.py        derivadas polinomiais, EWMA (λ), estatística, CAPM, Fisher, autovalores, Markowitz, Monte Carlo
    avisos.py, base.py  tipos compartilhados (Aviso, BaseDados)
  infra/            mundo externo
    repositorio.py    leitura tipada e gravação dos CSV de dados/
    yahoo.py          único ponto que acessa a internet (Yahoo Finance)
    mercado.py        atualiza cotação e proventos 12m em ativos.csv
    historico.py      histórico semanal em dados/historico.csv
  aplicacao/        casos de uso
    importacao.py     importar o que ficou preenchido no Cadastro (uso sem macros)
    pipeline.py       atualizar o mercado; conferir e analisar a base
  apresentacao/     saídas
    excel/            abas, painel, formulário do Cadastro, análises, tabelas dinâmicas, macros (VBA), estilo
    google.py         versão Google Planilhas (GOOGLEFINANCE, QUERY, Apps Script)
tests/          testes automáticos (python -m unittest discover -s tests -t .)
arquivo/        planilha antiga e dados simulados antigos (backup)
gestao.py       comando único (raiz de composição: único arquivo que liga as camadas)
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
- Testes automáticos (não usam internet nem seus dados): `python -m unittest discover -s tests -t .`

## Planilha
Menu no topo: Painel · Carteira · Análises · Meta · Quant · Modelos · Projeção · ✚ Lançar · Lançamentos · Proventos · Premissas · Verificar.
- **Carteira** é a lista única de ações/FIIs: posição, proventos, peso e meta de renda na mesma linha
  (ativos que você ainda não tem aparecem em cinza). A aba Ativos fica oculta, só como base.
- **Painel (padrão Casa Organizada, tema claro):** faixa azul-marinho com a situação dos dados e o botão **✚ Novo lançamento**; menu de abas com a aba aberta sublinhada; painel **Filtros** à esquerda (Todos / Ações / FIIs, atalhos e progresso da meta); cards de indicadores; 6 gráficos (proventos por mês, posição por tipo e por ativo, patrimônio por classe, progresso da meta, proventos por ativo); meta de renda, destaques e alocação por classe.
- Azul em fundo amarelo = entrada (vem de `dados/`); verde = vínculo com outra aba; preto = fórmula.
- Todas as tabelas têm filtro no cabeçalho (ex.: filtrar só FIIs).
- **Análises:** tabelas dinâmicas (carteira por tipo, proventos por ativo, lançamentos) e gráfico dinâmico com segmentação por Tipo; atualizam ao abrir o arquivo.
- Paleta executiva: azul = dado principal, cinza = referência; verde/vermelho só para ganho/perda.
- A aba **Verificar** lista inconsistências encontradas nos dados — inclusive as da análise quantitativa
  (cotações espúrias do Yahoo corrigidas, provento do cadastro diferente do histórico, queda forte de proventos,
  tendência sem confiabilidade estatística, ativo que concentra o risco).

## Meta de renda por ativo (aba Meta)
Você decide, para cada ação e cada FII, **quanto quer receber por mês** e **em quantos meses** (campos amarelos).
Padrão para todos em `dados/config.json` (`meta_mensal_por_ativo`: R$ 1.099, `prazo_meta_meses`: 60) ou na aba
Premissas. Para personalizar, digite por cima do campo amarelo (fica em negrito); apague para voltar ao padrão.
As personalizações são salvas em `dados/metas.csv` ao clicar em **⟳ Atualizar** (ou rodar `atualizar.bat`).

Por ativo a aba calcula: valor da cota, provento mensal por cota, **cotas necessárias**, **valor total a investir**,
cotas que você tem, preço médio e valor investido hoje, renda atual, % da meta, **cotas faltantes**, **valor faltante**,
**cotas por mês** e **aporte mensal** para cumprir o prazo (também reinvestindo os proventos) e a data prevista.
No rodapé: aporte que todos os prazos exigem × aporte mensal disponível (sobra ou falta) e subtotais de ações e FIIs.

## Análise quantitativa (abas Quant, Modelos e Projeção)
Calculada a cada execução a partir de `dados/historico.csv` e dos seus lançamentos. Os pesos usados são os da
**carteira real** (lançamentos × cotação); sem lançamentos, os pesos ideais.
- **Quant** (uma linha por ativo, mesma ordem da Carteira): qualidade dos dados, retorno 12m, volatilidade
  (histórica e EWMA λ), Sharpe, Sortino, queda máxima, VaR/CVaR 95%, beta e correlação com o Ibovespa,
  CAPM e alfa de Jensen, assimetria, curtose, teste de normalidade Jarque-Bera, autocorrelação,
  **derivadas analíticas** de um polinômio de grau 3 ajustado ao log-preço (taxa instantânea f′ e curvatura f″),
  regressão log-linear de 3 anos com R² e faixa de preço 12m, crescimento dos proventos e contribuição ao risco.
  Abaixo: resumo da carteira (volatilidade com correlações, diversificação, beta, DY, TIR dos seus lançamentos)
  e matriz de correlação.
- **Modelos**: retorno real (equação de Fisher), autovalores/autovetores da matriz de covariância (fatores de
  risco), carteiras de Markowitz (mínima variância e máximo Sharpe, sem venda a descoberto) comparadas com os
  pesos ideais e a carteira real, e Monte Carlo (movimento browniano geométrico) com percentis por ano,
  probabilidade de vencer o CDI e de atingir a meta de renda.
- **Projeção**: evolução mês a mês com fórmulas (parte da Carteira e do aporte das Premissas) em 3 cenários +
  CDI; os parâmetros em amarelo podem ser alterados na própria aba.
- Parâmetros em `dados/config.json` → `analise`: `cdi_anual`, `inflacao_anual`, `horizonte_anos`,
  `reinvestir_proventos`, `historico_anos`, `lambda_ewma`, `simulacoes_monte_carlo`.
- Estimativas estatísticas a partir do passado — não são recomendação nem garantia de retorno.

## Privacidade
- `.htaccess` impede o Apache/WAMP de servir esta pasta (ela fica dentro de `www`).
- Só os tickers são enviados ao Yahoo Finance.

## Google Planilhas / Google Drive
- `python gestao.py --google --drive` gera `google/` e copia para `G:\Meu Drive\InvestERP` (Google Drive para computador).
- No Drive: abra o .xlsx com Planilhas Google, salve como Planilhas Google e cole `InvestERP.gs` em
  Extensões › Apps Script (detalhes em `google/COMO_USAR_NO_GOOGLE.md`).
- No Google: cotações ao vivo (GOOGLEFINANCE), análises com QUERY e registro pelo menu InvestERP ou caixa "Registrar".
- A planilha do Google vira base própria: o que for lançado lá não volta sozinho para `dados/`.
