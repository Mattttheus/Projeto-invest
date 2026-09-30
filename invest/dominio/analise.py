"""Análise quantitativa: risco/retorno, derivadas da tendência, regressão e parâmetros da projeção.

Recebe o histórico semanal e os valores internos da carteira (lançamentos, proventos, cotações e premissas)
já carregados; tudo é calculado aqui e a planilha só apresenta (abas Quant, Modelos e Projeção).
A data de referência (`hoje`) é parâmetro: o resultado é reproduzível e testável.

Convenções:
- retornos semanais totais = (preço + provento da semana) / preço anterior - 1;
- janela de risco e regressão: últimos 3 anos (156 semanas); "12m" = últimas 52 semanas;
- derivadas: polinômio de grau 3 ajustado ao log-preço de 12 meses, derivado analiticamente hoje:
  1ª = taxa instantânea anual (velocidade), 2ª = curvatura (aceleração). Detalhes em modelos.py.
"""
import numpy as np
import pandas as pd

from . import modelos
from .avisos import AVISO, INFO, Aviso
from .carteira import fluxos_internos, posicoes, tir, valor_posicoes

ANO = 52              # semanas por ano
JANELA = 3 * ANO      # risco, beta e regressão
IBOV = "^BVSP"        # referência de mercado no histórico
MIN_SEMANAS_PAR = 26  # mínimo de semanas em comum para covariância/beta


def parametros(cfg):
    """Premissas da análise (config.json → analise), com valores padrão."""
    a = cfg.get("analise", {})
    return {"cdi": a.get("cdi_anual", 0.14), "horizonte": a.get("horizonte_anos", 10),
            "reinvestir": a.get("reinvestir_proventos", True), "inflacao": a.get("inflacao_anual", 0.045),
            "lambda": a.get("lambda_ewma", 0.94), "simulacoes": a.get("simulacoes_monte_carlo", 5000)}


def _anual(taxa_semana):
    return np.expm1(taxa_semana * ANO)


def _regressao(logp):
    """Tendência log-linear: retorna (inclinação/semana, R², desvio dos resíduos, valor ajustado hoje)."""
    x = np.arange(len(logp))
    b, a = np.polyfit(x, logp, 1)
    ajuste = a + b * x
    res = logp - ajuste
    tot = ((logp - logp.mean()) ** 2).sum()
    r2 = 1 - (res ** 2).sum() / tot if tot > 0 else 0.0
    return b, r2, res.std(ddof=2), a + b * x[-1]


def _leitura(vel, acel):
    if np.isnan(vel) or np.isnan(acel):
        return "-"
    if vel >= 0:
        return "Alta acelerando" if acel >= 0 else "Alta perdendo força"
    return "Queda acelerando" if acel < 0 else "Queda perdendo força"


def limpar(p, limite=0.3):
    """Remove cotações espúrias: salto > limite (em log) desfeito na semana seguinte (ex.: 109 → 1,06 → 109).
    Retorna a série corrigida (interpolada) e quantos pontos foram corrigidos."""
    lr = np.log(p).diff()
    pico = (lr.abs() > limite) & (lr.shift(-1).abs() > limite) & (np.sign(lr) != np.sign(lr.shift(-1)))
    if not pico.any():
        return p, 0
    return p.mask(pico).interpolate(limit_area="inside"), int(pico.sum())


def _qualidade(obs, dias, lacunas, corrigidos=0):
    if obs >= 2.5 * ANO and dias <= 10 and lacunas == 0 and corrigidos == 0:
        return "Alta"
    if obs >= ANO and dias <= 30:
        return "Média"
    return "Baixa"


def _max_queda(indice):
    return float((indice / indice.cummax() - 1).min())


def metricas_ativo(p, d, ibov_r, c, hoje):
    """p: fechamentos semanais; d: proventos por semana (mesmo índice); ibov_r: retornos semanais do Ibov."""
    p, corrigidos = limpar(p.dropna())
    d = d.reindex(p.index).fillna(0)
    m = {"obs": len(p), "ultima": p.index[-1] if len(p) else pd.NaT, "corrigidos": corrigidos}
    if len(p) < 8:
        return m
    m["dias"] = (hoje - p.index[-1]).days
    m["lacunas"] = int((p.index.to_series().diff().dt.days > 21).sum())   # semanas sem negociação
    m["qualidade"] = _qualidade(m["obs"], m["dias"], m["lacunas"], corrigidos)
    r = ((p + d) / p.shift(1) - 1).dropna()
    rj = r.iloc[-JANELA:]
    u = p.iloc[-ANO - 1:]
    m["ret12_preco"] = u.iloc[-1] / u.iloc[0] - 1
    m["ret12_total"] = (1 + r.iloc[-ANO:]).prod() - 1
    m["vol"] = rj.std() * np.sqrt(ANO)
    m["ret_anual"] = (1 + rj).prod() ** (ANO / len(rj)) - 1
    cdi = c["cdi"]
    m["sharpe"] = (m["ret_anual"] - cdi) / m["vol"] if m["vol"] > 0 else np.nan
    m["vol_ewma"] = modelos.vol_ewma(rj.values, c["lambda"])
    m["max_queda"] = _max_queda((1 + rj).cumprod())
    par = pd.concat([rj, ibov_r], axis=1, join="inner").dropna()
    if len(par) >= MIN_SEMANAS_PAR and par.iloc[:, 1].var() > 0:
        m["beta"] = par.cov().iloc[0, 1] / par.iloc[:, 1].var()
        m["corr_ibov"] = par.corr().iloc[0, 1]
    ibov_anual = (1 + ibov_r).prod() ** (ANO / len(ibov_r)) - 1 if len(ibov_r) >= ANO else np.nan
    m.update(modelos.estatisticas(rj, (1 + cdi) ** (1 / ANO) - 1, ibov_anual, cdi, m.get("beta", np.nan),
                                  m["ret_anual"]))
    # derivadas analíticas do polinômio ajustado ao log-preço (12 meses)
    lp = np.log(p)
    pol = modelos.derivadas_polinomiais(lp.iloc[-ANO:].values)
    m["velocidade"], m["aceleracao"], m["r2_poli"] = (pol.get(k, np.nan) for k in ("taxa_inst", "curvatura", "r2_poli"))
    m["leitura"] = _leitura(m["velocidade"], m["aceleracao"])
    # regressão (tendência de longo prazo) e projeção de 12 meses
    lj = lp.iloc[-JANELA:].values
    if len(lj) >= ANO:
        b, r2, s, aj = _regressao(lj)
        m["tendencia"], m["r2"] = _anual(b), r2
        m["vs_tendencia"] = p.iloc[-1] / np.exp(aj) - 1
        centro = aj + b * ANO
        m["proj12"], m["proj12_min"], m["proj12_max"] = (float(np.exp(centro + k * s)) for k in (0, -1, 1))
    # proventos: 12m (por data, como em infra/yahoo.py) e crescimento sobre os 12m anteriores
    um_ano, dois = hoje - pd.DateOffset(years=1), hoje - pd.DateOffset(years=2)
    fim_semana = d.index + pd.Timedelta(days=6)          # cada linha soma os proventos da semana inteira
    m["prov12"] = d[fim_semana > um_ano].sum()
    ant = d[(fim_semana > dois) & (fim_semana <= um_ano)].sum() if p.index[0] <= dois else np.nan
    m["cresc_prov"] = m["prov12"] / ant - 1 if ant and ant > 0 else np.nan
    m["retornos"] = rj
    return m


def analisar(ativos, lanc, prov, cfg, hist, hoje=None):
    """Retorna dict com: por_ativo (DataFrame por ticker), carteira (dict), correlacao (DataFrame),
    projecao (parâmetros), modelos e avisos [Aviso]. hoje: data de referência (padrão: hoje)."""
    hoje = pd.Timestamp(hoje or pd.Timestamp.today()).normalize()
    c = parametros(cfg)
    avisos = []
    if hist.empty:
        return {"vazio": True, "avisos": [Aviso(AVISO, "-", "Sem histórico de cotações (dados/historico.csv): "
                                           "rode com internet para gerar a análise quantitativa.")], "cfg": c}
    px = hist.pivot_table(index="data", columns="ticker", values="fechamento")
    dv = hist.pivot_table(index="data", columns="ticker", values="provento", aggfunc="sum").fillna(0)
    ibov_r = px[IBOV].dropna().pct_change().dropna().iloc[-JANELA:] if IBOV in px else pd.Series(dtype=float)

    linhas, rets = {}, {}
    for _, a in ativos.iterrows():
        t = a["ticker"]
        if t not in px:
            avisos.append(Aviso(AVISO, t, "Sem histórico de cotações — fica fora da análise quantitativa."))
            continue
        m = metricas_ativo(px[t], dv[t], ibov_r, c, hoje)
        if "retornos" in m:
            rets[t] = m.pop("retornos")
        linhas[t] = m
        # qualificação da informação: confronta o cadastro com o histórico
        if m.get("corrigidos"):
            avisos.append(Aviso(INFO, t, f"{m['corrigidos']} cotação(ões) espúria(s) do Yahoo no histórico foram ignoradas na análise."))
        if m.get("dias", 0) > 10:
            avisos.append(Aviso(AVISO, t, f"Última cotação do histórico tem {m['dias']} dias — ativo pouco negociado ou ticker mudou."))
        p12 = m.get("prov12", 0)
        if a["provento_anual_cota"] > 0 and p12 > 0 and abs(a["provento_anual_cota"] / p12 - 1) > 0.15:
            avisos.append(Aviso(AVISO, t, f"Provento 12m do cadastro (R$ {a['provento_anual_cota']:.2f}) difere do histórico "
                                       f"(R$ {p12:.2f}) — confira antes de usar a meta de renda."))
        if not np.isnan(m.get("cresc_prov", np.nan)) and m["cresc_prov"] < -0.3:
            avisos.append(Aviso(AVISO, t, f"Proventos 12m caíram {-m['cresc_prov']:.0%} sobre o ano anterior — o DY atual pode não se repetir."))
        if m.get("r2", 1) < 0.2 and "tendencia" in m:
            r2 = f"{m['r2']:.2f}".replace(".", ",")
            avisos.append(Aviso(INFO, t, f"Tendência de 3 anos com R² de {r2}: projeção por regressão pouco confiável."))
    por_ativo = pd.DataFrame.from_dict(linhas, orient="index")

    # valores internos: pesos da carteira real (ou ideais, se ainda não há lançamentos)
    qtd = posicoes(ativos, lanc)
    precos = ativos.set_index("ticker")["preco"]
    valor = valor_posicoes(ativos, qtd)
    if valor.sum() > 0:
        pesos, base = valor / valor.sum(), "carteira real (lançamentos)"
    else:
        pesos, base = ativos.set_index("ticker")["peso_ideal"], "pesos ideais (ainda sem lançamentos)"
        pesos = pesos / pesos.sum() if pesos.sum() > 0 else pesos
    R = pd.DataFrame(rets)
    tk = [t for t in pesos.index if t in R and pesos[t] > 0]
    w = pesos[tk] / pesos[tk].sum() if tk else pesos[tk]
    cov = R[tk].cov(min_periods=MIN_SEMANAS_PAR) * ANO
    carteira = {"base": base, "valor_rv": float(valor.sum())}
    if tk:
        cov = cov.fillna(0)
        var = float(w @ cov @ w)
        vol_p = np.sqrt(var)
        contrib = w * (cov @ w) / var if var > 0 else w * 0
        por_ativo["peso_analise"] = pesos.reindex(por_ativo.index).fillna(0)
        por_ativo["contrib_risco"] = contrib.reindex(por_ativo.index).fillna(0)
        met = por_ativo.loc[tk]
        media_vol = float((w * met["vol"]).sum())
        dy = (ativos.set_index("ticker")["provento_anual_cota"] / precos).replace([np.inf, -np.inf], 0).fillna(0)
        carteira.update({
            "vol": vol_p, "ret_anual": float((w * met["ret_anual"]).sum()),
            "ret12": float((w * met["ret12_total"]).sum()),
            "beta": float((w * met["beta"].fillna(1)).sum()) if "beta" in met else np.nan,
            "dy": float((w * dy.reindex(tk)).sum()),
            "tendencia": float((w * met["tendencia"].fillna(0)).sum()) if "tendencia" in met else 0.0,
            "r2": float((w * met["r2"].fillna(0)).sum()) if "r2" in met else 0.0,
            "diversificacao": media_vol / vol_p if vol_p > 0 else np.nan,
            "max_queda": _max_queda((1 + (R[tk].fillna(0) @ w)).cumprod()),
        })
        carteira["sharpe"] = (carteira["ret_anual"] - c["cdi"]) / vol_p if vol_p > 0 else np.nan
        if len(ibov_r) >= ANO:
            carteira["ibov12"] = float((1 + ibov_r.iloc[-ANO:]).prod() - 1)
        for t in tk:
            if w[t] > 0.05 and contrib[t] > 2 * w[t]:
                avisos.append(Aviso(INFO, t, f"Responde por {contrib[t]:.0%} do risco com {w[t]:.0%} do peso — maior fonte de volatilidade."))
    carteira["tir"] = tir(fluxos_internos(ativos, lanc, prov, qtd, hoje)) if not lanc.empty else None

    # projeção: valorização base = tendência encolhida pelo R² (sem ajuste, a reta vale ~0) e limitada;
    # cenários = base ± 1 desvio-padrão do retorno médio no horizonte (σ/√H)
    H = c["horizonte"]
    base_g = float(np.clip(carteira.get("tendencia", 0) * carteira.get("r2", 0), -0.10, 0.20))
    banda = carteira.get("vol", 0.20) / np.sqrt(H)
    projecao = {"horizonte": H, "dy": carteira.get("dy", 0.0), "cdi": c["cdi"], "reinvestir": c["reinvestir"],
                "g_base": base_g, "g_pess": base_g - banda, "g_otim": base_g + banda}

    todos = [t for t in ativos["ticker"] if t in R]
    corr = R[todos].corr(min_periods=MIN_SEMANAS_PAR)
    mod = _modelos(R[todos], por_ativo, w, carteira, projecao, ativos, cfg, c) if len(todos) >= 2 else None
    return {"vazio": False, "por_ativo": por_ativo, "carteira": carteira, "correlacao": corr,
            "projecao": projecao, "modelos": mod, "avisos": avisos, "cfg": c,
            "data_base": hist["data"].max()}


def _modelos(R, por_ativo, w, carteira, projecao, ativos, cfg, c):
    """Álgebra linear (fatores, Markowitz), economia (Fisher) e Monte Carlo sobre a carteira."""
    cov = R.cov(min_periods=MIN_SEMANAS_PAR).fillna(0) * ANO
    # retorno esperado: metade histórico, metade CAPM (encolhimento reduz o erro de estimação de μ)
    hist_mu = por_ativo["ret_anual"].reindex(cov.index)
    capm = por_ativo["capm"].reindex(cov.index) if "capm" in por_ativo else hist_mu
    mu = (0.5 * hist_mu + 0.5 * capm.fillna(hist_mu)).fillna(0)
    rf = c["cdi"]
    otimas = modelos.markowitz(cov, mu, rf)
    ideal = ativos.set_index("ticker")["peso_ideal"].reindex(cov.index).fillna(0)
    carteiras = {"Pesos ideais": ideal / ideal.sum() if ideal.sum() > 0 else ideal,
                 "Carteira analisada": w.reindex(cov.index).fillna(0)}
    carteiras.update({"Mínima variância": otimas["minima_variancia"], "Máximo Sharpe": otimas["maximo_sharpe"]})
    carteiras = {k: v for k, v in carteiras.items() if v is not None and v.sum() > 0}
    metricas = {k: modelos.risco_retorno(v, cov, mu, rf) for k, v in carteiras.items()}

    meta = cfg.get("meta_mensal_por_ativo", 0) * len(ativos)
    mu_mc = projecao["g_base"] + (projecao["dy"] if c["reinvestir"] else 0)
    mc = modelos.monte_carlo(carteira.get("valor_rv", 0), cfg.get("aporte_mensal", 0), mu_mc,
                             carteira.get("vol", 0.2), c["horizonte"], meta, projecao["dy"], rf, n=c["simulacoes"])
    economia = {"inflacao": c["inflacao"], "cdi_real": modelos.real(rf, c["inflacao"]),
                "carteira_real": modelos.real(carteira.get("ret_anual", 0), c["inflacao"]),
                "dy_real": modelos.real(carteira.get("dy", 0), c["inflacao"]),
                "premio_risco": carteira.get("ret_anual", 0) - rf}
    return {"fatores": modelos.fatores(cov), "mu": mu, "carteiras": carteiras, "metricas": metricas,
            "monte_carlo": mc, "economia": economia}
