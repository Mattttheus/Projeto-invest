"""Modelos matemáticos da análise: cálculo, álgebra linear, estatística e processos estocásticos.

- Cálculo: regressão polinomial (grau 3) do log-preço e suas derivadas analíticas no ponto atual;
  volatilidade EWMA com fator de decaimento λ (RiskMetrics: σ²ₜ = λσ²ₜ₋₁ + (1-λ)r²ₜ).
- Estatística: assimetria, curtose, teste de normalidade Jarque-Bera, autocorrelação, VaR/CVaR históricos.
- Finanças/economia: CAPM, alfa de Jensen, Sortino, retorno real pela equação de Fisher.
- Álgebra linear: autovalores/autovetores da matriz de covariância (fatores de risco, PCA) e
  carteiras de Markowitz (mínima variância e máximo Sharpe) sem venda a descoberto.
- Processo estocástico: Monte Carlo de movimento browniano geométrico com aportes mensais.
Só numpy/pandas.
"""
import numpy as np
import pandas as pd

ANO = 52


# ------------------------------------------------------------------ cálculo

def derivadas_polinomiais(logp, grau=3):
    """Ajusta log-preço(t) por polinômio (t em anos, 0 = hoje) e deriva analiticamente em t = 0.
    Retorna taxa instantânea anual (e^{f'} - 1), curvatura f'' (por ano²) e R² do ajuste."""
    y = np.asarray(logp, float)
    if len(y) < grau + 5:
        return {}
    t = (np.arange(len(y)) - (len(y) - 1)) / ANO
    pol = np.polynomial.Polynomial.fit(t, y, grau).convert()
    d1, d2 = pol.deriv(1), pol.deriv(2)
    res = y - pol(t)
    tot = ((y - y.mean()) ** 2).sum()
    return {"taxa_inst": float(np.expm1(d1(0.0))), "curvatura": float(d2(0.0)),
            "r2_poli": float(1 - (res ** 2).sum() / tot) if tot > 0 else 0.0}


def vol_ewma(r, lam=0.94):
    """Volatilidade anual EWMA: σ²ₜ = λσ²ₜ₋₁ + (1-λ)r²ₜ. Como no RiskMetrics, média zero: começa em
    média(r²) das 10 primeiras semanas (e não na variância com a média subtraída)."""
    r = np.asarray(r, float)
    if len(r) < 10:
        return np.nan
    v = float((r[:10] ** 2).mean())
    for x in r:
        v = lam * v + (1 - lam) * x * x
    return float(np.sqrt(v * ANO))


# ------------------------------------------------------------------ estatística

def estatisticas(r, rf_semana, ibov_ret_anual, rf_anual, beta, ret_anual):
    r = pd.Series(r).dropna()
    if len(r) < 26:
        return {}
    x = r - r.mean()
    s = x.std(ddof=0)
    assim = float((x ** 3).mean() / s ** 3) if s > 0 else 0.0
    curt = float((x ** 4).mean() / s ** 4 - 3) if s > 0 else 0.0
    jb = len(r) / 6 * (assim ** 2 + curt ** 2 / 4)
    corte = r.quantile(0.05)
    abaixo = r[r < rf_semana] - rf_semana
    desvio_neg = np.sqrt((abaixo ** 2).sum() / len(r)) * np.sqrt(ANO)
    m = {"assimetria": assim, "curtose": curt, "jb_p": float(np.exp(-jb / 2)),     # χ² com 2 g.l.
         "autocorr": float(r.autocorr(1)), "var95": float(-corte), "cvar95": float(-r[r <= corte].mean()),
         "sortino": (ret_anual - rf_anual) / desvio_neg if desvio_neg > 0 else np.nan}
    if not np.isnan(beta) and not np.isnan(ibov_ret_anual):
        m["capm"] = rf_anual + beta * (ibov_ret_anual - rf_anual)
        m["alfa"] = ret_anual - m["capm"]
    return m


def real(nominal, inflacao):
    """Equação de Fisher: (1 + nominal) = (1 + real)(1 + inflação)."""
    return (1 + nominal) / (1 + inflacao) - 1


# ------------------------------------------------------------------ álgebra linear

def fatores(cov):
    """Decomposição espectral Σ = VΛVᵀ: autovalores (variância de cada fator) e autovetores (cargas)."""
    val, vet = np.linalg.eigh(cov.values)
    ordem = np.argsort(val)[::-1]
    val, vet = val[ordem], vet[:, ordem]
    vet = vet * np.sign(vet.sum(axis=0))                  # sinal convencional: fator 1 positivo
    k = min(3, len(val))
    return {"autovalores": val, "explicado": val / val.sum(),
            "cargas": pd.DataFrame(vet[:, :k], index=cov.index, columns=[f"Fator {i + 1}" for i in range(k)]),
            "condicao": float(val[0] / val[-1]) if val[-1] > 0 else np.inf}


def _long_only(cov, alvo):
    """Resolve Σw ∝ alvo com w ≥ 0 e Σw = 1 (conjunto ativo: remove quem fica negativo e resolve de novo)."""
    ativos = list(cov.index)
    while ativos:
        S = cov.loc[ativos, ativos].values + np.eye(len(ativos)) * 1e-10
        w = np.linalg.solve(S, alvo.loc[ativos].values)
        if w.sum() <= 0:
            return None
        w = w / w.sum()
        if (w >= -1e-12).all():
            return pd.Series(w, index=ativos).reindex(cov.index).fillna(0).clip(lower=0)
        ativos = [a for a, x in zip(ativos, w) if x > 0]
    return None


def _simplex(v):
    """Projeção euclidiana no simplex {w ≥ 0, Σw = 1}."""
    u = np.sort(v)[::-1]
    css = np.cumsum(u) - 1
    k = np.nonzero(u - css / np.arange(1, len(v) + 1) > 0)[0][-1]
    return np.maximum(v - css[k] / (k + 1), 0)


def _media_variancia(S, mu, gama, iters=800):
    """max wᵀμ − (γ/2)wᵀΣw no simplex, por gradiente projetado (passo 1/L, L = γ·λmáx)."""
    passo = 1 / (gama * np.linalg.eigvalsh(S)[-1])
    w = np.full(len(mu), 1 / len(mu))
    for _ in range(iters):
        w = _simplex(w + passo * (mu - gama * S @ w))
    return w


def markowitz(cov, mu, rf):
    """Mínima variância: w ∝ Σ⁻¹1 (conjunto ativo). Máximo Sharpe: varre a fronteira eficiente
    (max wᵀμ − γ/2·wᵀΣw para vários γ) e fica com a maior razão (retorno − rf)/vol. Sem vendas a descoberto."""
    um = pd.Series(1.0, index=cov.index)
    S, m = cov.values, mu.reindex(cov.index).values
    melhor, w_best = -np.inf, None
    for gama in np.logspace(-0.5, 3, 30):
        w = _media_variancia(S, m, gama)
        vol = np.sqrt(w @ S @ w)
        sh = (w @ m - rf) / vol if vol > 0 else -np.inf
        if sh > melhor:
            melhor, w_best = sh, w
    tangente = pd.Series(np.where(w_best > 1e-4, w_best, 0), index=cov.index) if w_best is not None else None
    return {"minima_variancia": _long_only(cov, um),
            "maximo_sharpe": tangente / tangente.sum() if tangente is not None else None}


def risco_retorno(w, cov, mu, rf):
    w = w.reindex(cov.index).fillna(0)
    vol = float(np.sqrt(w @ cov @ w))
    ret = float(w @ mu)
    return {"ret": ret, "vol": vol, "sharpe": (ret - rf) / vol if vol > 0 else np.nan}


# ------------------------------------------------------------------ processo estocástico

def monte_carlo(inicial, aporte, mu, sigma, anos, meta_renda, dy, cdi, n=5000, semente=42):
    """Movimento browniano geométrico mensal: Sₜ₊₁ = Sₜ·exp((μ - σ²/2)Δt + σ√Δt·Z) + aporte.
    μ inclui proventos reinvestidos. Retorna percentis por ano e probabilidades."""
    rng = np.random.default_rng(semente)
    dt, meses = 1 / 12, int(anos * 12)
    mu_log = np.log1p(mu)
    S = np.full(n, float(inicial))
    cdi_val, aportado = float(inicial), float(inicial)
    ano_meta = np.full(n, np.nan)
    linhas = []
    for m in range(1, meses + 1):
        z = rng.standard_normal(n)
        S = S * np.exp((mu_log - sigma ** 2 / 2) * dt + sigma * np.sqrt(dt) * z) + aporte
        cdi_val = cdi_val * (1 + cdi) ** dt + aporte
        aportado += aporte
        if meta_renda > 0:
            atingiu = np.isnan(ano_meta) & (S * dy / 12 >= meta_renda)
            ano_meta[atingiu] = m / 12
        if m % 12 == 0:
            p = np.percentile(S, [5, 25, 50, 75, 95])
            linhas.append({"ano": m // 12, "p5": p[0], "p25": p[1], "p50": p[2], "p75": p[3], "p95": p[4],
                           "cdi": cdi_val, "aportado": aportado, "prob_cdi": float((S > cdi_val).mean()),
                           "prob_perda": float((S < aportado).mean())})
    return {"anos": pd.DataFrame(linhas), "prob_meta": float((~np.isnan(ano_meta)).mean()),
            "ano_meta_mediano": float(np.nanmedian(ano_meta)) if (~np.isnan(ano_meta)).any() else None,
            "n": n, "mu": mu, "sigma": sigma}
