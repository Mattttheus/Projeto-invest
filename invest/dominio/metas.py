"""Meta de renda por ativo: quanto cada ação/FII deve pagar por mês e em quanto tempo.

Padrão para todos (config.json: meta_mensal_por_ativo, prazo_meta_meses); o usuário pode personalizar
qualquer ativo (dados/metas.csv). Campo vazio = volta ao padrão.
"""
import pandas as pd

from .avisos import AVISO, Aviso

PRAZO_PADRAO = 60                   # meses, quando config.json não informa
COLUNAS = ["ticker", "meta_mensal", "prazo_meses"]


def padroes(cfg):
    return float(cfg.get("meta_mensal_por_ativo", 0)), int(cfg.get("prazo_meta_meses", PRAZO_PADRAO))


def vazias():
    return pd.DataFrame(columns=COLUNAS)


def resolver(ativos, metas, cfg):
    """Meta e prazo efetivos de cada ativo do cadastro (personalizados ou padrão).
    Retorna DataFrame indexado por ticker: meta_mensal, prazo_meses, personalizada."""
    meta_p, prazo_p = padroes(cfg)
    m = metas.drop_duplicates("ticker", keep="last").set_index("ticker").reindex(ativos["ticker"])
    out = pd.DataFrame(index=ativos["ticker"])
    out["meta_mensal"] = m["meta_mensal"].astype(float).fillna(meta_p)
    out["prazo_meses"] = m["prazo_meses"].astype(float).fillna(prazo_p).astype(int)
    out["personalizada"] = m["meta_mensal"].notna() | m["prazo_meses"].notna()
    return out


def validar(metas, ativos):
    out = []
    for t in sorted(set(metas["ticker"]) - set(ativos["ticker"])):
        out.append(Aviso(AVISO, t, "Meta em metas.csv para ticker que não está em ativos.csv — ignorada."))
    for _, r in metas.iterrows():
        if pd.notna(r["meta_mensal"]) and r["meta_mensal"] < 0:
            out.append(Aviso(AVISO, r["ticker"], "Meta mensal negativa em metas.csv — use um valor positivo."))
        if pd.notna(r["prazo_meses"]) and r["prazo_meses"] <= 0:
            out.append(Aviso(AVISO, r["ticker"], "Prazo da meta deve ser de pelo menos 1 mês (metas.csv)."))
    return out


def entrada(ticker, meta, prazo):
    """Normaliza o que foi digitado na aba Meta. Vazio/fórmula = padrão (None). ValueError se inválido."""
    meta = _numero_ou_nada(meta)
    prazo = _numero_ou_nada(prazo)
    if meta is not None and meta < 0:
        raise ValueError(f"{ticker}: meta mensal negativa")
    if prazo is not None and prazo < 1:
        raise ValueError(f"{ticker}: prazo deve ser de pelo menos 1 mês")
    return {"ticker": ticker, "meta_mensal": meta, "prazo_meses": None if prazo is None else int(round(prazo))}


def _numero_ou_nada(v):
    if v is None or (isinstance(v, str) and (not v.strip() or v.startswith("="))):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace("R$", "").replace(".", "").replace(",", ".").strip())
