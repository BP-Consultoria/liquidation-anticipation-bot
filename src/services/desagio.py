import numpy as np
import pandas as pd
import pyodbc
from datetime import date, datetime, timedelta
from typing import Any

from config import settings

REGRAS_ANTECIPACAO = [
    {
        "sacado_contem": ["M. DIAS", "M DIAS"],
        "cedente_contem": ["TEC TRANSPORTES", "IG TRANSPORTES"],
        "dias": 15,
    },
    {
        "sacado_contem": ["PETROLEO", "PETROBRAS"],
        "cedente_contem": ["GAIA"],
        "dias": 15,
    },
    {
        "sacado_contem": ["PETROLEO", "PETROBRAS"],
        "cedente_contem": ["techprime", "Techprime", "TECHPRIME"],
        "dias": 21,
    },
    {
        "sacado_contem": ["BRASKEM"],
        "cedente_contem": ["Usiface", "USIFACE"],
        "dias": 21,
    },
    {
        "sacado_contem": ["NESTLE"],
        "cedente_contem": ["JTD TRANSPORTES"],
        "dias": 21,
    },
    {
        "sacado_contem": ["NESTLE"],
        "cedente_contem": ["L B R TRANSPORTES"],
        "dias": 21,
    },
]


def _normalize(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).split()).upper()


def _regra_aplica(regra: dict, cedente: str, sacado: str) -> bool:
    cedente_n = " ".join(cedente.split()).upper()
    sacado_n = " ".join(sacado.split()).upper()

    cedente_ok = True
    sacado_ok = True

    cedentes_contem = regra.get("cedente_contem")
    if cedentes_contem:
        cedente_ok = any(sub.upper() in cedente_n for sub in cedentes_contem)

    sacados_contem = regra.get("sacado_contem")
    if sacados_contem:
        sacado_ok = any(sub.upper() in sacado_n for sub in sacados_contem)

    return cedente_ok and sacado_ok


def _ajustar_para_dia_util(data: date) -> date:
    """Se a data cair em sábado (5) ou domingo (6), recua para sexta-feira."""
    if data.weekday() == 5:  # sábado → recua 1 dia
        return data - timedelta(days=1)
    if data.weekday() == 6:  # domingo → recua 2 dias
        return data - timedelta(days=2)
    return data


def obter_dias_antecipacao(cedente: str, sacado: str) -> int | None:
    """Prazo mínimo em dias (regra cedente/sacado) usado em ``calcular_desagio``.

    Quando não houver regra cadastrada, retorna ``None``; o chamador usa ``0`` no deságio.
    """
    cedente_n = _normalize(cedente)
    sacado_n = _normalize(sacado)
    for regra in REGRAS_ANTECIPACAO:
        if _regra_aplica(regra, cedente_n, sacado_n):
            return regra["dias"]
    return None


def filtrar_titulos_para_hoje(
    df: pd.DataFrame,
    data_ref: date | None = None,
    col_emissao: str = "EMISSAO",
    col_cedente: str = "CEDENTE",
    col_sacado: str = "SACADO",
) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    dia_atual = data_ref if data_ref is not None else date.today()
    if isinstance(dia_atual, datetime):
        dia_atual = dia_atual.date()
    elif not isinstance(dia_atual, date):
        dia_atual = getattr(dia_atual, "date", lambda: dia_atual)()

    if (
        col_emissao not in df.columns
        or col_cedente not in df.columns
        or col_sacado not in df.columns
    ):
        return pd.DataFrame()

    emissao = pd.to_datetime(df[col_emissao], errors="coerce", dayfirst=True).dt.date
    cedentes = df[col_cedente].astype(str)
    sacados = df[col_sacado].astype(str)

    def to_date(value: Any) -> date | None:
        if pd.isna(value):
            return None
        if isinstance(value, date):
            return value
        if hasattr(value, "date"):
            return value.date()
        return None

    mask: list[bool] = []
    for i in range(len(df)):
        e = to_date(emissao.iloc[i])
        if e is None:
            mask.append(False)
            continue
        ced = cedentes.iloc[i]
        sac = sacados.iloc[i]
        dias = obter_dias_antecipacao(ced, sac)
        if dias is None:
            mask.append(False)
            continue
        data_alvo = _ajustar_para_dia_util(e + timedelta(days=dias))
        mask.append(data_alvo == dia_atual)
    return df.loc[mask].copy().reset_index(drop=True)


def _get_sec_connection():
    """Cria conexão com o banco de securitização."""
    conn_str = (
        f"DRIVER={{{settings.sec_odbc_driver}}};"
        f"SERVER={settings.sec_host},{settings.sec_port};"
        f"DATABASE={settings.sec_name};"
        f"UID={settings.sec_user};"
        f"PWD={settings.sec_password}"
    )
    return pyodbc.connect(conn_str)


def get_dtpgto() -> str:
    """Retorna a data de hoje formatada como dd-mm-aaaa."""
    return date.today().strftime("%d-%m-%Y")


def get_fator(bordero: int) -> float:
    """Retorna a taxa (fator) do borderô a partir da tabela sigbors."""
    conn = _get_sec_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT fator FROM sigbors WHERE bordero = ?", bordero)
        row = cursor.fetchone()
        if row is None:
            raise ValueError(f"Fator não encontrado para o borderô {bordero}.")
        return float(row[0])
    finally:
        conn.close()


def calcular_desagio(df: pd.DataFrame, prazo_minimo: int) -> pd.DataFrame:
    df = df.copy()

    df["bordero"] = df["bordero"].astype(int)
    df["fator"] = df["bordero"].apply(get_fator)
    df["valor_desagio"] = np.nan

    df["vencimento"] = pd.to_datetime(df["vencimento"], dayfirst=True)
    df["dtpgto"] = get_dtpgto()
    df["dtpgto"] = pd.to_datetime(df["dtpgto"], dayfirst=True)
    df["emissao"] = pd.to_datetime(df["emissao"], dayfirst=True)

    # Diferença entre data de pagamento e emissão
    df["diferenca_dias"] = (df["dtpgto"] - df["emissao"]).dt.days

    df = df.dropna(subset=["fator"])
    df["Valor"] = df["Valor"].fillna(0)

    # Calcula data ajustada quando diferença < prazo mínimo
    nova_data = df["emissao"] + pd.to_timedelta(prazo_minimo, unit="D")
    cond = df["diferenca_dias"] < prazo_minimo

    # Dias finais para cálculo
    qt_dias = np.where(
        cond,
        (df["vencimento"] - nova_data).dt.days,
        (df["vencimento"] - df["dtpgto"]).dt.days,
    )

    # Cálculo deságio
    valor = ((df["fator"] / 100) / 30) * qt_dias * df["Valor"]
    valor = valor.clip(lower=0)

    df["valor_desagio"] = valor
    df = df[df["diferenca_dias"] > 0]

    return df
