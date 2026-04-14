"""Funções auxiliares para texto/valores usados no RPA WBA."""

from datetime import date, datetime

import pandas as pd


def valor_monetario_br(valor: float) -> str:
    """Formata valor como 1.234,56 (milhar com ponto, decimal com vírgula)."""
    neg = valor < 0
    v = abs(valor)
    inteiro_s, frac = f"{v:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro_s):,}".replace(",", ".")
    corpo = f"{inteiro_fmt},{frac}"
    return f"-{corpo}" if neg else corpo


def valor_monetario_wba_campo_float(valor: float) -> str:
    """Campo numérico do WBA que valida como float: sem separador de milhar, ponto decimal."""
    neg = valor < 0
    v = abs(valor)
    s = f"{v:.2f}"
    return f"-{s}" if neg else s


def mes_abreviatura_pt_br(mes: int) -> str:
    """Retorna abreviação de 3 letras do mês (JAN … DEZ). ``mes``: 1–12."""
    abrev = (
        "JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
        "JUL", "AGO", "SET", "OUT", "NOV", "DEZ",
    )
    if not 1 <= mes <= 12:
        raise ValueError(f"Mês inválido: {mes} (esperado 1–12)")
    return abrev[mes - 1]


def texto_historico_desagio_padrao(data_ref: date | None = None) -> str:
    """Texto do histórico WBA: ``7-DESAGIO ANT MDIAS-{MMM}/{AAAA}`` (mês e ano dinâmicos)."""
    d = data_ref if data_ref is not None else date.today()
    if isinstance(d, datetime):
        d = d.date()
    mes_str = mes_abreviatura_pt_br(d.month)
    return f"7-DESAGIO ANT MDIAS-{mes_str}/{d.year}"


def valor_total_desagio_unico(df: pd.DataFrame) -> float:
    """Lê **um** valor de ``Valor_Total_Desagio`` (o mesmo em todas as linhas do lote; não soma)."""
    if "Valor_Total_Desagio" not in df.columns:
        raise ValueError("Coluna Valor_Total_Desagio ausente.")
    vals = pd.to_numeric(df["Valor_Total_Desagio"], errors="coerce").dropna()
    if vals.empty:
        raise ValueError("Valor_Total_Desagio vazio ou inválido.")
    v = float(vals.iloc[0])
    if v <= 0:
        raise ValueError("Valor_Total_Desagio deve ser > 0.")
    return v


def codigo_cedente_unico(df: pd.DataFrame) -> int:
    """Valida um único ``codigo_cedente`` no lote (um cedente por DataFrame) e retorna o código."""
    if df.empty:
        raise ValueError("DataFrame vazio.")
    if "codigo_cedente" not in df.columns:
        raise ValueError("Coluna 'codigo_cedente' ausente.")
    codigos = df["codigo_cedente"].dropna().unique()
    if len(codigos) != 1:
        raise ValueError(
            "O lote deve ter um único codigo_cedente; agrupe por cedente antes do lançamento."
        )
    c = codigos[0]
    if isinstance(c, (float, int)) and not isinstance(c, bool):
        return int(c)
    return int(str(c).strip())
