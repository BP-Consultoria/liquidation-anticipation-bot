"""Funções auxiliares para texto/valores usados no RPA WBA."""

import re
from datetime import date, datetime
from typing import Any

import pandas as pd


def normalizar_id_titulo_dcto(valor: Any) -> str:
    """Normaliza ``Titulo`` / dcto do WBA para comparar com texto copiado da grid."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, bool):
        return str(int(valor))
    if isinstance(valor, (int, float)):
        f = float(valor)
        if f.is_integer():
            return str(int(f))
        return str(valor).strip()
    s = str(valor).strip()
    try:
        f = float(s.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return s


def normalizar_texto_copiado_grid_dcto(texto: str) -> str:
    """Limpa o que costuma vir do Ctrl+C na célula (tabs, quebras, espaços)."""
    if not texto:
        return ""
    linha = texto.replace("\r", "\n").split("\n")[0]
    parte = linha.split("\t")[0].strip()
    return parte.strip()


def texto_copiado_indica_dcto(copiado_bruto: str, dcto_alvo: str) -> bool:
    """True se o texto do ``Ctrl+C`` corresponde ao documento ``dcto_alvo``.

    Aceita célula só com o código, ou texto onde o código aparece delimitado por não-dígitos
    (ex.: histórico que inclui o número do título).
    """
    alvo = normalizar_id_titulo_dcto(dcto_alvo)
    if not alvo:
        return False
    cop = normalizar_texto_copiado_grid_dcto(copiado_bruto)
    if normalizar_id_titulo_dcto(cop) == alvo:
        return True
    if not cop:
        return False
    return bool(re.search(rf"(?:^|\D){re.escape(alvo)}(?:\D|$)", cop))


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


def calcular_ajuste_dinamico(
    df: pd.DataFrame,
    coluna_identificador_grid: str = "Titulo",
) -> tuple[pd.DataFrame, float, dict[str, Any] | None]:
    """Usa ``Debito_Credito`` do lote (primeira linha). Se **negativo**, abate no maior ``Valor``.

    O ``df`` deve estar na **mesma ordem** do grid da Recompra (ex.: ``preparar_df_para_rpa``).
    ``coluna_identificador_grid``: coluna do ``df`` que corresponde ao **Dcto** copiado no WBA
    (por padrão ``Titulo``; use outra se o ERP exibir outro código na coluna Dcto).
    Retorna o DataFrame atualizado (cópia), valor residual positivo ou ``0.00``, e metadados do
    ajuste ou ``None``.
    """
    if df.empty:
        return df, 0.00, None

    out = df.copy()
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")

    if "Debito_Credito" not in out.columns:
        raise ValueError("Coluna Debito_Credito ausente.")

    dc = pd.to_numeric(out["Debito_Credito"].iloc[0], errors="coerce")
    if pd.isna(dc):
        dc = 0.0
    diferenca = round(float(dc), 2)

    posicao_maior = int(out["Valor"].values.argmax())
    valor_original = float(out.iloc[posicao_maior]["Valor"])

    if diferenca < 0:
        novo_valor = round(valor_original + diferenca, 2)

        if coluna_identificador_grid not in out.columns:
            raise ValueError(
                f"Coluna {coluna_identificador_grid!r} ausente; necessária para localizar o dcto no grid."
            )
        dcto_alvo = normalizar_id_titulo_dcto(out.iloc[posicao_maior][coluna_identificador_grid])

        print("--- Relatório de Ajuste ---")
        print(f"Maior Valor (alvo): {valor_original} na posição (índice df) {posicao_maior}")
        print(f"Dcto ({coluna_identificador_grid}) alvo para busca no grid: {dcto_alvo}")
        print(f"Debito_Credito (diferença): {diferenca}")
        print(f"Cálculo: {valor_original} + ({diferenca}) = {novo_valor}")
        print("---------------------------")

        out.iat[posicao_maior, out.columns.get_loc("Valor")] = novo_valor

        return out, 0.00, {
            "posicao": posicao_maior,
            "valor": novo_valor,
            "dcto_alvo": dcto_alvo,
            "valor_original": valor_original,
        }

    return out, diferenca if diferenca > 0 else 0.00, None


def atualizar_valor_no_df_por_identificador(
    df: pd.DataFrame,
    coluna_identificador: str,
    id_documento: Any,
    novo_valor: float,
) -> pd.DataFrame:
    """Cópia do ``df`` com ``Valor`` atualizado na linha cujo identificador coincide com o do grid.

    Usado quando o dcto e o valor a colar já foram definidos fora de ``calcular_ajuste_dinamico``.
    """
    out = df.copy()
    if coluna_identificador not in out.columns:
        raise ValueError(f"Coluna {coluna_identificador!r} ausente.")
    if "Valor" not in out.columns:
        raise ValueError("Coluna Valor ausente.")
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")
    alvo = normalizar_id_titulo_dcto(id_documento)
    chave = out[coluna_identificador].map(normalizar_id_titulo_dcto)
    mask = chave == alvo
    if not bool(mask.any()):
        raise ValueError(
            f"Nenhuma linha com {coluna_identificador!r} compatível com dcto {alvo!r}."
        )
    novo = round(float(novo_valor), 2)
    out.loc[mask, "Valor"] = novo
    return out
