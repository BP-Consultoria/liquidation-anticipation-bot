import pandas as pd

from services.database import db_service
from services.desagio import calcular_desagio, obter_dias_antecipacao
from utils.extrair_extratos import (
    CONTAS, renovar_token, consultar_extrato, buscar_valor_liquido,
)

from rpa.Wba import WBA

# Mapa de código do cedente no sistema
CODIGO_CEDENTE = {
    "IG TRANSPORTES": 16634,
    "TEC TRANSPORTES": 5088,
    "GAIA EMPREENDIMENTOS": 21498,
}


def buscar_codigo_cedente(cedente_db: str) -> int | None:
    """Retorna o código do cedente a partir do nome do banco."""
    cedente_upper = cedente_db.upper().strip()
    for prefixo, codigo in CODIGO_CEDENTE.items():
        if cedente_upper.startswith(prefixo):
            return codigo
    return None


def preparar_df_para_rpa(df: pd.DataFrame) -> pd.DataFrame:
    """Garante ``Valor_Liquido`` e ordena colunas para o RPA."""
    out = df.copy()
    if "Valor_Liquido" not in out.columns:
        if "Valor_Liquido_Final" in out.columns:
            out["Valor_Liquido"] = out["Valor_Liquido_Final"]
        else:
            raise ValueError("DataFrame sem Valor_Liquido nem Valor_Liquido_Final.")
    out["Vencimento"] = pd.to_datetime(out["Vencimento"], errors="coerce", dayfirst=True)
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")
    out = out.sort_values(
        by=["Vencimento", "Valor", "Titulo"],
        ascending=[True, True, False],
    ).reset_index(drop=True)
    return out


def buscar_conta_por_cedente(cedente_db: str) -> str | None:
    """Mapeia nome do cedente do banco → número da conta no Arbi.
    Compara pelo início do nome para tolerar diferenças como EIRELI/LTDA."""
    cedente_upper = cedente_db.upper().strip()
    for conta, nome_arbi in CONTAS.items():
        # Compara as duas primeiras palavras (ex: "TEC TRANSPORTES")
        palavras_arbi = nome_arbi.upper().split()[:2]
        palavras_db = cedente_upper.split()[:2]
        if palavras_arbi == palavras_db:
            return conta
    return None


def obter_valor_liquido_arbi_todos_cedentes(
    borderos_por_cedente: dict[str, list],
) -> dict[str, tuple[float, list]] | None:
    """Para cada cedente, obtém o valor líquido (TED REMESSA) no Arbi.

    Se **qualquer** cedente falhar (sem conta, erro de API ou sem TED), retorna ``None``
    e o fluxo não deve atualizar o banco nem seguir para o RPA.
    """
    resultado: dict[str, tuple[float, list]] = {}
    for cedente, borderos in borderos_por_cedente.items():
        conta = buscar_conta_por_cedente(cedente)
        if not conta:
            print(
                f"\n[FLOW] Cedente '{cedente}' sem conta Arbi no mapa. "
                "Valor líquido obrigatório via API — fluxo abortado."
            )
            return None

        print(f"\n[FLOW] Consultando extrato Arbi de {cedente} (conta {conta})...")

        extrato_api = consultar_extrato(conta)
        if isinstance(extrato_api, dict) and "erro" in extrato_api:
            print(
                f"  ERRO na API Arbi: {extrato_api['erro']}. "
                "Valor líquido não obtido — fluxo abortado."
            )
            return None

        valor = buscar_valor_liquido(extrato_api)
        if valor is None:
            print(
                f"  Nenhuma TED REMESSA no extrato Arbi para {cedente}. "
                "Valor líquido obrigatório — fluxo abortado."
            )
            return None

        print(f"  TED REMESSA (valor líquido): R$ {valor:,.2f}")
        resultado[cedente] = (float(valor), borderos)

    return resultado


def run():
    print("=" * 60)
    print("INICIANDO FLUXO DE ANTECIPAÇÃO")
    print("=" * 60)

    # 1. Conectar no banco
    db_service.conectar()

    try:
        # 2. Buscar antecipações do dia
        antecipacoes = db_service.buscar_antecipacoes_hoje()
        print(f"[FLOW] Antecipações encontradas hoje: {len(antecipacoes)}")

        if not antecipacoes:
            print("[FLOW] Nenhuma antecipação para processar.")
            return

        # 3. Renovar token da API Arbi
        renovar_token()

        # 4. Agrupar borderôs por cedente (nome do banco)
        borderos_por_cedente = {}
        for row in antecipacoes:
            cedente = row["Cedente"]
            borderos_por_cedente.setdefault(cedente, []).append(row["Bordero"])

        # 5. Valor líquido obrigatório via Arbi (TED REMESSA) para todos os cedentes; só então atualiza o banco
        valores_arbi = obter_valor_liquido_arbi_todos_cedentes(borderos_por_cedente)
        if valores_arbi is None:
            print("\n[FLOW] Valor líquido Arbi incompleto ou ausente. RPA não será iniciado.")
            return

        cedentes_atualizados = []
        for cedente, (valor, borderos) in valores_arbi.items():
            for bordero in set(borderos):
                db_service.atualizar_valor_liquido(bordero, valor)
            cedentes_atualizados.append(cedente)

        # 6. Buscar DF para o RPA (borderôs já atualizados com líquido do Arbi)

        print("\n[RPA] Buscando dados para liquidação...")
        df = db_service.buscar_dados_para_rpa(cedentes_atualizados)
        df["codigo_cedente"] = df["Cedente"].apply(buscar_codigo_cedente)

        # 7. Recalcular deságio por cedente (total usado no WBA; não persiste valor_desagio no DF)
        print("\n[DESAGIO] Recalculando deságio dos títulos...")
        _COLS_CALC_INTERNO = ("valor_desagio", "fator", "diferenca_dias", "dtpgto")
        dfs_final = []
        lotes_wba: list[tuple[pd.DataFrame, float]] = []

        for cedente, group in df.groupby("Cedente"):
            sacado = str(group["Sacado"].iloc[0])
            dias_regra = obter_dias_antecipacao(str(cedente), sacado)
            prazo_minimo = dias_regra if dias_regra is not None else 0

            df_calc = group.rename(columns={
                "Bordero": "bordero",
                "Titulo": "titulo",
                "Emissao": "emissao",
                "Vencimento": "vencimento",
            })

            df_desagio = calcular_desagio(df_calc, prazo_minimo=prazo_minimo)

            if df_desagio.empty:
                print(f"  {cedente}: nenhum título elegível para deságio")
                continue

            df_desagio = df_desagio.rename(columns={
                "bordero": "Bordero",
                "titulo": "Titulo",
                "emissao": "Emissao",
                "vencimento": "Vencimento",
            })

            total_desagio = float(df_desagio["valor_desagio"].sum())
            print(f"  {cedente}: {len(df_desagio)} títulos | Deságio total: R$ {total_desagio:,.2f}")

            df_desagio = df_desagio.drop(
                columns=[c for c in _COLS_CALC_INTERNO if c in df_desagio.columns],
                errors="ignore",
            )
            df_prep = preparar_df_para_rpa(df_desagio.copy())
            dfs_final.append(df_prep)
            lotes_wba.append((df_prep, total_desagio))

        if not dfs_final:
            print("\n[FLOW] Nenhum título elegível para deságio. RPA não será iniciado.")
            return

        df_concat = pd.concat(dfs_final, ignore_index=True)

        cols_show = [
            "Bordero",
            "Cedente",
            "codigo_cedente",
            "Titulo",
            "Vencimento",
            "Valor",
            "Valor_Liquido",
            "Valor_Liquido_Final",
        ]
        cols_show = [c for c in cols_show if c in df_concat.columns]
        print(f"\n[RPA] {len(df_concat)} títulos prontos para liquidar:")
        print(df_concat[cols_show].to_string(index=False))

        # RPA WBA: um lançamento por cedente (total_desagio calculado acima)
        wba = WBA()
        try:
            wba.login()
            for df_lote, total_desagio in lotes_wba:
                wba.lancar_desagio_contas_lancamentos(df_lote, total_desagio=total_desagio)
        finally:
            wba.close_wba_application()

        print("\n" + "=" * 60)
        print("FLUXO FINALIZADO")
        print("=" * 60)

    finally:
        db_service.desconectar()
