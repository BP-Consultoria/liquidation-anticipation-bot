import pandas as pd

from services.database import db_service
from services.desagio import (
    calcular_debito_credito_lote,
    calcular_desagio,
    obter_dias_antecipacao,
)
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
    """Garante ``Valor_Liquido`` e ordena o lote no mesmo espírito do grid do WBA.

    Ordem de processamento no RPA (recompra, etc.): **vencimento** mais cedo primeiro,
    depois **valor** menor primeiro, depois **título** do menor para o maior (numérico se
    ``Titulo`` for número).
    """
    out = df.copy()
    if "Valor_Liquido" not in out.columns:
        if "Valor_Liquido_Final" in out.columns:
            out["Valor_Liquido"] = out["Valor_Liquido_Final"]
        else:
            raise ValueError("DataFrame sem Valor_Liquido nem Valor_Liquido_Final.")
    out["Vencimento"] = pd.to_datetime(out["Vencimento"], errors="coerce", dayfirst=True)
    out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")
    # Título como número quando possível (evita ordem lexicográfica 10 antes de 9)
    titulo_num = pd.to_numeric(out["Titulo"], errors="coerce")
    out["_titulo_ordem"] = titulo_num
    out = out.sort_values(
        by=["Vencimento", "Valor", "_titulo_ordem", "Titulo"],
        ascending=[True, True, True, True],
        na_position="last",
    ).drop(columns=["_titulo_ordem"])
    out = out.reset_index(drop=True)
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


def _recalcular_debito_credito_e_persistir(g: pd.DataFrame) -> pd.DataFrame:
    """Recalcula ``Debito_Credito`` por borderô e grava ``Valor_Total_Desagio`` + ``Debito_Credito`` no banco."""
    partes: list[pd.DataFrame] = []
    for bordero, gb in g.groupby("Bordero", sort=False):
        gb = gb.copy()
        vlf = float(pd.to_numeric(gb["Valor_Liquido_Final"].iloc[0], errors="coerce") or 0.0)
        soma_val = float(pd.to_numeric(gb["Valor"], errors="coerce").sum())
        vtd = float(pd.to_numeric(gb["Valor_Total_Desagio"].iloc[0], errors="coerce") or 0.0)
        dc = calcular_debito_credito_lote(vlf, soma_val, vtd)
        gb["Debito_Credito"] = dc
        print(f"    Borderô {int(bordero)}: Debito_Credito recalculado = {dc}")
        try:
            db_service.atualizar_desagio_e_debito_credito(int(bordero), vtd, dc)
        except Exception as exc:
            print(f"    [DB] Aviso: não atualizou borderô {bordero} ({exc})")
        partes.append(gb)
    return pd.concat(partes, ignore_index=True)


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

        # 7. Lotes por cedente — recalcula deságio (``services.desagio``); se falhar, usa Valor_Total_Desagio do banco
        print("\n[RPA] Agrupando por cedente (deságio recalculado ou valor do banco)...")
        dfs_final = []

        for cedente, group in df.groupby("Cedente"):
            sacado = str(group["Sacado"].iloc[0]) if "Sacado" in group.columns else ""
            dias_regra = obter_dias_antecipacao(str(cedente), sacado)
            prazo_minimo = dias_regra if dias_regra is not None else 0

            df_calc = group.rename(
                columns={
                    "Bordero": "bordero",
                    "Titulo": "titulo",
                    "Emissao": "emissao",
                    "Vencimento": "vencimento",
                }
            )

            try:
                df_desagio = calcular_desagio(
                    df_calc,
                    prazo_minimo,
                    coluna_valor_base="Valor_Liquido_Final",
                )
            except Exception as exc:
                print(f"  {cedente}: erro no cálculo de deságio ({exc}); usando Valor_Total_Desagio do banco.")
                df_desagio = pd.DataFrame()

            if df_desagio.empty:
                vals = pd.to_numeric(group["Valor_Total_Desagio"], errors="coerce").dropna()
                um_valor = float(vals.iloc[0]) if len(vals) else 0.0
                if um_valor <= 0:
                    print(f"  {cedente}: sem deságio calculado e Valor_Total_Desagio ausente/zero — pulando.")
                    continue
                print(
                    f"  {cedente}: {len(group)} título(s) | Valor_Total_Desagio (banco): R$ {um_valor:,.2f}"
                )
                g = group.copy()
            else:
                total_desagio = float(df_desagio["valor_desagio"].sum())
                df_keys = df_desagio.rename(
                    columns={
                        "bordero": "Bordero",
                        "titulo": "Titulo",
                    }
                )
                chaves = df_keys[["Bordero", "Titulo"]].drop_duplicates()
                g = group.merge(chaves, on=["Bordero", "Titulo"], how="inner")
                if g.empty:
                    print(f"  {cedente}: deságio calculado não casou com títulos do banco — pulando.")
                    continue
                g = g.drop(columns=["Valor_Total_Desagio"], errors="ignore")
                g["Valor_Total_Desagio"] = total_desagio
                print(
                    f"  {cedente}: {len(g)} título(s) | Deságio recalculado (soma p/ WBA): R$ {total_desagio:,.2f}"
                )

            g = _recalcular_debito_credito_e_persistir(g)
            df_prep = preparar_df_para_rpa(g)
            dfs_final.append(df_prep)

        if not dfs_final:
            print("\n[FLOW] Nenhum cedente com Valor_Total_Desagio válido. RPA não será iniciado.")
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
            "Valor_Total_Desagio",
            "Debito_Credito",
        ]
        cols_show = [c for c in cols_show if c in df_concat.columns]
        print(f"\n[RPA] {len(df_concat)} títulos prontos para liquidar:")
        print(df_concat[cols_show].to_string(index=False))

        # RPA WBA: um lançamento por cedente (valor único lido de Valor_Total_Desagio no próprio df)
        wba = WBA()
        try:
            wba.login()
            for df_lote in dfs_final:
                wba.lancar_desagio_contas_lancamentos(df_lote)
                wba.recompra_carteira_propria(df_lote)
                wba.inserir_desagio_apos_recompra()
                df_lote = wba.aplicar_ajuste_debito_credito_recompra(df_lote)
        finally:
            wba.close_wba_application()

        print("\n" + "=" * 60)
        print("FLUXO FINALIZADO")
        print("=" * 60)

    finally:
        db_service.desconectar()
