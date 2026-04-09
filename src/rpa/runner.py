from services.database import db_service
from utils.extrair_extratos import (
    CONTAS, renovar_token, consultar_extrato, parsear_movimentacoes,
)

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

        # 5. Para cada cedente, mapear para conta Arbi, consultar extrato e atualizar
        atualizou = False
        cedentes_atualizados = []
        for cedente, borderos in borderos_por_cedente.items():
            conta = buscar_conta_por_cedente(cedente)
            if not conta:
                print(f"\n[FLOW] Cedente '{cedente}' não encontrado no mapa de contas. Pulando.")
                continue

            print(f"\n[FLOW] Consultando extrato de {cedente} (conta {conta})...")

            extrato_api = consultar_extrato(conta)
            if isinstance(extrato_api, dict) and "erro" in extrato_api:
                print(f"  ERRO ao consultar extrato: {extrato_api['erro']}")
                continue

            movimentacoes = parsear_movimentacoes(extrato_api)

            if not movimentacoes:
                print(f"  Nenhuma movimentação encontrada para {cedente}")
                continue

            valor = movimentacoes[0]["valor"]
            print(f"  Movimentação encontrada: R$ {valor:,.2f}")

            borderos_unicos = set(borderos)
            for bordero in borderos_unicos:
                db_service.atualizar_valor_liquido(bordero, valor)
                atualizou = True
            cedentes_atualizados.append(cedente)

        # 6. Se atualizou, buscar DF para o RPA liquidar
        if not atualizou:
            print("\n[FLOW] Nenhum borderô atualizado. RPA não será iniciado.")
            return

        print("\n[RPA] Buscando dados para liquidação...")
        df = db_service.buscar_dados_para_rpa(cedentes_atualizados)
        df["codigo_cedente"] = df["Cedente"].apply(buscar_codigo_cedente)
        print(f"[RPA] {len(df)} títulos para liquidar:")
        print(df[["Bordero", "Cedente", "codigo_cedente", "Titulo", "Vencimento", "Valor_Liquido_Final", "Valor_Total_Desagio"]].to_string(index=False))

        # TODO: iniciar RPA com o df
        # rpa_liquidar(df)

        print("\n" + "=" * 60)
        print("FLUXO FINALIZADO")
        print("=" * 60)

    finally:
        db_service.desconectar()
