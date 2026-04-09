from services.database import db_service
from utils.extrair_extratos import renovar_token, consultar_extrato, parsear_movimentacoes


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

        # 4. Agrupar borderôs por cedente (conta)
        borderos_por_conta = {}
        for row in antecipacoes:
            conta = row["Cedente"]
            borderos_por_conta.setdefault(conta, []).append(row["Bordero"])

        # 5. Para cada conta, consultar extrato e atualizar Valor_Liquido_Final
        atualizou = False
        for conta, borderos in borderos_por_conta.items():
            print(f"\n[FLOW] Consultando extrato da conta {conta}...")

            extrato_api = consultar_extrato(conta)
            if isinstance(extrato_api, dict) and "erro" in extrato_api:
                print(f"  ERRO ao consultar extrato da conta {conta}: {extrato_api['erro']}")
                continue

            movimentacoes = parsear_movimentacoes(extrato_api)

            if not movimentacoes:
                print(f"  Nenhuma movimentação encontrada para conta {conta}")
                continue

            valor = movimentacoes[0]["valor"]
            print(f"  Movimentação encontrada: R$ {valor:,.2f}")

            for bordero in borderos:
                db_service.atualizar_valor_liquido(bordero, valor)
                atualizou = True

        # 6. Se atualizou, buscar DF para o RPA liquidar
        if not atualizou:
            print("\n[FLOW] Nenhum borderô atualizado. RPA não será iniciado.")
            return

        print("\n[RPA] Buscando dados para liquidação...")
        df = db_service.buscar_dados_para_rpa()
        print(f"[RPA] {len(df)} títulos para liquidar:")
        print(df[["Bordero", "Nome_Cedente", "Titulo", "Vencimento", "Valor_Liquido_Final", "Valor_Total_Desagio"]].to_string(index=False))

        # TODO: iniciar RPA com o df
        # rpa_liquidar(df)

        print("\n" + "=" * 60)
        print("FLUXO FINALIZADO")
        print("=" * 60)

    finally:
        db_service.desconectar()
