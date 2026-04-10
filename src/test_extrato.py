import ast
import json
from datetime import datetime
from utils.extrair_extratos import renovar_token, consultar_extrato, ler_yaml

# Renovar token
renovar_token()

# Consultar extrato da IG TRANSPORTES (conta que teve movimentação hoje)
conta = "0003717752"
print(f"Consultando extrato completo da conta {conta}...")
extrato_api = consultar_extrato(conta)

if isinstance(extrato_api, dict) and "erro" in extrato_api:
    print(f"ERRO: {extrato_api}")
    exit()

dia_atual = datetime.now().strftime("%d/%m/%Y")
print(f"Data filtro: {dia_atual}")
print(f"Total de itens retornados pela API: {len(extrato_api)}\n")

# Mostrar TODAS as movimentações do dia sem filtro
for i, item in enumerate(extrato_api):
    try:
        resultado = ast.literal_eval(item["resultado"])
        data_mov = resultado.get("datamovimento", "")
        if data_mov != dia_atual:
            continue

        print(f"--- Movimentação {i} ---")
        print(f"  Data: {data_mov}")
        print(f"  Documento: {resultado.get('nrodocto', '')}")
        print(f"  Historico: {resultado.get('historico', '')}")
        print(f"  Finalidade: {resultado.get('finalidade', '')}")
        print(f"  Valor: {resultado.get('valor', '')}")
        print(f"  Natureza: {resultado.get('natureza', '')} ({'CRÉDITO' if resultado.get('natureza') == 'C' else 'DÉBITO'})")
        print()
    except (ValueError, SyntaxError, KeyError) as e:
        print(f"  Erro ao parsear item {i}: {e}")
