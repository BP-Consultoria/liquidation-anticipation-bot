import pyodbc
import pandas as pd
from config import settings
from utils.extrair_extratos import (
    CONTAS, renovar_token, consultar_extrato, parsear_movimentacoes,
)

CEDENTE_PARA_CONTA = {nome: conta for conta, nome in CONTAS.items()}

conn_str = (
    f"DRIVER={{{settings.db_odbc_driver}}};"
    f"SERVER={settings.db_host},{settings.db_port};"
    f"DATABASE={settings.db_name};"
    f"UID={settings.db_user};"
    f"PWD={settings.db_password}"
)

print("=" * 60)
print("TESTE DO FLUXO COMPLETO")
print("=" * 60)

# 1. Conectar e buscar antecipações de hoje
print(f"\nConectando em {settings.db_host}/{settings.db_name}...")
conn = pyodbc.connect(conn_str)
print("Conectado!\n")

cursor = conn.cursor()
cursor.execute(
    "SELECT "
    "    a.Bordero, a.Cedente, a.Titulo, a.Valor, "
    "    a.Is_inserted, a.Control_id, a.Is_send, "
    "    a.Titulo_Completo, a.created_at, a.Vencimento, "
    "    a.Valor_Desagio, b.Valor_Liquido_Final "
    "FROM anticipation_db.dbo.anticipation_db AS a "
    "INNER JOIN anticipation_db.dbo.borderos AS b "
    "    ON a.Bordero = b.Bordero "
    "WHERE "
    "    a.Is_inserted = 1 "
    "    AND a.Control_id = 4 "
    "    AND a.created_at >= CAST(GETDATE() AS DATE) "
    "    AND a.created_at < DATEADD(DAY, 1, CAST(GETDATE() AS DATE))"
)
colunas = [col[0] for col in cursor.description]
linhas = cursor.fetchall()
antecipacoes = [dict(zip(colunas, row)) for row in linhas]
conn.close()

print(f"[DB] Antecipações encontradas hoje: {len(antecipacoes)}")

if not antecipacoes:
    print("Nenhuma antecipação para processar.")
    exit()

# 2. Agrupar por cedente
borderos_por_cedente = {}
for row in antecipacoes:
    cedente = row["Cedente"]
    borderos_por_cedente.setdefault(cedente, []).append(row["Bordero"])

print(f"[DB] Cedentes encontrados: {list(borderos_por_cedente.keys())}")

# 3. Renovar token e consultar extrato
renovar_token()

for cedente, borderos in borderos_por_cedente.items():
    conta = CEDENTE_PARA_CONTA.get(cedente)
    if not conta:
        print(f"\n[SKIP] Cedente '{cedente}' não tem conta mapeada.")
        continue

    print(f"\n[API] Consultando extrato de {cedente} (conta {conta})...")

    extrato_api = consultar_extrato(conta)
    if isinstance(extrato_api, dict) and "erro" in extrato_api:
        print(f"  ERRO: {extrato_api['erro']}")
        continue

    movimentacoes = parsear_movimentacoes(extrato_api)
    print(f"  Movimentações filtradas: {len(movimentacoes)}")

    if movimentacoes:
        valor = movimentacoes[0]["valor"]
        borderos_unicos = set(borderos)
        print(f"  Valor da movimentação: R$ {valor:,.2f}")
        print(f"  Borderôs que seriam atualizados: {borderos_unicos}")
    else:
        print(f"  Nenhuma movimentação de BANCO ITAU encontrada hoje.")

print("\n" + "=" * 60)
print("TESTE FINALIZADO (nenhum UPDATE foi executado)")
print("=" * 60)
