import requests
import json
import yaml
import ast
from datetime import datetime
from pathlib import Path

PROJECT_PATH = Path(__file__).resolve().parent.parent.parent  # raiz do projeto
YAML_PATH = PROJECT_PATH / "auth" / "paramenters.yml"
BASE_URL = "https://gapp.bancoarbi.com.br"

CONTAS = {
    '0003717752': "IG TRANSPORTES LTDA",
    '0003715768': "TEC TRANSPORTES LTDA",
    '0003730040': "GAIA EMPREENDIMENTOS CONSTRUCOES E",
}

# ──────────────────────────────────────────────
# YAML - leitura e escrita
# ──────────────────────────────────────────────

def ler_yaml():
    with open(YAML_PATH, 'r') as f:
        return yaml.load(f, Loader=yaml.FullLoader)


def gravar_yaml(config):
    with open(YAML_PATH, 'w') as f:
        yaml.dump(config, f)


# ──────────────────────────────────────────────
# ETAPA 1 - Autenticação OAuth2
# ──────────────────────────────────────────────

def obter_grant_code(client_id):
    response = requests.post(
        f"{BASE_URL}/oauth/grant-code",
        headers={"Content-Type": "application/json"},
        data=json.dumps({
            "client_id": client_id,
            "redirect_uri": "http://localhost/"
        })
    )
    if response.status_code == 201:
        redirect_uri = response.json()["redirect_uri"]
        return redirect_uri.split("=")[1]
    raise Exception(f"Erro ao obter grant code: {response.status_code} - {response.text}")


def obter_access_token(authorization, client_id):
    code = obter_grant_code(client_id)
    response = requests.post(
        f"{BASE_URL}/oauth/access-token",
        headers={
            "Content-Type": "application/json",
            "Authorization": authorization
        },
        data=json.dumps({
            "grant_type": "authorization_code",
            "code": code
        })
    )
    if response.status_code == 201:
        return response.json()["access_token"]
    raise Exception(f"Erro ao obter access token: {response.status_code} - {response.text}")


def renovar_token():
    config = ler_yaml()
    arbi = config["paramenters_arbi"]
    token = obter_access_token(arbi["authorization"], arbi["client_id"])
    config["paramenters_arbi"]["token"] = token
    gravar_yaml(config)
    print(f"[AUTH] Token renovado com sucesso")
    return token


# ──────────────────────────────────────────────
# ETAPA 2 - ID de Requisição
# ──────────────────────────────────────────────

def gerar_idrequisicao():
    config = ler_yaml()
    novo_id = config["paramenters_arbi"]["idrequisicao"] + 1
    config["paramenters_arbi"]["idrequisicao"] = novo_id
    gravar_yaml(config)
    return novo_id


# ──────────────────────────────────────────────
# ETAPA 3 - Chamada à API do Banco Arbi
# ──────────────────────────────────────────────

def chamar_api_arbi(idrequisicao, conta, idtransacao, datainicial, datafinal):
    config = ler_yaml()
    arbi = config["paramenters_arbi"]

    payload = {
        "contacorrente": {
            "inscricaoparceiro": "26845323000170",
            "tokenusuario": "RtOResyNtOResepOSpEcIaNTIonSFACk",
            "idrequisicao": str(idrequisicao),
            "idmodulo": "1",
            "idtransacao": str(idtransacao),
            "bancoorigem": "213",
            "agenciaorigem": "00019",
            "contaorigem": str(conta),
            "tipocontadebitada": "CC",
            "bancodestino": "",
            "agenciadestino": "",
            "contadestino": "",
            "tipocontacreditada": "",
            "cnpjcpfclicred": "",
            "nomeclicred": "",
            "tipopessoaclicred": "",
            "finalidade": "",
            "historico": "",
            "dataagendamento": str(datafinal),
            "valor": "0",
            "datainicial": str(datainicial),
            "datafinal": str(datafinal),
            "periodoemdias": "",
            "canalentrada": "E"
        }
    }

    headers = {
        "Content-Type": "application/json",
        "client_id": arbi["client_id"],
        "access_token": arbi["token"]
    }

    response = requests.post(
        f"{BASE_URL}/contacorrente/v2/contacorrente/",
        headers=headers,
        data=json.dumps(payload)
    )

    if response.status_code in (200, 201, 204):
        return response.json()
    return {"erro": response.status_code, "mensagem": response.text}


def consultar_extrato(conta):
    dia_atual = datetime.now().strftime("%Y-%m-%d")
    return chamar_api_arbi(gerar_idrequisicao(), conta, 4, dia_atual, dia_atual)


def consultar_saldo(conta):
    dia_atual = datetime.now().strftime("%Y-%m-%d")
    return chamar_api_arbi(gerar_idrequisicao(), conta, 3, dia_atual, dia_atual)


# ──────────────────────────────────────────────
# ETAPA 4 - Parsear movimentações da API
# ──────────────────────────────────────────────

def parsear_movimentacoes(dados_api):
    """Parseia todas as movimentações do dia atual do extrato."""
    if not isinstance(dados_api, list):
        return []

    if not dados_api or dados_api[0].get("descricaostatus") != "Sucesso":
        return []

    dia_atual = datetime.now().strftime("%d/%m/%Y")

    movimentacoes = []
    for item in dados_api:
        try:
            resultado = ast.literal_eval(item["resultado"])

            if resultado.get("datamovimento", "") != dia_atual:
                continue

            historico = resultado.get("historico", "")
            if "-" in historico:
                historico = historico.split("-", 1)[1].strip()

            movimentacoes.append({
                "data_movimento": resultado.get("datamovimento", ""),
                "documento": resultado.get("nrodocto", ""),
                "historico": historico,
                "finalidade": resultado.get("finalidade", ""),
                "valor": float(resultado.get("valor", 0)),
                "natureza": resultado.get("natureza", ""),
                "tipo": "credito" if resultado.get("natureza") == "C" else "debito"
            })
        except (ValueError, SyntaxError, KeyError) as e:
            movimentacoes.append({"erro_parse": str(e), "dado_original": item})

    return movimentacoes


def buscar_valor_liquido(dados_api):
    """Busca o valor do débito TED - REMESSA no extrato do dia (Valor_Liquido_Final)."""
    movimentacoes = parsear_movimentacoes(dados_api)
    for mov in movimentacoes:
        if mov.get("tipo") == "debito" and "TED" in mov.get("historico", "").upper() and "REMESSA" in mov.get("historico", "").upper():
            return mov["valor"]
    return None


def parsear_saldo(dados_api):
    if not isinstance(dados_api, list) or not dados_api:
        return None
    try:
        return float(dados_api[0]["resultado"])
    except (ValueError, KeyError, TypeError):
        return None


# ──────────────────────────────────────────────
# ETAPA 5 - Extração completa
# ──────────────────────────────────────────────

def extrair_extratos():
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO DE EXTRATOS - BANCO ARBI")
    print(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"Total de contas: {len(CONTAS)}")
    print("=" * 60)

    renovar_token()

    resultado_geral = {
        "data_execucao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "periodo": {
            "inicio": datetime.now().strftime("%d/%m/%Y"),
            "fim": datetime.now().strftime("%d/%m/%Y")
        },
        "total_contas": len(CONTAS),
        "resumo": {
            "com_saldo": 0,
            "sem_saldo": 0,
            "com_movimento": 0,
            "sem_movimento": 0,
            "erros": 0
        },
        "contas": []
    }

    for i, (conta, empresa) in enumerate(CONTAS.items(), 1):
        print(f"[{i}/{len(CONTAS)}] Processando: {empresa} (conta {conta})")

        dados_conta = {
            "conta": conta,
            "empresa": empresa,
            "agencia": "0001-9",
            "banco": "213 - Banco Arbi",
            "saldo": None,
            "tem_saldo": False,
            "movimentacoes": [],
            "tem_movimento": False,
            "erro": None
        }

        # Consultar saldo
        saldo_api = consultar_saldo(conta)
        if isinstance(saldo_api, dict) and "erro" in saldo_api:
            dados_conta["erro"] = f"Erro saldo: {saldo_api['erro']} - {saldo_api['mensagem']}"
            resultado_geral["resumo"]["erros"] += 1
            resultado_geral["contas"].append(dados_conta)
            print(f"  ERRO ao consultar saldo: {saldo_api['erro']}")
            continue

        saldo_valor = parsear_saldo(saldo_api)
        dados_conta["saldo"] = saldo_valor
        dados_conta["tem_saldo"] = saldo_valor is not None and saldo_valor > 0

        if dados_conta["tem_saldo"]:
            resultado_geral["resumo"]["com_saldo"] += 1
        else:
            resultado_geral["resumo"]["sem_saldo"] += 1

        # Consultar extrato
        extrato_api = consultar_extrato(conta)
        if isinstance(extrato_api, dict) and "erro" in extrato_api:
            dados_conta["erro"] = f"Erro extrato: {extrato_api['erro']} - {extrato_api['mensagem']}"
            resultado_geral["resumo"]["erros"] += 1
            resultado_geral["contas"].append(dados_conta)
            print(f"  ERRO ao consultar extrato: {extrato_api['erro']}")
            continue

        movimentacoes = parsear_movimentacoes(extrato_api)
        dados_conta["movimentacoes"] = movimentacoes
        dados_conta["tem_movimento"] = len(movimentacoes) > 0

        if dados_conta["tem_movimento"]:
            resultado_geral["resumo"]["com_movimento"] += 1
        else:
            resultado_geral["resumo"]["sem_movimento"] += 1

        saldo_str = f"R$ {saldo_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if saldo_valor is not None else "N/A"
        print(f"  Saldo: {saldo_str} | Movimentações: {len(movimentacoes)}")

        resultado_geral["contas"].append(dados_conta)

    print("\n" + "=" * 60)
    print("EXTRAÇÃO FINALIZADA")
    print(f"Com saldo: {resultado_geral['resumo']['com_saldo']}")
    print(f"Sem saldo: {resultado_geral['resumo']['sem_saldo']}")
    print(f"Com movimento: {resultado_geral['resumo']['com_movimento']}")
    print(f"Sem movimento: {resultado_geral['resumo']['sem_movimento']}")
    print(f"Erros: {resultado_geral['resumo']['erros']}")
    print("=" * 60)

    return resultado_geral


# ──────────────────────────────────────────────
# Execução
# ──────────────────────────────────────────────

if __name__ == "__main__":
    resultado = extrair_extratos()

    output_path = PROJECT_PATH / "extratos_resultado.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"\nJSON salvo em: {output_path}")
