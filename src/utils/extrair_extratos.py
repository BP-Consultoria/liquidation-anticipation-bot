import requests
import json
import yaml
import ast
from datetime import datetime
from pathlib import Path

# Configurações de Caminho
PROJECT_PATH = Path(__file__).resolve().parent.parent.parent
YAML_PATH = PROJECT_PATH / "auth" / "paramenters.yml"
BASE_URL = "https://gapp.bancoarbi.com.br"

# DATA_API = "2026-04-17"
# DATA_BR  = "17/04/2026"
_HOJE = datetime.now()
DATA_API = _HOJE.strftime("%Y-%m-%d")   # Formato para envio na URL/Payload
DATA_BR  = _HOJE.strftime("%d/%m/%Y")   # Formato para validar o retorno da API

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
#  - Chamada à API do Banco Arbi
# ──────────────────────────────────────────────


def consultar_saldo(conta):
    return chamar_api_arbi(gerar_idrequisicao(), conta, 3, DATA_API, DATA_API)

def consultar_extrato(conta):
    """Esta é a função que o seu runner.py estava sentindo falta."""
    return chamar_api_arbi(gerar_idrequisicao(), conta, 4, DATA_API, DATA_API)

# ──────────────────────────────────────────────
# ETAPA 3 - Chamada à API do Banco Arbi
# ──────────────────────────────────────────────


def buscar_valor_liquido(dados_api):
    """
    Busca o valor líquido no extrato com a seguinte prioridade:
      1. PIX REMESSA (débito) para GRLIS SECURITIZADORA
    """
    movimentacoes = parsear_movimentacoes(dados_api)

    for mov in movimentacoes:
        h = mov.get("historico", "").upper()
        finalidade = mov.get("finalidade", "").upper()
        tipo = mov.get("tipo", "")

        if tipo == "debito" and "REMESSA" in h and "SECURITIZADORA" in finalidade:
            print(f"    [MATCH] Valor líquido encontrado (REMESSA): R$ {mov['valor']}")
            return mov["valor"]

    return None

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

# ──────────────────────────────────────────────
# ETAPA 4 - Parsear movimentações
# ──────────────────────────────────────────────

def _tipo_e_natureza_movimento(natureza_raw: str) -> tuple[str, str]:
    n = (natureza_raw or "").strip().upper()
    if n in ("C", "CRÉDITO", "CREDITO", "CREDIT"):
        return "credito", "CRÉDITO"
    if n in ("D", "DÉBITO", "DEBITO", "DEBIT"):
        return "debito", "DÉBITO"
    return "desconhecido", n or "(vazio)"

def parsear_movimentacoes(dados_api):
    """Filtra e organiza as movimentações retornadas pela API."""
    if not isinstance(dados_api, list) or not dados_api:
        return []

    if dados_api[0].get("descricaostatus") != "Sucesso":
        return []

    movimentacoes = []
    for item in dados_api:
        try:
            # Arbi retorna o campo 'resultado' como string de um dict Python
            res_raw = item.get("resultado")
            if not res_raw: continue
            
            resultado = ast.literal_eval(res_raw) if isinstance(res_raw, str) else res_raw

            # Validação de data com strip() para evitar erros de espaços invisíveis
            data_mov_api = str(resultado.get("datamovimento", "")).strip()
            
            if data_mov_api != DATA_BR:
                continue

            historico = resultado.get("historico", "")
            if "-" in historico:
                historico = historico.split("-", 1)[1].strip()

            nat_raw = resultado.get("natureza", "")
            tipo_mov, natureza_label = _tipo_e_natureza_movimento(str(nat_raw))

            movimentacoes.append({
                "data_movimento": data_mov_api,
                "documento": resultado.get("nrodocto", ""),
                "historico": historico,
                "finalidade": resultado.get("finalidade", ""),
                "valor": float(resultado.get("valor", 0)),
                "natureza": nat_raw,
                "natureza_normalizada": natureza_label,
                "tipo": tipo_mov,
            })
        except Exception as e:
            movimentacoes.append({"erro_parse": str(e), "dado_original": item})

    return movimentacoes

def parsear_saldo(dados_api):
    if not isinstance(dados_api, list) or not dados_api:
        return None
    try:
        return float(dados_api[0]["resultado"])
    except (ValueError, KeyError, TypeError):
        return None

# ──────────────────────────────────────────────
# ETAPA 5 - Execução Principal
# ──────────────────────────────────────────────

def extrair_extratos():
    print("=" * 60)
    print("INICIANDO EXTRAÇÃO - BANCO ARBI")
    print(f"Data Alvo: {DATA_BR}")
    print("=" * 60)

    try:
        renovar_token()
    except Exception as e:
        print(f"FALHA NA AUTENTICAÇÃO: {e}")
        return

    resultado_geral = {
        "data_execucao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "periodo_consultado": DATA_BR,
        "contas": []
    }

    resumo = {"com_saldo": 0, "com_movimento": 0, "erros": 0}

    for conta, empresa in CONTAS.items():
        print(f" PROCESSANDO: {empresa} ({conta})")
        
        dados_conta = {
            "conta": conta, "empresa": empresa,
            "saldo": None, "movimentacoes": [], "erro": None
        }

        # 1. Saldo
        res_saldo = chamar_api_arbi(gerar_idrequisicao(), conta, 3, DATA_API, DATA_API)
        if isinstance(res_saldo, dict) and "erro" in res_saldo:
            dados_conta["erro"] = res_saldo["mensagem"]
            resumo["erros"] += 1
        else:
            dados_conta["saldo"] = parsear_saldo(res_saldo)
            if dados_conta["saldo"] and dados_conta["saldo"] > 0: resumo["com_saldo"] += 1

        # 2. Extrato
        res_extrato = chamar_api_arbi(gerar_idrequisicao(), conta, 4, DATA_API, DATA_API)
        if isinstance(res_extrato, dict) and "erro" in res_extrato:
            dados_conta["erro"] = res_extrato["mensagem"]
        else:
            movs = parsear_movimentacoes(res_extrato)
            dados_conta["movimentacoes"] = movs
            if movs: resumo["com_movimento"] += 1

        resultado_geral["contas"].append(dados_conta)
        
        # Log rápido
        status_mov = f"{len(dados_conta['movimentacoes'])} movs"
        print(f"   -> Saldo: {dados_conta['saldo']} | Movimentos: {status_mov}")

    resultado_geral["resumo"] = resumo
    return resultado_geral

if __name__ == "__main__":
    resultado = extrair_extratos()
    
    output_path = PROJECT_PATH / "extratos_resultado.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f"\nFinalizado! JSON salvo em: {output_path}")