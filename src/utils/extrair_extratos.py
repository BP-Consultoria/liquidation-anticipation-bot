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


def buscar_conta_por_cedente(cedente_db: str) -> str | None:
    """Mapeia nome do cedente do banco → número da conta no Arbi (duas primeiras palavras)."""
    cedente_upper = cedente_db.upper().strip()
    for conta, nome_arbi in CONTAS.items():
        palavras_arbi = nome_arbi.upper().split()[:2]
        palavras_db = cedente_upper.split()[:2]
        if palavras_arbi == palavras_db:
            return conta
    return None


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

# Faixa: [valor_liquido - tolerância, valor_liquido]; entre os que passam, fica o maior.
TOLERANCIA_VALOR_LIQUIDO_REMESSA = 500.0


def _classificar_pix_ted_remesa_grlis(mov: dict) -> str | None:
    """``'pix'``, ``'ted'`` ou ``None``."""
    h = mov.get("historico", "").upper()
    finalidade = mov.get("finalidade", "").upper()
    tipo = mov.get("tipo", "")

    if (
        tipo != "debito"
        or "REMESSA" not in h
        or "GRLIS SECURITIZADORA" not in finalidade
    ):
        return None
    if "PIX" in h:
        return "pix"
    if "TED" in h:
        return "ted"
    return None


def buscar_valor_liquido(
    dados_api,
    valor_liquido: float,
    tolerancia: float = TOLERANCIA_VALOR_LIQUIDO_REMESSA,
) -> float | None:
    """
    Entre débitos PIX REMESSA e TED REMESSA (GRLIS SECURITIZADORA), considera só valores em
    [valor_liquido - tolerância, valor_liquido] (teto = referência do banco / df).
    Sobre os candidatos da mesma categoria, usa o **maior** valor. Prioridade: PIX, depois TED.
    """
    ref = float(valor_liquido)
    tol = float(tolerancia)
    vmin = ref - tol
    vmax = ref

    movimentacoes = parsear_movimentacoes(dados_api)
    pix_nos: list[float] = []
    ted_nos: list[float] = []

    for mov in movimentacoes:
        kind = _classificar_pix_ted_remesa_grlis(mov)
        if not kind:
            continue
        v = float(mov["valor"])
        if v < vmin or v > vmax:
            continue
        if kind == "pix":
            pix_nos.append(v)
        else:
            ted_nos.append(v)

    if pix_nos:
        escolhido = max(pix_nos)
        print(
            f"    [MATCH PIX] Valor líquido: R$ {escolhido:,.2f} (maior na faixa "
            f"R$ {vmin:,.2f} – R$ {vmax:,.2f}, ref R$ {ref:,.2f})"
        )
        return escolhido
    if ted_nos:
        escolhido = max(ted_nos)
        print(
            f"    [MATCH TED] Valor líquido: R$ {escolhido:,.2f} (maior na faixa "
            f"R$ {vmin:,.2f} – R$ {vmax:,.2f}, ref R$ {ref:,.2f})"
        )
        return escolhido

    print(
        f"    [MATCH] Nenhum REMESSA GRLIS entre R$ {vmin:,.2f} e R$ {vmax:,.2f} "
        f"(ref R$ {ref:,.2f}, tolerância R$ {tol:,.2f})"
    )
    return None


def _valor_liquido_referencia_por_bordero(antecipacoes: list, bordero: object) -> float | None:
    """Líquido de referência do **borderô** (``Valor_Liquido_Final`` ou ``Valor_Liquido`` no join).

    Não soma vários borderôs: no Arbi, cada remessa bate o líquido **daquela** operação (ex. 1.734,99),
    não a soma de todos os borderôs do cedente no dia.
    """
    for r in antecipacoes:
        if r.get("Bordero") != bordero:
            continue
        v = r.get("Valor_Liquido_Final")
        if v is None:
            v = r.get("Valor_Liquido")
        if v is not None:
            try:
                return float(v)
            except (TypeError, ValueError):
                return None
    return None


def obter_valor_liquido_arbi_todos_cedentes(
    borderos_por_cedente: dict[str, list],
    antecipacoes: list,
) -> dict[str, dict[int, float]] | None:
    """Para cada cedente, consulta o extrato da **conta** Arbi e, para **cada borderô** distinto,
    localiza a remessa com ``buscar_valor_liquido`` usando a referência daquele borderô.
    Retorno: ``cedente`` → ``{ bordero: valor no extrato }``.
    """
    resultado: dict[str, dict[int, float]] = {}
    for cedente, borderos in borderos_por_cedente.items():
        conta = buscar_conta_por_cedente(cedente)
        if not conta:
            print(
                f"\n[FLOW] Cedente '{cedente}' sem conta Arbi no mapa. "
                "Valor líquido obrigatório via API — fluxo abortado."
            )
            return None

        print(
            f"\n[FLOW] Consultando extrato Arbi de {cedente} (conta {conta}); "
            f"um match por borderô (ref = líquido do borderô no banco)..."
        )

        extrato_api = consultar_extrato(conta)
        if isinstance(extrato_api, dict) and "erro" in extrato_api:
            print(
                f"  ERRO na API Arbi: {extrato_api['erro']}. "
                "Valor líquido não obtido — fluxo abortado."
            )
            return None

        por_bordero: dict[int, float] = {}
        for b in sorted(set(borderos)):
            ref = _valor_liquido_referencia_por_bordero(antecipacoes, b)
            if ref is None or ref <= 0:
                print(
                    f"  Borderô {b}: sem Valor_Liquido/Valor_Liquido_Final para referência. "
                    "Fluxo abortado."
                )
                return None

            print(
                f"  Borderô {b}: ref R$ {ref:,.2f} (faixa R$ {ref - TOLERANCIA_VALOR_LIQUIDO_REMESSA:,.2f} "
                f"a R$ {ref:,.2f})"
            )
            valor = buscar_valor_liquido(extrato_api, ref)
            if valor is None:
                print(
                    f"  Borderô {b}: nenhuma REMESSA GRLIS na faixa no extrato. "
                    "Fluxo abortado."
                )
                return None

            print(f"  Borderô {b}: valor no extrato R$ {valor:,.2f}")
            por_bordero[int(b)] = float(valor)

        resultado[cedente] = por_bordero

    return resultado

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