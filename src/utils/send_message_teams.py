import html
from datetime import datetime
from typing import Any

import pandas as pd
import requests

from services.teams import get_teams_headers


def format_text(text: str) -> str:
    """Escapa HTML para o corpo da mensagem, preservando quebras ``<br>``."""
    if not text:
        return ""
    return "<br>".join(html.escape(part.strip()) for part in text.split("<br>"))


def formatar_mensagem(
    data_atual: str,
    valor_liquido_final: float,
    array_titulos: list[Any],
    nome_portal: str,
    valor_deixado: float | None = None,
) -> tuple[str, str]:
    def formatar_valor(valor: float) -> str:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    valor_formatado = formatar_valor(valor_liquido_final)

    titulos = "/".join(str(t) for t in array_titulos)

    hora = datetime.now().strftime("%H:%M")
    original = (
        f'<small>Lis  <span style="color:gray">{data_atual}, {hora}</span></small><br>'
        f"R$ {valor_formatado} - TEC TRANSPORTES EIRELI "
        f"R$ {valor_formatado} - Entrou no dia "
        f"{data_atual} na conta INTER - GRLIS DEIXADO NO 001"
    )

    label_tit = "TIT." if len(array_titulos) == 1 else "TITS."
    resposta = f"{label_tit} {titulos} - ANT {nome_portal}"

    if valor_deixado is not None and valor_deixado > 0:
        valor_deixado_fmt = formatar_valor(valor_deixado)
        resposta += f"<br><br>DEIXADO NO 001 R$ {valor_deixado_fmt} PARA AVERIGUAÇÃO"

    return original, resposta


def reply_simulado(chat_id: str, original_text: str, resposta: str) -> requests.Response:
    """Publica uma mensagem num chat do Teams (Microsoft Graph)."""
    headers = get_teams_headers()
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"

    resposta_html = format_text(resposta)

    content = (
        f"<blockquote>{original_text}</blockquote>"
        f"<p>{resposta_html}</p>"
    )

    body = {
        "body": {
            "contentType": "html",
            "content": content,
        }
    }

    response = requests.post(url, headers=headers, json=body, timeout=60)
    response.raise_for_status()
    return response


def notificar_liquidacao_conta_corrente(
    df: pd.DataFrame,
    nome_portal: str,
    chat_id: str,
    valor_deixado: float | None,
) -> requests.Response:
    """Monta data/valor/títulos a partir do ``df`` e envia ao Teams após liquidação na C/Corrente.

    ``valor_deixado``: saldo positivo deixado no 001 (fluxo *alterar*); ``None`` no fluxo *excluir*
    (``Debito_Credito`` negativo).
    """
    if "Titulo" not in df.columns:
        raise ValueError("DataFrame sem coluna Titulo para notificação Teams.")
    if "Valor_Liquido_Final" not in df.columns:
        raise ValueError("DataFrame sem coluna Valor_Liquido_Final para notificação Teams.")

    titulos = [str(x) for x in df["Titulo"].tolist()]
    ser_vlf = pd.to_numeric(df["Valor_Liquido_Final"], errors="coerce").dropna()
    if ser_vlf.empty:
        raise ValueError("Valor_Liquido_Final inválido ou nulo para notificação Teams.")
    valor_liquido_final = round(float(ser_vlf.iloc[0]), 2)

    data_atual = datetime.now().strftime("%d/%m/%Y")
    original, resposta = formatar_mensagem(
        data_atual=data_atual,
        valor_liquido_final=valor_liquido_final,
        array_titulos=titulos,
        nome_portal=nome_portal,
        valor_deixado=valor_deixado,
    )
    return reply_simulado(chat_id, original, resposta)