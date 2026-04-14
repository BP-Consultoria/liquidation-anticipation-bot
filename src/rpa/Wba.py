import os
from datetime import date
from typing import Optional, List, Dict, Any

import pandas as pd
from dotenv import load_dotenv
import time
import subprocess
import pyautogui
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import Desktop
import sys
import psutil

from utils.wba_helpers import (
    codigo_cedente_unico,
    texto_historico_desagio_padrao,
    valor_monetario_br,
    valor_total_desagio_unico,
)


class WBA:
    def __init__(self):
        load_dotenv()
        self.wba_username = os.getenv("WBA_USERNAME")
        self.wba_password = os.getenv("WBA_PASSWORD")
        
        self.start_wba_application()
    
    def login(self) -> bool:
        try:
            actual_window = self.app.window(title="Gerenciador de Conexões")
            
            time.sleep(1)
            actual_window.child_window(title="Conectar", class_name="TBitBtn").click()
            time.sleep(1)
            actual_window = self.app.window(title="Identificação")
            
            login_input = actual_window.child_window(class_name="TEdit", found_index=1).wrapper_object()
            time.sleep(1)
            login_input.set_focus()
            time.sleep(1)
            login_input.set_edit_text(self.wba_username)
            
            password_input = actual_window.child_window(class_name="TEdit", found_index=0).wrapper_object()
            time.sleep(1)
            password_input.set_focus() 
            time.sleep(1)
            password_input.set_edit_text(self.wba_password)
            
            actual_window.child_window(title="OK", class_name="TBitBtn").click()
            time.sleep(5)

        except Exception as exc:
            print(f"Erro ao realizar login: {exc}")
    
    def start_wba_application(self) -> None:
        caminho_executavel = r"C:\Program Files (x86)\WBA\Securitizacao\Securitizacao.exe"
        pasta_trabalho = r"C:\Program Files (x86)\WBA\Securitizacao"
        nome_processo = "Securitizacao.exe"

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == nome_processo.lower():
                    proc.terminate()
                    proc.wait(timeout=5)
            except:
                pass

        time.sleep(1)

        try:
            self.app = Application(backend="uia").start(
                cmd_line=caminho_executavel,
                work_dir=pasta_trabalho
            )
        except Exception as e:
            print(f"Erro ao iniciar: {e}")
            
    def close_wba_application(self) -> None:
        nome_processo = "Securitizacao.exe"
        
        try:
            if hasattr(self, 'app') and self.app:
                try:
                    self.app.kill()
                except:
                    pass
            
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and proc.info['name'].lower() == nome_processo.lower():
                        proc.terminate()
                        proc.wait(timeout=5)
                except:
                    pass
            
            print("WBA application closed")
        except Exception as e:
            print(f"Error closing WBA application: {e}")

    def press_keys(self, tecla: str, vezes: int, delay: float = 0.3) -> None:
        for _ in range(vezes):
            send_keys(tecla)
            time.sleep(delay)

    def lancar_desagio_contas_lancamentos(
        self,
        df: pd.DataFrame,
        imagem_open_titulo: str = r"C:\Users\suporte\Documents\imagens\open_titulo.png",
        imagem_salvar: str = r"C:\Users\suporte\Documents\imagens\salvar.png",
        titulo_janela_principal: str = "WBA Securitização - Versão: 24.7.1 (Build: 6847)",
        codigo_conta: str = "286",
        codigo_busca: str = "7-COM",
        texto_historico: str | None = None,
        data_referencia_historico: date | None = None,
        confidence: float = 0.8,
    ) -> None:
        """Contas → Lançamentos: lança deságio na Manutenção do Fluxo de Caixa (Carteira Própria).

        O valor digitado no WBA é **um único** ``Valor_Total_Desagio`` (primeiro valor válido da coluna;
        não soma — o banco repete o mesmo número em cada linha do lote).
        ``df`` é o lote de um único cedente (``codigo_cedente`` no DataFrame).

        Se ``texto_historico`` for omitido, usa ``texto_historico_desagio_padrao(data_referencia_historico)``
        (mês abreviado + ano dinâmicos).
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")

        hist = texto_historico if texto_historico is not None else texto_historico_desagio_padrao(
            data_referencia_historico
        )

        valor_str = valor_monetario_br(valor_total_desagio_unico(df))
        codigo_cedente_str = str(codigo_cedente_unico(df))

        app = self.app
        janela = app.window(title=titulo_janela_principal)
        janela.wait("ready", timeout=10)
        janela.menu_select("Contas->Lançamentos")

        janela = app.window(title="Manutenção do Fluxo de Caixa (Carteira Própria)")
        janela.wait("ready", timeout=10)

        posicao = pyautogui.locateCenterOnScreen(imagem_open_titulo, confidence=confidence)
        if not posicao:
            raise RuntimeError("Imagem open_titulo não encontrada na tela.")

        pyautogui.click(posicao)

        self.press_keys("{TAB}", 2)
        send_keys(codigo_conta)

        self.press_keys("{TAB}", 2)
        send_keys(codigo_busca)

        self.press_keys("{TAB}", 1)
        send_keys(valor_str)

        self.press_keys("{TAB}", 2)
        send_keys(hist, with_spaces=True)

        self.press_keys("{TAB}", 1)
        send_keys(codigo_cedente_str)

        self.press_keys("{TAB}", 1)

        posicao_salvar = pyautogui.locateCenterOnScreen(imagem_salvar, confidence=confidence)
        if not posicao_salvar:
            raise RuntimeError("Imagem salvar não encontrada na tela.")

        pyautogui.click(posicao_salvar)

        historico = app.window(title="Adiciona Histórico")
        historico.wait("ready", timeout=10)
        historico.child_window(title="OK", control_type="Button").click()

        janela = app.window(title="Manutenção do Fluxo de Caixa (Carteira Própria)")
        janela.wait("ready", timeout=10)
        janela.child_window(title="Fechar", control_type="Button").click()

    def recompra_carteira_propria(
        self,
        df: pd.DataFrame,
        imagem_path: str = r"C:\Users\suporte\Documents\imagens\novo.png",
        selectall_path: str = r"C:\Users\suporte\Documents\imagens\selectall_button.png",
        titulo_janela_principal: str = "WBA Securitização - Versão: 24.7.1 (Build: 6847)",
        confidence: float = 0.8,
    ) -> None:
        """Recompra (Carteira Própria): informa código do cedente e inclui cada ``Titulo`` do ``df`` na busca."""
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")
        if "Titulo" not in df.columns:
            raise ValueError("DataFrame deve conter a coluna 'Titulo'.")

        codigo_str = str(codigo_cedente_unico(df))
        titulos = df["Titulo"].tolist()
        if not titulos:
            raise ValueError("Nenhum título no DataFrame.")

        app = self.app
        janela = app.window(title=titulo_janela_principal)
        janela.wait("visible", timeout=10)
        janela.set_focus()

        botao = janela.child_window(auto_id="BotaoRegresso", control_type="Image").wrapper_object()
        botao.click_input()

        time.sleep(3)
        janela = app.window(title="Recompra (Carteira Própria)")

        posicao = pyautogui.locateCenterOnScreen(imagem_path, confidence=confidence)
        if not posicao:
            raise RuntimeError("Imagem salvar não encontrada na tela.")

        pyautogui.click(posicao)

        janela.type_keys(codigo_str)
        janela.type_keys("{TAB}")

        janela = app.window(title="Alertas")
        time.sleep(4)
        janela.child_window(title="Fechar", control_type="Button", found_index=0).click()

        janela = app.window(title="Recompra (Carteira Própria)")

        time.sleep(1)
        janela.type_keys("{TAB}")
        time.sleep(1)
        janela.type_keys("0")
        time.sleep(1)
        janela.type_keys("{TAB}")
        time.sleep(1)
        janela.type_keys("0")
        time.sleep(1)
        janela.type_keys("{TAB}")
        time.sleep(1)
        janela.type_keys("0")
        time.sleep(1)
        janela.type_keys("{TAB}")
        time.sleep(1)
        janela.type_keys("0")
        time.sleep(1)
        janela.type_keys("{TAB}")
        time.sleep(1)

        janela.type_keys("{ENTER}")

        janela_busca = app.window(title="Busca Avançada")

        for i, titulo in enumerate(titulos):
            is_last = i == len(titulos) - 1
            todos = janela_busca.descendants(title="Todos", control_type="CheckBox")
            todos[1].click()
            self.press_keys("{TAB}", 4)
            time.sleep(0.5)

            janela_busca.type_keys("{UP}")
            time.sleep(0.3)

            self.press_keys("{TAB}", 1)
            time.sleep(0.3)

            janela_busca.type_keys("^a{BACKSPACE}")
            janela_busca.type_keys(str(titulo))
            janela_busca.type_keys("{ENTER}")

            janela_selecao = app.window(title="Seleção de títulos a Recomprar (Carteira Própria)")
            janela_selecao.wait("visible", timeout=10)

            time.sleep(4)

            posicao_sel = pyautogui.locateCenterOnScreen(selectall_path, confidence=confidence)
            if not posicao_sel:
                raise RuntimeError(f"Imagem Select All não encontrada para título {titulo}")

            pyautogui.click(posicao_sel)

            time.sleep(1)

            janela_selecao.child_window(title="Recomprar", control_type="Button").click()

            time.sleep(1)

            janela_recompra = app.window(title="Recompra (Carteira Própria)")
            janela_recompra.wait("visible", timeout=10)
            janela_recompra.set_focus()

            if not is_last:
                janela_recompra.type_keys("{ENTER}")
            else:
                time.sleep(1)
                btn_recalcular = janela_recompra.child_window(title="Recalcular", control_type="Button")
                btn_recalcular.click()

            time.sleep(2)

            if not is_last:
                janela_busca = app.window(title="Busca Avançada")
                janela_busca.wait("visible", timeout=10)