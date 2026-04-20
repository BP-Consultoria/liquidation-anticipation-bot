import os
from datetime import date
from typing import Optional, List, Dict, Any

import pandas as pd
from dotenv import load_dotenv
import time
import subprocess
import pyautogui
import pyperclip
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import Desktop
import sys
import psutil

from utils.wba_helpers import (
    atualizar_valor_no_df_por_identificador,
    calcular_ajuste_dinamico,
    codigo_cedente_unico,
    normalizar_id_titulo_dcto,
    texto_copiado_indica_dcto,
    texto_historico_desagio_padrao,
    valor_monetario_wba_campo_float,
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

        valor_str = valor_monetario_wba_campo_float(valor_total_desagio_unico(df))
        codigo_cedente_str = str(codigo_cedente_unico(df))

        app = self.app
        janela = app.window(title=titulo_janela_principal)
        janela.wait("ready", timeout=30)
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

        try:
            if janela.exists(timeout=5):
                print("[WBA] Janela 'Alertas' encontrada. Fechando...")
                time.sleep(4)
                janela.child_window(title="Fechar", control_type="Button", found_index=0).click()
            else:
                print("[WBA] Janela 'Alertas' não apareceu. Continuando...")

        except Exception as e:
            print(f"[WBA] Aviso: Erro ao interagir com alerta (pode ter fechado sozinho): {e}")

        print("[WBA] Seguindo com o fluxo...")

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

    def inserir_desagio_apos_recompra(
        self,
        codigo_busca: str = "7-COM",
        titulo_recompra: str = "Recompra (Carteira Própria)",
        titulo_deducao: str = "Dedução do Contas a Pagar",
    ) -> None:
        """Após ``recompra_carteira_propria``: aba Liberação (data do dia), dedução 7-COM e Recalcular.

        A data digitada no WBA é sempre a **data atual** (formato ``DDMMYYYY``).
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")

        app = self.app

        d = date.today()
        data_inserir = f"{d.day:02d}{d.month:02d}{d.year}"

        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.set_focus()

        aba_liberacao = janela_recompra.child_window(title="Liberação", control_type="TabItem")
        aba_liberacao.click_input()
        time.sleep(1)

        self.press_keys("{TAB}", 8)
        time.sleep(1)
        janela_recompra.type_keys("^a{BACKSPACE}")
        time.sleep(1)
        janela_recompra.type_keys(data_inserir + "{ENTER}")
        time.sleep(1)

        self.press_keys("{TAB}", 2)
        time.sleep(1)
        janela_recompra.type_keys("{ENTER}")
        time.sleep(1)

        print(f"Data {data_inserir} inserida com sucesso na aba Liberação.")

        janela_deducao = app.window(title=titulo_deducao)
        janela_deducao.wait("visible", timeout=10)
        time.sleep(1)
        janela_deducao.child_window(title="Selecionar...", control_type="Button").click()

        janela_busca = app.window(title="Busca Avançada")
        janela_busca.wait("visible", timeout=10)
        
        todos = janela_busca.descendants(title="Todos", control_type="CheckBox")
        todos[1].click()
        self.press_keys("{TAB}", 4)
        time.sleep(1)

        janela_busca.type_keys("{UP}")
        time.sleep(1)

        self.press_keys("{TAB}", 1)
        time.sleep(1)

        janela_busca.type_keys("^a{BACKSPACE}")
        time.sleep(1)
        janela_busca.type_keys(codigo_busca)
        rect = janela_busca.rectangle()
        janela_busca.type_keys("{ENTER}")
        time.sleep(0.5)
        
        janela_selecao = app.window(title="Seleção de títulos a Pagar (Carteira Própria)")
        
        grid = janela_selecao.child_window(control_type="Pane", found_index=0)  # ajuste se necessário
        rect = grid.rectangle()

        # Clica na primeira célula da primeira linha (coluna Data)
        pyautogui.click(
            x=rect.left + 40,
            y=rect.top + 25
        )

        janela_selecao = app.window(title="Seleção de títulos a Pagar (Carteira Própria)")
        janela_selecao.wait("visible", timeout=10)
        janela_selecao.child_window(title="Pagar", control_type="Button").click()

        janela_deducao = app.window(title=titulo_deducao)
        janela_deducao.wait("visible", timeout=10)
        janela_deducao.child_window(title="Fechar", control_type="Button", found_index=0).click()
        
        time.sleep(2)

        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.wait("visible", timeout=10)
        janela_recompra.set_focus()
        janela_recompra.child_window(title="Recalcular", control_type="Button").click()

    def _type_keys_horizontal(self, janela, direcao: str, passos: int, delay: float) -> None:
        tecla = "{LEFT}" if direcao.lower() == "left" else "{RIGHT}"
        for _ in range(passos):
            janela.type_keys(tecla)
            time.sleep(delay)

    def _grid_recompra_buscar_dcto_por_copia(
        self,
        janela_recompra,
        dcto_alvo: str,
        max_linhas: int,
        delay_tecla_grid: float,
        delay_apos_ctrl_c: float,
    ) -> None:
        """Com o foco já na coluna Dcto: em cada linha, Shift esquerdo+2 → Ctrl+C → compara; senão DOWN.

        O Shift+2 usa a tecla **2** da fileira superior (VK + dígito 2), não o teclado numérico.
        Pausas maiores evitam a grid “pular” linha antes do Ctrl+C. Ao achar o dcto, envia **Enter**
        para sair do modo edição/cópia da célula antes das setas até o valor.
        """
        amostras: list[tuple[int, str]] = []
        for linha in range(max_linhas):
            janela_recompra.set_focus()
            time.sleep(0.25)
            janela_recompra.type_keys("{VK_LSHIFT down}2{VK_LSHIFT up}")
            time.sleep(0.5)

            janela_recompra.set_focus()
            time.sleep(0.15)
            janela_recompra.type_keys("^c")
            time.sleep(delay_apos_ctrl_c)
            bruto = pyperclip.paste()
            if len(amostras) < 5:
                amostras.append((linha + 1, repr((bruto or "")[:100])))

            if texto_copiado_indica_dcto(bruto, dcto_alvo):
                cop_curto = (bruto or "").replace("\n", " ")[:80]
                print(
                    f"[WBA] Dcto encontrado (linha ~{linha + 1}): alvo {dcto_alvo!r} "
                    f"(trecho copiado: {cop_curto!r})"
                )
                janela_recompra.set_focus()
                time.sleep(0.15)
                janela_recompra.type_keys("{ENTER}")
                time.sleep(0.4)
                return

            janela_recompra.type_keys("{DOWN}")
            time.sleep(max(delay_tecla_grid, 0.55))

        detalhe = "; ".join(f"linha~{a[0]}: {a[1]}" for a in amostras)
        raise RuntimeError(
            f"Dcto alvo {dcto_alvo!r} não encontrado após {max_linhas} linhas. "
            "Confira setas até a coluna Dcto, Shift esquerdo+2 (fileira principal) e Ctrl+C. "
            f"Amostras: {detalhe}"
        )

    def aplicar_ajuste_debito_credito_recompra(
        self,
        df: pd.DataFrame,
        titulo_recompra: str = "Recompra (Carteira Própria)",
        *,
        coluna_identificador_grid: str = "Titulo",
        dcto_documento: str | int | None = None,
        valor_ajustado: float | None = None,
        passos_horizontal_ate_dcto: int = 8,
        direcao_horizontal_ate_dcto: str = "right",
        passos_horizontal_ate_valor: int = 8,
        delay_tecla_coluna_valor: float = 0.55,
        delay_tecla_grid: float = 0.55,
        delay_apos_ctrl_c: float = 1.1,
        delay_antes_recalcular: float = 2.5,
    ) -> tuple[pd.DataFrame, str | None]:
        """Se ``Debito_Credito`` for negativo, ajusta o valor na linha certa da grid (aba Títulos).

        Deve rodar **depois** de ``inserir_desagio_apos_recompra`` (fluxo do ``runner``), com a
        janela *Recompra (Carteira Própria)* ativa. Não usa mais contagem fixa de ``DOWN`` pela
        posição no DataFrame: vai à coluna Dcto, em cada linha faz **Shift esquerdo + 2** (tecla 2
        da fileira principal), **Ctrl+C** e compara com o dcto alvo; se não bater, **DOWN** e
        repete. Depois **8× seta esquerda** até a coluna do valor, **Backspace** (limpa a célula)
        e **Ctrl+V** com o valor do cálculo no clipboard (vírgula decimal). Fecha “Atenção”,
        aguarda e clica **Recalcular**.

        Modo **automático** (padrão): ``calcular_ajuste_dinamico``. **Explícito**: ``dcto_documento``
        e ``valor_ajustado`` juntos. Se ``Debito_Credito`` ≥ 0, devolve o ``df`` sem teclado.

        Retorno: ``(df, dcto_ajustado)``. ``dcto_ajustado`` é o identificador do título editado no
        grid (``None`` quando não houve ajuste), para fluxos como ``inserir_tag_documento_fluxo_caixa``.
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")

        explicito = dcto_documento is not None or valor_ajustado is not None
        if explicito and (dcto_documento is None or valor_ajustado is None):
            raise ValueError(
                "Modo explícito: informe os dois, ``dcto_documento`` e ``valor_ajustado``."
            )

        if explicito:
            dcto_norm = normalizar_id_titulo_dcto(dcto_documento)
            valor_novo = round(float(valor_ajustado), 2)
            df_out = atualizar_valor_no_df_por_identificador(
                df,
                coluna_identificador_grid,
                dcto_documento,
                valor_novo,
            )
            ajuste_valor = valor_novo
            print(
                f"[WBA] Ajuste explícito no df e na grid: dcto={dcto_norm!r}, valor={valor_novo}"
            )
        else:
            df_out, residual, ajuste = calcular_ajuste_dinamico(
                df, coluna_identificador_grid=coluna_identificador_grid
            )
            if not ajuste:
                if residual > 0:
                    print(f"[WBA] Debito_Credito positivo ({residual}); sem edição no grid.")
                else:
                    print("[WBA] Debito_Credito ≥ 0; sem ajuste no grid.")
                return df_out, None
            dcto_norm = str(ajuste["dcto_alvo"])
            ajuste_valor = float(ajuste["valor"])

        app = self.app
        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.wait("visible", timeout=10)
        janela_recompra.set_focus()

        tab_titulos = janela_recompra.child_window(
            title="Títulos",
            control_type="TabItem"
        )
        tab_titulos.click_input()

        time.sleep(1)

        janela_recompra.set_focus()

        self.press_keys("{TAB}", 7)
        time.sleep(0.65)

        self.press_keys("{ENTER}", 1)
        time.sleep(0.65)

        self._type_keys_horizontal(
            janela_recompra,
            direcao_horizontal_ate_dcto,
            passos_horizontal_ate_dcto,
            delay_tecla_grid,
        )

        max_linhas = max(len(df) + 20, 60)
        self._grid_recompra_buscar_dcto_por_copia(
            janela_recompra,
            dcto_norm,
            max_linhas,
            delay_tecla_grid,
            delay_apos_ctrl_c,
        )

        janela_recompra.set_focus()
        time.sleep(0.55)
        self._type_keys_horizontal(
            janela_recompra,
            "left",
            passos_horizontal_ate_valor,
            delay_tecla_coluna_valor,
        )

        valor_str = f"{ajuste_valor:.2f}".replace(".", ",")
        time.sleep(0.45)
        janela_recompra.set_focus()
        janela_recompra.type_keys("^a")
        time.sleep(0.15)
        janela_recompra.type_keys("{BACKSPACE}")
        time.sleep(0.2)
        pyperclip.copy(valor_str)
        time.sleep(0.15)
        janela_recompra.set_focus()
        janela_recompra.type_keys("^v")
        time.sleep(0.35)
        janela_recompra.type_keys("{ENTER}")

        print(f"[WBA] Valor colado (Ctrl+V): {valor_str} (dcto {dcto_norm})")

        time.sleep(1.0)

        janela_atencao = janela_recompra.child_window(
            title="Atenção", control_type="Window"
        )
        janela_atencao.wait("visible", timeout=10)
        janela_atencao.child_window(title="OK", control_type="Button").click()

        time.sleep(delay_antes_recalcular)
        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.wait("visible", timeout=15)
        janela_recompra.set_focus()
        janela_recompra.child_window(title="Recalcular", control_type="Button").click()
        time.sleep(3)
        print("[WBA] Recalcular clicado após Atenção.")

        return df_out, dcto_norm

    def preencher_valor_total_aba_renegociacao(
        self,
        df: pd.DataFrame,
        titulo_recompra: str = "Recompra (Carteira Própria)",
        *,
        coluna_valor: str = "Valor",
        passos_tab_ate_campo: int = 8,
        delay_apos_aba: float = 1.0,
        delay_apos_tabs: float = 1.0,
        delay_apos_enter_valor: float = 0.5,
        delay_apos_recalcular: float = 2.0,
    ) -> None:
        """Aba *Renegociação*: preenche o campo com a **soma da coluna ``Valor``** (face de cada título).

        Cada linha do ``df`` é um título; o valor enviado ao WBA é ``sum(Valor)`` de todas as linhas.
        Deve rodar **depois** de ``aplicar_ajuste_debito_credito_recompra``. Ajuste
        ``passos_tab_ate_campo`` se o foco não cair no campo certo.
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")
        if coluna_valor not in df.columns:
            raise ValueError(f"DataFrame sem coluna {coluna_valor!r}.")

        n_titulos = len(df)
        soma_valor_titulos = round(
            float(pd.to_numeric(df[coluna_valor], errors="coerce").fillna(0).sum()),
            2,
        )
        valor_formatado = f"{soma_valor_titulos:.2f}".replace(".", ",")

        print(
            f"[WBA] Renegociação — soma de {coluna_valor} ({n_titulos} título(s)): R$ {valor_formatado}"
        )

        app = self.app
        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.wait("visible", timeout=15)
        janela_recompra.set_focus()

        aba_renegociacao = janela_recompra.child_window(
            title="Renegociação", control_type="TabItem"
        )
        aba_renegociacao.click_input()
        time.sleep(delay_apos_aba)

        janela_recompra.set_focus()
        self.press_keys("{TAB}", passos_tab_ate_campo)
        time.sleep(delay_apos_tabs)

        janela_recompra.type_keys(valor_formatado + "{ENTER}")
        time.sleep(delay_apos_enter_valor)

        btn_recalcular = janela_recompra.child_window(
            title="Recalcular", control_type="Button"
        )
        btn_recalcular.click_input()
        time.sleep(delay_apos_recalcular)

        print(
            f"[WBA] Renegociação: total (soma {coluna_valor}) R$ {valor_formatado} — Recalcular acionado."
        )

    def liberar_concluir_etapa_recompra(
        self,
        titulo_recompra: str = "Recompra (Carteira Própria)",
        *,
        passos_tab_ate_campo: int = 16,
        texto_campo: str = "77",
        passos_tab_apos_campo: int = 2,
        delay_apos_aba: float = 1.0,
        delay_apos_tabs_inicial: float = 0.3,
        delay_apos_digitacao: float = 0.3,
        delay_apos_tabs_final: float = 0.3,
        delay_apos_recalcular: float = 2.0,
        delay_apos_liberar: float = 2.0,
    ) -> None:
        """Último passo do fluxo na *Recompra*: aba *Liberação*, campo (TAB), digitação, *Recalcular* e *Liberar*.

        Rodar **depois** de ``preencher_valor_total_aba_renegociacao``. Ajuste
        ``passos_tab_ate_campo`` / ``texto_campo`` conforme o layout do WBA.
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")

        app = self.app
        janela_recompra = app.window(title=titulo_recompra)
        janela_recompra.wait("visible", timeout=15)
        janela_recompra.set_focus()

        aba_liberacao = janela_recompra.child_window(
            title="Liberação", control_type="TabItem"
        )
        aba_liberacao.click_input()
        time.sleep(delay_apos_aba)

        janela_recompra.set_focus()
        self.press_keys("{TAB}", passos_tab_ate_campo, delay=delay_apos_tabs_inicial)
        time.sleep(delay_apos_tabs_inicial)

        janela_recompra.type_keys(texto_campo)
        time.sleep(delay_apos_digitacao)

        self.press_keys("{TAB}", passos_tab_apos_campo, delay=delay_apos_tabs_final)
        time.sleep(delay_apos_tabs_final)

        janela_recompra.set_focus()
        btn_recalcular = janela_recompra.child_window(
            title="Recalcular", control_type="Button"
        )
        btn_recalcular.click_input()
        time.sleep(delay_apos_recalcular)

        janela_recompra.set_focus()
        btn_liberar = janela_recompra.child_window(
            title="Liberar", control_type="Button"
        )
        btn_liberar.click_input()
        time.sleep(delay_apos_liberar)
        
        janela_recompra.set_focus()
        janela_recompra.child_window(title="Fechar", control_type="Button").click()

        print(
            "[WBA] Liberação: campo preenchido, Recalcular e Liberar acionados (etapa concluída)."
        )

    @staticmethod
    def _resolver_lista_tags(janela_tags) -> object:
        """Retorna o controle *List* (ou *Table*) onde os ``ListItem`` das tags aparecem — como no Jupyter."""
        for fi in (0, 1, 2):
            try:
                lst = janela_tags.child_window(control_type="List", found_index=fi)
                lst.wait("exists", timeout=2)
                return lst
            except Exception:
                continue
        desc = janela_tags.descendants(control_type="List")
        if desc:
            return desc[0]
        tab = janela_tags.descendants(control_type="Table")
        if tab:
            return tab[0]
        return janela_tags

    def inserir_tag_documento_fluxo_caixa(
        self,
        df: pd.DataFrame,
        dcto: str | int,
        *,
        titulo_janela_principal: str = "WBA Securitização - Versão: 24.7.1 (Build: 6847)",
        titulo_manutencao_fluxo: str = "Manutenção do Fluxo de Caixa (Carteira Própria)",
        imagem_filtro: str = r"C:\Users\suporte\Documents\imagens\filtro.png",
        imagem_tags: str = r"C:\Users\suporte\Documents\imagens\tags.png",
        texto_tag: str = "SALDO DE ANTECIPAÇÂO INSUFICIENTE",
        confidence: float = 0.8,
        max_pgdn_lista_tag: int = 20,
        pyautogui_write_interval: float = 0.0,
    ) -> None:
        """Espelha o fluxo do Jupyter: Contas→Lançamentos, filtro, busca (cedente + dcto), tags, *Recompra Cedente*.

        ``lista`` é o ``List`` da janela ``Tags do documento: <dcto>`` (``found_index=0`` com fallback).
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")

        dcto_str = str(normalizar_id_titulo_dcto(dcto)).strip()
        codigo_str = str(codigo_cedente_unico(df))
        titulo_tags = f"Tags do documento: {dcto_str}"

        app = self.app
        janela = app.window(title=titulo_janela_principal)
        janela.wait("visible", timeout=10)
        janela.set_focus()
        janela.menu_select("Contas->Lançamentos")
        time.sleep(2)

        posicao = pyautogui.locateCenterOnScreen(imagem_filtro, confidence=confidence)
        if not posicao:
            raise RuntimeError("Imagem filtro não encontrada na tela.")
        pyautogui.click(posicao)

        janela = app.window(title="Busca Avançada")
        janela.wait("visible", timeout=15)
        time.sleep(1)

        janela.child_window(
            title="Todos", control_type="CheckBox", found_index=1
        ).click_input()
        time.sleep(1)

        cedente_checkbox = janela.child_window(title="Cedente", control_type="CheckBox")
        cedente_checkbox.set_focus()
        cedente_checkbox.toggle()
        time.sleep(1)

        self.press_keys("{TAB}", 2, delay=0.05)
        janela.type_keys(codigo_str)
        time.sleep(1)
        self.press_keys("{TAB}", 8, delay=0.05)
        time.sleep(1)
        self.press_keys("{UP}", 1, delay=0.05)
        time.sleep(1)
        self.press_keys("{TAB}", 1, delay=0.05)
        janela.type_keys(dcto_str)
        time.sleep(2)
        self.press_keys("{ENTER}", 1, delay=0.05)

        janela = app.window(title=titulo_manutencao_fluxo)
        janela.wait("visible", timeout=15)

        posicao = pyautogui.locateCenterOnScreen(imagem_tags, confidence=confidence)
        if not posicao:
            raise RuntimeError("Imagem tags não encontrada na tela.")
        pyautogui.click(posicao)

        janela = app.window(title=titulo_tags)
        janela.wait("visible", timeout=15)

        try:
            lista = janela.child_window(control_type="List", found_index=0)
            lista.wait("exists", timeout=5)
        except Exception:
            lista = self._resolver_lista_tags(janela)

        lista.set_focus()

        item_ok = False
        for _ in range(max_pgdn_lista_tag):
            try:
                item = lista.child_window(
                    title_re=".*Recompra Cedente.*",
                    control_type="ListItem",
                )
                item.double_click_input()
                item_ok = True
                break
            except Exception:
                lista.type_keys("{PGDN}")

        if not item_ok:
            raise RuntimeError(
                "ListItem com título contendo 'Recompra Cedente' não encontrado na lista de tags."
            )

        time.sleep(1)
        pyautogui.write(texto_tag, interval=pyautogui_write_interval or 0)
        time.sleep(1)
        self.press_keys("{TAB}", 1, delay=0.05)
        time.sleep(1)
        self.press_keys("{ENTER}", 1, delay=0.05)
        time.sleep(1)
        self.press_keys("{ENTER}", 1, delay=0.05)

        janela = app.window(title=titulo_manutencao_fluxo)
        janela.wait("visible", timeout=15)
        time.sleep(2)
        janela.set_focus()
        janela.child_window(title="Fechar", control_type="Button").click()

        print(
            f"[WBA] Tag inserida no documento {dcto_str!r} (texto: {texto_tag!r})."
        )

    def _enviar_teams_liquidacao_cc(
        self,
        df: pd.DataFrame,
        valor_deixado: Optional[float],
        nome_portal_teams: Optional[str],
        teams_chat_id: Optional[str],
    ) -> None:
        portal = (nome_portal_teams or "").strip()
        cid = (teams_chat_id or "").strip()
        if not portal or not cid:
            return
        try:
            from utils.send_message_teams import notificar_liquidacao_conta_corrente

            notificar_liquidacao_conta_corrente(df, portal, cid, valor_deixado)
            print("[WBA] Notificação Teams enviada (liquidação C/Corrente).")
        except Exception as exc:
            print(f"[WBA] Aviso: notificação Teams não enviada ({exc}).")

    def processar_conta_corrente_pos_liberacao(
        self,
        df: pd.DataFrame,
        *,
        nome_portal_teams: Optional[str] = None,
        teams_chat_id: Optional[str] = None,
        titulo_janela_principal: str = "WBA Securitização - Versão: 24.7.1 (Build: 6847)",
        titulo_manutencao_cc: str = "Manutenção do Conta Corrente (Carteira Própria)",
        codigo_filtro_conta: str = "001",
        imagem_filtro: str = r"C:\Users\suporte\Documents\imagens\filtro.png",
        imagem_alterar: str = r"C:\Users\suporte\Documents\imagens\alterar.png",
        imagem_salvar: str = r"C:\Users\suporte\Documents\imagens\salvar.png",
        imagem_excluir: str = r"C:\Users\suporte\Documents\imagens\excluir.png",
        texto_motivo_exclusao: str = "lançamento liquidado no 040 by: Lis",
        confidence: float = 0.8,
        delay_menu: float = 2.0,
    ) -> None:
        """Após ``liberar_concluir_etapa_recompra``: C/Corrente → Lançamentos, *Busca Avançada* com
        ``Valor_Liquido_Final`` nas duas faixas **Valores de:** (mesmo valor repetido).

        Se ``nome_portal_teams`` e ``teams_chat_id`` estiverem preenchidos, envia Teams após
        excluir (``Debito_Credito`` < 0, ``valor_deixado`` = ``None``) ou após alterar valor
        (``Debito_Credito`` > 0, ``valor_deixado`` = saldo positivo ``dc``).
        """
        if not hasattr(self, "app") or self.app is None:
            raise RuntimeError("Application not started; call start_wba_application first.")
        if "Valor_Liquido_Final" not in df.columns:
            raise ValueError("Coluna Valor_Liquido_Final ausente.")
        if "Debito_Credito" not in df.columns:
            raise ValueError("Coluna Debito_Credito ausente.")

        ser_vlf = pd.to_numeric(df["Valor_Liquido_Final"], errors="coerce").dropna()
        if ser_vlf.empty:
            raise ValueError("Valor_Liquido_Final inválido ou nulo em todas as linhas do lote.")
        v_filtro = round(float(ser_vlf.iloc[0]), 2)
        col_usada = "Valor_Liquido_Final"

        valor_filtro_str = f"{v_filtro:.2f}".replace(".", ",")

        dc_num = pd.to_numeric(df["Debito_Credito"].iloc[0], errors="coerce")
        dc = 0.0 if pd.isna(dc_num) else round(float(dc_num), 2)

        if dc == 0.0:
            print(
                "[WBA] Conta Corrente pós-liberação: Debito_Credito = 0; "
                "fluxo excluir/alterar não executado."
            )
            return

        print(
            f"[WBA] Conta Corrente: filtro 'Valores de:' = R$ {valor_filtro_str} "
            f"(coluna {col_usada})."
        )

        def _fmt_dc_alteracao(valor: float) -> str:
            return f"{round(abs(float(valor)), 2):.2f}".replace(".", ",")

        app = self.app
        janela = app.window(title=titulo_janela_principal)
        janela.wait("visible", timeout=10)
        janela.set_focus()
        janela.menu_select("C/Corrente->Lançamentos de Caixa e Bancos")
        time.sleep(delay_menu)

        posicao = pyautogui.locateCenterOnScreen(imagem_filtro, confidence=confidence)
        if not posicao:
            raise RuntimeError("Imagem filtro não encontrada na tela.")

        pyautogui.click(posicao)

        janela_busca = app.window(title="Busca Avançada")
        janela_busca.wait("visible", timeout=15)
        time.sleep(1)

        checkbox_todos = janela_busca.child_window(
            title="Todos", control_type="CheckBox"
        )
        checkbox_todos.wait("visible", timeout=10)
        checkbox_todos.click_input()
        time.sleep(1)

        self.press_keys("{TAB}", 2)
        time.sleep(1)
        self.press_keys("{RIGHT}", 2)
        time.sleep(1)
        self.press_keys("{TAB}", 2)
        time.sleep(1)
        janela_busca.type_keys(codigo_filtro_conta)
        time.sleep(1)

        checkbox_valores = janela_busca.child_window(
            title="Valores de:", control_type="CheckBox"
        )
        checkbox_valores.wait("visible", timeout=10)
        checkbox_valores.click_input()
        time.sleep(1)

        self.press_keys("{TAB}", 1)
        time.sleep(1)
        janela_busca.type_keys(valor_filtro_str)
        time.sleep(1)
        self.press_keys("{TAB}", 1)
        time.sleep(1)
        janela_busca.type_keys(valor_filtro_str)
        time.sleep(1)
        self.press_keys("{TAB}", 2)
        self.press_keys("{ENTER}", 1)

        time.sleep(1)
        janela_cc = app.window(title=titulo_manutencao_cc)
        janela_cc.wait("visible", timeout=15)
        janela_cc.set_focus()

        if dc < 0:
            print(f"[WBA] Conta Corrente: Debito_Credito negativo ({dc:.2f}); fluxo excluir.")
            pos_exc = pyautogui.locateCenterOnScreen(imagem_excluir, confidence=confidence)
            if not pos_exc:
                raise RuntimeError("Imagem excluir não encontrada na tela.")
            pyautogui.click(pos_exc)

            time.sleep(1)
            self.press_keys("{LEFT}", 1)
            time.sleep(1)
            self.press_keys("{ENTER}", 1)
            time.sleep(2)
            send_keys(texto_motivo_exclusao, with_spaces=True)
            time.sleep(2)
            self.press_keys("{TAB}", 1)
            time.sleep(1)
            self.press_keys("{ENTER}", 1)

            time.sleep(2)
            janela_cc = app.window(title=titulo_manutencao_cc)
            janela_cc.set_focus()
            janela_cc.child_window(title="Fechar", control_type="Button").click_input()
            print("[WBA] Conta Corrente: exclusão concluída e janela fechada.")
            self._enviar_teams_liquidacao_cc(
                df, valor_deixado=None, nome_portal_teams=nome_portal_teams, teams_chat_id=teams_chat_id
            )
            return

        # dc > 0
        print(f"[WBA] Conta Corrente: Debito_Credito positivo ({dc:.2f}); fluxo alterar valor.")
        valor_alterar_str = _fmt_dc_alteracao(dc)
        pos_alt = pyautogui.locateCenterOnScreen(imagem_alterar, confidence=confidence)
        if not pos_alt:
            raise RuntimeError("Imagem alterar não encontrada na tela.")
        pyautogui.click(pos_alt)
        time.sleep(1)
        self.press_keys("{TAB}", 5)
        time.sleep(1)
        janela_cc.type_keys(valor_alterar_str)
        self.press_keys("{TAB}", 1)

        pos_salvar = pyautogui.locateCenterOnScreen(imagem_salvar, confidence=confidence)
        if not pos_salvar:
            raise RuntimeError("Imagem salvar não encontrada na tela.")
        pyautogui.click(pos_salvar)
        print(
            f"[WBA] Conta Corrente: valor alterado para {valor_alterar_str} (|Debito_Credito|) e salvo."
        )
        time.sleep(2)
        janela_cc = app.window(title=titulo_manutencao_cc)
        janela_cc.set_focus()
        janela_cc.child_window(title="Fechar", control_type="Button").click_input()
        print("[WBA] Conta Corrente: alteração concluída e janela fechada.")
        self._enviar_teams_liquidacao_cc(
            df, valor_deixado=dc, nome_portal_teams=nome_portal_teams, teams_chat_id=teams_chat_id
        )