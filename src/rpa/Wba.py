import os
from typing import Optional, List, Dict, Any
from dotenv import load_dotenv
import time
import subprocess
import pyautogui
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import Desktop
import sys
import psutil


class WBA:
    def __init__(self):
        load_dotenv()
        self.wba_username = os.getenv("WBA_USERNAME")
        self.wba_password = os.getenv("WBA_PASSWORD")
        self.X_RELATIVE = 8
        self.Y_RELATIVE = 7
        
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