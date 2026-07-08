from gerador_escalas_desktop import GeradorEscalasApp
import tkinter as tk

class DummyApp(GeradorEscalasApp):
    def __init__(self):
        self.log_messages = []
        self.btn_process = type('obj', (object,), {'config': lambda state=None: None})()
        pass
    def log(self, message):
        self.log_messages.append(message)
        print(message)
        
app = DummyApp()
app.process_data("Check pre envio-Macro Colunas - 15052026.xlsm")
