import customtkinter as ctk
import tkinter as tk
from pathlib import Path
from threading import Thread
import pythoncom
from almap_africa.almap_africa_refatorado_copy import executar as almap_africa
from Indicador_class.indicador_class_copy_2 import executar as indicador_class
from shared.logger_utils import get_logger, TextboxHandler

# Funções
def alterar_estado_botao(habilitado: bool):
    estado = "normal" if habilitado else "disabled"

    botoes = [botao_almap, botao_indicador_class, botao_limpar]

    for botao in botoes:
        botao.configure(state=estado)

def limpar_textbox(textbox):
    textbox.delete("1.0", tk.END)

def executar_almap_africa():

    alterar_estado_botao(False)

    def tarefa():
        try:
            pythoncom.CoInitialize()
            logger = get_logger(
                nome_arquivo="almap_africa",
                pasta_logs=BASE_DIR / "almap_africa",
                extra_handlers=[textbox_handler]
            )

            limpar_textbox(textbox_log)
            almap_africa(logger)
        
        finally:
            pythoncom.CoUninitialize()
            root.after(0, lambda: alterar_estado_botao(True))

    Thread(target=tarefa, daemon=True).start()

def executar_Indicador_class():

    alterar_estado_botao(False)

    def tarefa():
        try:
            pythoncom.CoInitialize()
            logger = get_logger(
                nome_arquivo="indicador_class",
                pasta_logs=BASE_DIR / "indicador_class",
                extra_handlers=[textbox_handler]
            )

            limpar_textbox(textbox_log)
            indicador_class(logger)
        
        finally:
            pythoncom.CoUninitialize()
            root.after(0, lambda: alterar_estado_botao(True))

    Thread(target=tarefa, daemon=True).start()

BASE_DIR = Path(__file__).resolve().parent

root = ctk.CTk()
root.geometry("400x300")
root.title("Automações")

# Tabs
tabControl = ctk.CTkTabview(root, segmented_button_selected_color="#7E15C0",
                            segmented_button_unselected_color="#666666",
                            segmented_button_selected_hover_color="#570D85",
                            segmented_button_unselected_hover_color="#555555")
tabControl.pack(expand=True, fill="both")

diario = tabControl.add("Diário")
terca = tabControl.add("Terça")
quarta = tabControl.add("Quarta")
quinta = tabControl.add("Quinta")
mensal = tabControl.add("Mensal")

#Aba Diário
botao_almap = ctk.CTkButton(
    diario, text="Almap Africa", command=executar_almap_africa)

botao_almap.pack(pady=10)

botao_indicador_class = ctk.CTkButton(diario, text="Indicador de classificação", command= lambda: executar_Indicador_class())
botao_indicador_class.pack(pady=10)

textbox_log = ctk.CTkTextbox(
    diario,
    width=350,
    height=150
)
textbox_log.pack(pady=10)
textbox_handler = TextboxHandler(textbox_log)

botao_limpar = ctk.CTkButton(diario, text="Limpar logs", command=lambda: limpar_textbox(textbox_log))
botao_limpar.pack(pady=10)

root.mainloop()
