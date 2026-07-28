import pandas as pd
from pandas import Timestamp
import win32com.client as win32
import xlwings as xw
import os
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()
# Caminhos
planilha_final_path = os.getenv("planilha_final_path")

def atualizar_dados(caminho):
    print("Iniciando atualização de dados no Excel")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        wb = excel.Workbooks.Open(caminho)

        wb.RefreshAll() # Atualiza todas as conexões de dados
        excel.CalculateUntilAsyncQueriesDone() # Aguarda o término das atualizações

        wb.Save()
        wb.Close()
        print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Dados atualizados com sucesso!")
    except Exception as e:
        print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Erro ao atualizar dados: {e}")
        raise
    finally:
        excel.Quit()

atualizar_dados(planilha_final_path)
