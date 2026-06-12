import pandas as pd
from pandas import Timestamp
import win32com.client as win32
import xlwings as xw
import pyautogui as tec
import time
import os
from dotenv import load_dotenv

load_dotenv()
# Caminhos
planilha_baixada_path = os.getenv("planilha_baixada_path")
planilha_final_path = os.getenv("planilha_final_path")
planilha_colab = os.getenv("planilha_colab")
# Destinatários
destinatarios_indicador = os.getenv("destinatarios_indicador")
cc_indicador = os.getenv("cc_indicador")

print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Processo iniciado")

def atualizar_dados(caminho):
    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Iniciando atualização de dados no Excel")
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
    finally:
        excel.Quit()

def enviar_email(destinatario, cc):
    saudacao = "Bom dia, pessoal." if Timestamp.now().hour < 12 else "Boa tarde, pessoal."

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Display()
    assinatura = mail.HtmlBody

    mail.to = destinatario
    mail.CC = cc
    mail.Subject = "Indicador de Classificação - 2026"

    mail.HtmlBody = Rf"""<div style="font-family: tahoma; font-size: 11pt">
        <p>{saudacao}</p>
        <p>O indicador de classificação está atualizado até {Timestamp.now().strftime('%d/%m/%Y')}</p>
        <p>Arquivo: <a href="https://adgbl.sharepoint.com/:x:/r/sites/Relatriosgerenciais/Arquivos/2026/Analises/Classifica%C3%A7%C3%A3o/Indicador - classifica%C3%A7%C3%A3o 2026.xlsx?d=w760aea90c357485589274b4504d94c34&csf=1&web=1&e=8rjrfK&xsdata=MDV8MDJ8bWF0aGV1cy5waW50b0BpYm9wZS5jb218MDc3ZDczNDFhMGQ5NGJhM2FlM2QwOGRlOTRiM2Q0YmR8YjI3NjcyNDFmYWI1NDU0YjhiNjJmNjMyNDY1MGUzMTZ8MHwwfDYzOTExMTY5NzI1Nzk4MTI4OHxVbmtub3dufFRXRnBiR1pzYjNkOGV5SkZiWEIwZVUxaGNHa2lPblJ5ZFdVc0lsWWlPaUl3TGpBdU1EQXdNQ0lzSWxBaU9pSlhhVzR6TWlJc0lrRk9Jam9pVFdGcGJDSXNJbGRVSWpveWZRPT18MHx8fA%3d%3d&sdata=QlpMbitCVUtmelAwcUIxQU5rTXpXZjNETnpuWE9Vb3lhR3dGaHQwT2hCbz0%3d">Indicador - classificação 2026.xlsx</a></p>
        <p>Para acessar o arquivo, acesse via desktop.</p>
        <p style="font-family: tahoma; font-size: 9pt; color: #555;"><i>E-mail enviado automaticamente.</i></p>
        {assinatura}
    </div>
    """
    mail.Display()
    time.sleep(3)
    tec.hotkey("ctrl", "enter")

print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Tranformando dados")

df_baixada = pd.read_excel(planilha_baixada_path, skiprows=1)
df_baixada = df_baixada.drop(columns=["Nome Tela", "Versão AG"])
df_baixada = df_baixada.drop(index=df_baixada[df_baixada.iloc[:, 3].isna()].index) # apaga a soma no fim da planilha

df_colab = pd.read_excel(planilha_colab, sheet_name="Planilha1")

df_baixada.iloc[:, 3] = pd.to_datetime(df_baixada.iloc[:, 3]) # garabte q a coluna está no formato datetime para extrair mes e ano

meses = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
    5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
    9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
}

df_baixada["CHAVE"] = (
    df_baixada.iloc[:, 2].astype(str).str.strip() + " - " + # pega o nome
    df_baixada.iloc[:, 3].apply( # pega a data
        lambda x: f"{meses[x.month]}/{str(x.year)[2:]}" # formata a data como "MES/ANO"
    ) # resultado final: "NOME - JUN/26"
)

df_colab = df_colab.rename(
    columns={
        df_colab.columns[2]: "CHAVE",
        df_colab.columns[3]: "EQUIPE"
    }
) # renomei as colunas na planilha de colab para facilitar o merge (n salva a alteração na planilha original)

df_baixada = df_baixada.merge(
    df_colab[["CHAVE", "EQUIPE"]],
    on="CHAVE",
    how="left"
)

equipes_validas = ["Tratamento Coleta TV"]

fora = df_baixada[~df_baixada["EQUIPE"].isin(equipes_validas)]
fora = fora["Usuário"].drop_duplicates().tolist()

for colab in fora:
    indices = df_baixada[df_baixada["Usuário"] == colab].index

    if (df_baixada.loc[indices, "EQUIPE"] == "Tratamento").any():
        print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Colaborador '{colab}' é Assistente ou Líder e será removido.")
        df_baixada = df_baixada.drop(index=indices)
    else:
        decisao = input(
            f"|{Timestamp.now().strftime('%H:%M:%S')}| Colaborador com equipe inválida: '{colab}'\n"
            f"[1] manter\n"
            f"[2] remover\n"
        ).strip()

        if decisao == "2":
            print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Colaborador '{colab}' removido")
            df_baixada = df_baixada.drop(index=indices)
        else:
            print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Colaborador '{colab}' mantido.")

fora = df_baixada[~df_baixada["EQUIPE"].isin(equipes_validas)]
fora = fora["Usuário"].drop_duplicates().tolist()
print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Colaboradores mantidos fora da equipe válidas: {fora}")

# ORDENANDO COLUNAS
print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Ordenando dados")
df_ordenado = df_baixada.copy()

colunas_ordenadas = ["Motivo Fechamento Tela", "ID", "Data Início", "Usuário"]

for colunas in colunas_ordenadas:
    df_ordenado = df_ordenado.sort_values(by=colunas, ascending=True, kind="stable")

print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Abrindo planilha destino")
app = xw.App(visible=False, add_book=False)
wb = app.books.open(planilha_final_path)
ws = wb.sheets["Sheet"]

tabela_excel = ws.api.ListObjects(1)
headers = [cell.Value for cell in tabela_excel.HeaderRowRange]

colunas_presentes = [c for c in headers if c in df_ordenado.columns]

df_final = df_ordenado[colunas_presentes]

ultima_linha_corpo = tabela_excel.Range.Rows.Count + tabela_excel.HeaderRowRange.Row

# Colamos apenas os valores (index=False e header=False para não repetir o cabeçalho)
print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Inserindo dados via xlwings")
dados = [[""] * len(headers) for _ in range(len(df_ordenado))]

for col in colunas_presentes:
    idx = headers.index(col)

    for i, valor in enumerate(df_ordenado[col]):
        dados[i][idx] = valor

ws.range((ultima_linha_corpo, tabela_excel.Range.Column)).value = dados

wb.save()
wb.close()
app.quit()

atualizar_dados(planilha_final_path)

enviar_email(destinatarios_indicador, cc_indicador)

print(f"|{Timestamp.now().strftime('%H:%M:%S')}| Processo finalizado")
