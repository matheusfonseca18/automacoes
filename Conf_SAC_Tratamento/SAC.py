import pandas as pd
from pandas import Timestamp, Timedelta
import xlwings as xw
import win32com.client as win32
from dotenv import load_dotenv
import os

print(f'Processo iniciado as {Timestamp.now().strftime("%H:%M:%S")}')

load_dotenv()
destinatarios_equipe = os.getenv("destinatarios_equipe")
cc_equipe = os.getenv("cc_equipe")
destinatarios_ocorrencias = os.getenv("destinatarios_ocorrencias")
cc_ocorrencias = os.getenv("cc_ocorrencias")
destinatarios_indicador = os.getenv("destinatarios_indicador")
cc_indicador = os.getenv("cc_indicador")

# Caminhos
planilha_baixada_path = os.getenv("planilha_baixada_path")
planilha_final_path = os.getenv("planilha_final_path")

def enviar_email(destinatario, cc, assunto, texto, acao, corpo=None):
    saudacao = "Bom dia, pessoal." if Timestamp.now().hour < 12 else "Boa tarde, pessoal."

    if corpo is None:
        html_tabela = ""
    elif isinstance(corpo, pd.DataFrame):
        html_tabela = corpo.fillna('').to_html(
            index=False, 
            border=1, 
            justify='center'
        )
    else:
        html_tabela = f"<p>{corpo}</p>"

    # Conecta Outlook aberto
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Display() 
    # Captura a assinatura que o Outlook
    assinatura = mail.HTMLBody

    mail.To = destinatario
    mail.CC = cc
    mail.Subject = assunto

    mail.HTMLBody = f"""<div style="font-family: tahoma; font-size: 11pt">
        <p>{saudacao}</p>
        <p>{texto}</p>
        {html_tabela}
        <p style="font-family: tahoma; font-size: 9pt; color: #555;"><i>E-mail enviado automaticamente.</i></p>
        {assinatura}
    </div>
    """

    # mail.Display() # pra ver o e-mail antes de enviar
    mail.Send()

    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| E-mail de {acao} enviado com sucesso!")

def atualizar_dados(caminho):
    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Iniciando atualização de dados no Excel...")
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

# Ler e processar CSV
df_csv = pd.read_csv(planilha_baixada_path, encoding='latin1', sep=';')
df_csv.columns = df_csv.columns.str.strip()

df_filtro = df_csv[df_csv["Ocorrência"].isnull()]

if df_filtro.empty:
    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Nenhuma ocorrência em branco encontrada. E-mail não será enviado.")
else: # E-mail ocorrencias
     enviar_email(destinatarios_ocorrencias , cc_ocorrencias, "Confirmação SAC - Ocorrências em branco", "Seguem as confirmações SAC com ocorrências em branco:", "Ocorrências em branco", df_filtro)

df_csv = df_csv.dropna(subset=["Ocorrência"]) # apagar ocorrencias em branco
print(f"|{Timestamp.now().strftime("%H:%M:%S")}| ocorrências em branco removidas")

print(f"|{Timestamp.now().strftime("%H:%M:%S")}| abrindo excel com xlwings")
# Abrir o Excel xlwings
# visible=False faz tudo em background, add_book=False evita abrir uma planilha em branco
app = xw.App(visible=False, add_book=False)
wb = app.books.open(planilha_final_path)
ws = wb.sheets["BD Confirmação SAC - tratamento"]

# Obter estrutura da Tabela (ListObject no Excel)
# xlwings acessa a tabela diretamente pelo nome ou índice
tabela_excel = ws.api.ListObjects(1) # Pega a primeira tabela da aba
headers = [cell.Value for cell in tabela_excel.HeaderRowRange]

# Filtrar colunas e converter datas
colunas_presentes = [c for c in headers if c in df_csv.columns]
df_csv = df_csv[colunas_presentes]

meses_pt = {'jan':'01','fev':'02','mar':'03','abr':'04','mai':'05','jun':'06',
            'jul':'07','ago':'08','set':'09','out':'10','nov':'11','dez':'12'}

for col_name in df_csv.columns:
    # Verificação para identificar colunas de data pelos nomes conhecidos (testar se é necessário)
    if "data" in col_name.lower() or "mês" in col_name.lower() or "veiculação" in col_name.lower():
        if col_name == "Mês Veiculação":
            for sigla, num in meses_pt.items():
                df_csv[col_name] = df_csv[col_name].astype(str).str.replace(sigla, f"01/{num}", case=False, regex=False)
        
        df_csv[col_name] = pd.to_datetime(df_csv[col_name], errors='coerce', dayfirst=True, format='mixed')

# Inserir dados no final da tabela
print(f"|{Timestamp.now().strftime("%H:%M:%S")}| inserindo dados via xlwings")
ultima_linha_corpo = tabela_excel.Range.Rows.Count + tabela_excel.HeaderRowRange.Row

# Colamos apenas os valores (index=False e header=False para não repetir o cabeçalho)
ws.range(f"A{ultima_linha_corpo}").options(pd.DataFrame, index=False, header=False).value = df_csv

# Salvar e Fechar
print(f"|{Timestamp.now().strftime("%H:%M:%S")}| salvando e fechando planilha")
wb.save()
wb.close()
app.quit()

print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Dados inseridos via xlwings com sucesso!")

df_final = pd.read_excel(planilha_final_path, sheet_name="BD Confirmação SAC - tratamento")

df_anunciante = df_final[(df_final['Anunciante AG'].isnull()) | (df_final['Anunciante AG'] == "#N/D")]
df_equipe = df_final[(df_final['Equipe'].isnull()) | (df_final['Equipe'] == "#N/D")]

equipe_erro = "<br>".join(df_equipe['Usuário'].drop_duplicates().tolist())
anunciante_erro = "<br>".join(df_anunciante['Anunciante'].drop_duplicates().tolist())

if not df_equipe.empty: # E-mail equipe
    enviar_email(destinatarios_equipe, cc_equipe, "Equipe não classificada - Confirmação SAC", "Por gentileza, poderiam classificar os colaboradores abaixo?", "equipe não classificada", equipe_erro)
else:
    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Nenhuma linha com Equipe em branco ou #N/D encontrada.")

if not df_anunciante.empty: # E-mail anunciante
    enviar_email(destinatarios_ocorrencias, cc_ocorrencias, "Anunciantes não classificados - Confirmação SAC", "Por gentileza, poderiam classificar os anunciantes abaixo, conforme o padrão que temos no AG?", "Anunciante não classificado", anunciante_erro)
else:
    print(f"|{Timestamp.now().strftime("%H:%M:%S")}| Nenhuma linha com Anunciante AG em branco ou #N/D encontrada.")
atualizar_dados(planilha_final_path)

try:
    df_sla = pd.read_excel(planilha_final_path, sheet_name="Area").loc[2:9, "Unnamed: 201":f"Unnamed: {201+int(Timestamp.now().strftime('%m'))+1}"].fillna("")
except Exception as e:
    print(f"Erro ao ler a planilha com o mês atual: {e}")
    df_sla = pd.read_excel(planilha_final_path, sheet_name="Area").loc[2:9, "Unnamed: 201":f"Unnamed: {201+int(Timestamp.now().strftime('%m'))}"].fillna("")

hoje = Timestamp.now()
domingo = (hoje - Timedelta(days=(hoje.weekday() + 1) % 7)).strftime('%d/%m/%Y')

# e-mail SLA
enviar_email(destinatarios_indicador, cc_indicador, "SLA - Confirmação SAC", f"Segue quadro de SLA confirmação atualizado até {domingo}", "SLA", df_sla.to_html(header=False, index=False))

# e-mail indicador
enviar_email(destinatarios_indicador, cc_indicador, "Indicador confirmação SAC - Tratamento", rf"O indicador de confirmação SAC está atualizado até {domingo}.<br><br><a href='https://adgbl.sharepoint.com/:x:/s/Relatriosgerenciais/IQAix4IxFd9oRJWoMkCYHeJlAYaJ_NBHG2V2hAQJuPERPW8?e=flab6c&xsdata=MDV8MDJ8bWF0aGV1cy5waW50b0BpYm9wZS5jb218ZmE3YmEyZTM4YWVmNGU5ZTU5MmIwOGRlOTU3NjFhYWR8YjI3NjcyNDFmYWI1NDU0YjhiNjJmNjMyNDY1MGUzMTZ8MHwwfDYzOTExMjUzMTYzODU5NjgyNHxVbmtub3dufFRXRnBiR1pzYjNkOGV5SkZiWEIwZVUxaGNHa2lPblJ5ZFdVc0lsWWlPaUl3TGpBdU1EQXdNQ0lzSWxBaU9pSlhhVzR6TWlJc0lrRk9Jam9pVFdGcGJDSXNJbGRVSWpveWZRPT18MHx8fA%3d%3d&sdata=ZDBWdElnTzhaZStMajMyTVozTFR0K2hBdEU3eHBwd2ZqZ3h6NWZmaFUvWT0%3d'>  Confirmação SAC - tratamento 2026.xlsx</a>", None)

print(f'Processo finalizado as {Timestamp.now().strftime("%H:%M:%S")}')
