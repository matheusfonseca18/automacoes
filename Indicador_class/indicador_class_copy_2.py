import pandas as pd
from pandas import Timestamp
import win32com.client as win32
import xlwings as xw
import time
import os
from dotenv import dotenv_values
from pathlib import Path
from shared.backup_utils import fazer_backup

def executar(logger, perguntar_colaborador):

    inicio_geral = time.perf_counter()

    # Carregar variáveis de ambiente
    BASE_DIR = Path(__file__).resolve().parent

    config = dotenv_values(BASE_DIR / ".env")

    # Caminhos
    base = Path(os.environ["OneDrive"])

    planilha_baixada_path = base / config["baixada_montar"]
    planilha_final_path = base / config["final_montar"]
    # planilha_final_path = config["final_local"]
    planilha_colab = base / config["colab_montar"]

    # Destinatários
    destinatarios_indicador = config["destinatarios_indicador"]
    cc_indicador = config["cc_indicador"]

    # funções
    def enviar_email(destinatario, cc):
        saudacao = "Bom dia, pessoal." if Timestamp.now().hour < 12 else "Boa tarde, pessoal."
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
        except Exception as e:
            logger.exception(f"Erro as criar instância do Outlook: {e}")
            raise

        mail.to = destinatario
        mail.CC = cc
        mail.Subject = "Indicador de Classificação - 2026"

        mail.HtmlBody = Rf"""<div style="font-family: tahoma; font-size: 11pt">
            <p>{saudacao}</p>
            <p>O indicador de classificação está atualizado até {Timestamp.now().strftime('%d/%m/%Y')}</p>
            <p>Arquivo: <a href="https://adgbl.sharepoint.com/:x:/r/sites/Relatriosgerenciais/Arquivos/2026/Analises/Classifica%C3%A7%C3%A3o/Indicador - classifica%C3%A7%C3%A3o 2026.xlsx?d=w760aea90c357485589274b4504d94c34&csf=1&web=1&e=8rjrfK&xsdata=MDV8MDJ8bWF0aGV1cy5waW50b0BpYm9wZS5jb218MDc3ZDczNDFhMGQ5NGJhM2FlM2QwOGRlOTRiM2Q0YmR8YjI3NjcyNDFmYWI1NDU0YjhiNjJmNjMyNDY1MGUzMTZ8MHwwfDYzOTExMTY5NzI1Nzk4MTI4OHxVbmtub3dufFRXRnBiR1pzYjNkOGV5SkZiWEIwZVUxaGNHa2lPblJ5ZFdVc0lsWWlPaUl3TGpBdU1EQXdNQ0lzSWxBaU9pSlhhVzR6TWlJc0lrRk9Jam9pVFdGcGJDSXNJbGRVSWpveWZRPT18MHx8fA%3d%3d&sdata=QlpMbitCVUtmelAwcUIxQU5rTXpXZjNETnpuWE9Vb3lhR3dGaHQwT2hCbz0%3d" target="_blank" rel="noopener noreferrer">Indicador - classificação 2026.xlsx</a></p>
            <p>Para acessar o arquivo, acesse via desktop.</p>
            <p style="font-family: tahoma; font-size: 9pt; color: #555;"><i>E-mail enviado automaticamente.</i></p>
        </div>
        """
        mail.Send() 
        logger.info(f"E-mail enviado\n"
                f"Destinatário: {destinatario}\n"
                f"CC: {cc}")

    logger.info("Processo iniciado")

    # backup
    fazer_backup(planilha_final_path, planilha_baixada_path, logger)

    # ler e proicessar dados
    logger.info("Tranformando dados")
    inicio = time.perf_counter()
    try:
        df_baixada = pd.read_excel(planilha_baixada_path, skiprows=1)
        logger.info("Arquivo Excel (planilha baixada) lido com sucesso")
    except Exception as e:
        logger.exception(f"Erro ao ler o arquivo Ecxel (planilha baixada): {e}")
        raise

    df_baixada = df_baixada.drop(columns=["Nome Tela", "Versão AG"])
    df_baixada = df_baixada.drop(index=df_baixada[df_baixada.iloc[:, 3].isna()].index) # apaga a soma no fim da planilha
    try:
        df_colab = pd.read_excel(planilha_colab, sheet_name="Planilha1")
        logger.info("Arquivo Excel (planilha colaboradore) lido com sucesso")
    except Exception as e:
        logger.exception(f"Erro ao ler o arquivo Ecxel (planilha colaboradores): {e}")
        raise

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

    try:
        df_colab = df_colab.rename(
            columns={
                df_colab.columns[2]: "CHAVE",
                df_colab.columns[3]: "EQUIPE"
            }
        ) # renomei as colunas na planilha de colab para facilitar o merge (n salva a alteração na planilha original)
    except Exception as e:
        logger.exception(f"Erro as renomear as colunas (planilha colaboradores: {e})")
        raise

    df_baixada = df_baixada.merge(
        df_colab[["CHAVE", "EQUIPE"]],
        on="CHAVE",
        how="left"
    )

    equipes_validas = ["Tratamento Coleta TV"]

    fora = df_baixada[~df_baixada["EQUIPE"].isin(equipes_validas)]
    fora = fora["Usuário"].drop_duplicates().tolist()

    logger.info(f"Leitura e tranformação planilha baixada: {time.perf_counter()-inicio:.2f}s")

    for colab in fora:
        indices = df_baixada[df_baixada["Usuário"] == colab].index

        if (df_baixada.loc[indices, "EQUIPE"] == "Tratamento").any():
            logger.info(f"Colaborador '{colab}' é Assistente ou Líder e será removido.")
            df_baixada = df_baixada.drop(index=indices)

        else:
            decisao = perguntar_colaborador(colab)

            if decisao:
                logger.info(f"Colaborador '{colab}' mantido.")
            else:
                df_baixada = df_baixada.drop(index=indices)
                logger.info(f"Colaborador '{colab}' removido")

    fora = df_baixada[~df_baixada["EQUIPE"].isin(equipes_validas)]
    fora = fora["Usuário"].drop_duplicates().tolist()
    if fora:
        logger.info(f"Colaboradores mantidos fora da equipe válida: {fora}")
    else:
        logger.info("Nenhum colaborador mantido fora da equipe válida")

    # ORDENANDO COLUNAS
    logger.info(f"Ordenando dados")
    inicio = time.perf_counter()
    df_ordenado = df_baixada.copy()

    colunas_ordenadas = ["Motivo Fechamento Tela", "ID", "Data Início", "Usuário"]

    for colunas in colunas_ordenadas:
        df_ordenado = df_ordenado.sort_values(by=colunas, ascending=True, kind="stable")

    logger.info(f"Colunas ordenadas: {time.perf_counter()-inicio:.2f}s")

    logger.info(f"Abrindo planilha destino")
    inicio = time.perf_counter()
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(planilha_final_path)
        app.screen_updating = False
        app.display_alerts = False

        app.api.EnableEvents = False
        app.api.Calculation = -4135  # Manual
        ws = wb.sheets["Sheet"]
    except Exception as e:
        logger.exception(f"Erro ao abrir planilha destino com xlwings: {e}")
        raise

    logger.info(f"Abrindo com xlwings: {time.perf_counter()-inicio:.2f}s")

    inicio = time.perf_counter()

    tabela_excel = ws.api.ListObjects(1)
    headers = [cell.Value for cell in tabela_excel.HeaderRowRange]

    colunas_presentes = [c for c in headers if c in df_ordenado.columns]

    df_final = df_ordenado[colunas_presentes]

    ultima_linha_corpo = tabela_excel.Range.Rows.Count + tabela_excel.HeaderRowRange.Row

    logger.info(f"Identificando tabela e conlunas: {time.perf_counter()-inicio:.2f}s")

    # Colamos apenas os valores (index=False e header=False para não repetir o cabeçalho)
    logger.info("Inserindo dados via xlwings")

    inicio = time.perf_counter()

    dados = [[""] * len(headers) for _ in range(len(df_ordenado))]

    for col in colunas_presentes:
        idx = headers.index(col)

        for i, valor in enumerate(df_ordenado[col]):
            dados[i][idx] = valor

    ws.range((ultima_linha_corpo, tabela_excel.Range.Column)).value = dados

    logger.info(f"Inserir dados: {time.perf_counter()-inicio:.2f}s")

    try:
        inicio = time.perf_counter()
        app.api.Calculation = -4105  # Automático
        app.api.Calculate()

        inicio = time.perf_counter()
        logger.info("Iniciando atualização")
        wb_api = wb.api

        wb_api.RefreshAll()
        app.api.CalculateUntilAsyncQueriesDone()

        logger.info(f"Atualização finalizada: {time.perf_counter()-inicio:.2f}s")

    except Exception as e:
        logger.exception(f"Erro ao atualizar a planilha final com xlwings: {e}")
        raise

    try:
        inicio = time.perf_counter()
        logger.info("Salvando planilha")

        wb.save()
        logger.info(f"Planilha salva: {time.perf_counter()-inicio:.2f}s")

    except Exception as e:
        logger.exception(f"Erro ao salvar a planilha final com xlwings: {e}")
        raise
    
    finally:
        try:
            app.api.EnableEvents = True
        except:
            pass
        inicio = time.perf_counter()
        wb.close()
        app.quit()
        logger.info(f"Fechar: {time.perf_counter()-inicio:.2f}s")

    enviar_email(destinatarios_indicador, cc_indicador)

    logger.info(f"Processo finalizado: {time.perf_counter()-inicio_geral:.2f}s")
