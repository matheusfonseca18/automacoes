# import pandas as pd
from dotenv import load_dotenv
import os
from pathlib import Path
from shared.logger_utils import get_logger
from shared.backup_utils import fazer_backup

load_dotenv()
planilha_baixada_path = os.getenv("planilha_baixada_path")
planilha_final_path = os.getenv("planilha_final_path")

logger = get_logger(
    nome_arquivo="teste.log",
    pasta_logs=Path(__file__).parent
)

logger.info("teste de log")

fazer_backup(planilha_final_path, planilha_baixada_path, logger, manter=30)
