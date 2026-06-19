import logging
from logging.handlers import TimedRotatingFileHandler

def get_logger(nome_arquivo, pasta_logs ):
    logger = logging.getLogger(nome_arquivo)

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    logger.propagate = False

    log_dir = pasta_logs / "logs"
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / nome_arquivo

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(filename)s:%(lineno)d | %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S"
    )

    file_handler = TimedRotatingFileHandler(
        log_file,
        when="midnight",
        interval=1,
        backupCount=30,
        encoding="utf-8"
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
