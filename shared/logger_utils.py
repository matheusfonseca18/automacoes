import logging
from logging.handlers import TimedRotatingFileHandler

def get_logger(nome_arquivo, pasta_logs, extra_handlers=None):
    logger = logging.getLogger(nome_arquivo)

    formatter = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(filename)s:%(lineno)d | %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S"
    )

    if not logger.handlers:

        logger.setLevel(logging.INFO)
        logger.propagate = False

        log_dir = pasta_logs / "logs"
        log_dir.mkdir(exist_ok=True)

        log_file = log_dir / f"{nome_arquivo}.log"

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

    if extra_handlers:
        for handler in extra_handlers:
            handler.setFormatter(formatter)

            if not any(isinstance(h, type(handler))for h in logger.handlers):
                logger.addHandler(handler)

    return logger
class TextboxHandler(logging.Handler):
    def __init__(self, textbox):
        super().__init__()
        self.textbox = textbox

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.textbox.insert("end", msg + "\n")
            self.textbox.see("end")

        self.textbox.after(0, append)
        