def executar_automacao(nome_arquivo, pasta_logs, automacao):

    alterar_estado_botao(False)

    def tarefa():
        try:
            pythoncom.CoInitialize()
            logger = get_logger(
                nome_arquivo=nome_arquivo,
                pasta_logs=BASE_DIR / pasta_logs,
                extra_handlers=[textbox_handler]
            )

            limpar_textbox(textbox_log)
            automacao(logger)
        
        finally:
            pythoncom.CoUninitialize()
            root.after(0, lambda: alterar_estado_botao(True))

    Thread(target=tarefa, daemon=True).start()


def executar_automacao(nome_arquivo, pasta_logs, automacao):

    alterar_estado_botao(False)

    def tarefa():
        try:
            pythoncom.CoInitialize()

            logger = get_logger(
                nome_arquivo=nome_arquivo,
                pasta_logs=pasta_logs,
                extra_handlers=[textbox_handler]
            )

            limpar_textbox(textbox_log)
            automacao(logger)

        finally:
            pythoncom.CoUninitialize()
            root.after(0, lambda: alterar_estado_botao(True))

    Thread(target=tarefa, daemon=True).start()