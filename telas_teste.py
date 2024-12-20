def tamanho_aplicacao(self):
    try:
        # Obtém as dimensões do monitor
        monitor = QDesktopWidget().screenGeometry()
        monitor_width = monitor.width()
        monitor_height = monitor.height()
        print(monitor_width, monitor_height)

        # monitor_width = 1366
        # monitor_height = 768

        if monitor_width > 1200:
            print("entrei")

        # Centraliza a aplicação no monitor
        x = (monitor_width - monitor_width) // 2
        y = (monitor_height - monitor_height) // 2

        print(monitor_width, monitor_height)

        # Define as dimensões e posição
        self.setGeometry(x, y, monitor_width, monitor_height)

    except Exception as e:
        nome_funcao = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        trata_excecao(nome_funcao, str(e), nome_arquivo, exc_traceback)