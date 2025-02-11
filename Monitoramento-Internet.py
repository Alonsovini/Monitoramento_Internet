# Monitoramento Internet Escritório

import sys
import time
from datetime import datetime
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon
import schedule
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import win32com.client as win32


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.next_run = None
        self.update_next_run()

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)
        self.timer.start(1000)  # Update every second

    def initUI(self):
        self.setWindowTitle("Monitoramento de Internet")
        self.setGeometry(100, 100, 500, 300)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        # Define o ícone da janela
        self.setWindowIcon(QIcon('Internet.ico'))

        self.label_status = QLabel("Aguardando Próximo Evento...", self)
        self.label_status.setStyleSheet("font-size: 18px;")  # Set font size to 12
        self.layout.addWidget(self.label_status)

        self.label_timer = QLabel("Tempo Restante Até O Próximo Teste:", self)
        self.label_timer.setStyleSheet("font-size: 18px;")  # Set font size to 12
        self.layout.addWidget(self.label_timer)

        # Barra de progresso para indicar o monitoramento
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setMaximum(0)  # Indeterminado
        self.progress_bar.setMinimum(0)
        self.progress_bar.setVisible(True)  # Torna visível a barra de progresso
        self.layout.addWidget(self.progress_bar)

        self.central_widget.setLayout(self.layout)

    def update_timer(self):
        if self.next_run:
            now = datetime.now()
            remaining_time = self.next_run - now
            if remaining_time.total_seconds() > 0:
                self.label_timer.setText(f"Tempo Restante Até O Próximo Teste: {str(remaining_time).split('.')[0]}")
            else:
                self.update_next_run()
                executar_teste()

    def update_next_run(self):
        schedule.run_pending()
        self.next_run = self.get_next_run_time()
        if self.next_run:
            self.total_time = (self.next_run - datetime.now()).total_seconds()

    def get_next_run_time(self):
        next_runs = [job.next_run for job in schedule.jobs]
        if next_runs:
            return min(next_runs)
        return None


def executar_teste():
    # Defina o caminho onde os prints serão salvos
    diretorio_screenshots = "C:\\Prints"  # Altere este caminho para onde você deseja salvar as imagens

    # Crie o diretório se não existir
    if not os.path.exists(diretorio_screenshots):
        os.makedirs(diretorio_screenshots)

    # Configuração do Selenium
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")  # Executa o Chrome em segundo plano (opcional)
    options.add_argument("--window-size=1920,1080")  # Define o tamanho da janela
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()

    # Site e botões
    site_url = "https://www.brasilbandalarga.com.br/"
    botao_aceitar_cookies_class = "cc-btn cc-dismiss"
    botao_id = "btnIniciar"

    # Acessar o site
    driver.get(site_url)

    # Aceitar cookies (se o botão existir) - Clicando com JavaScript
    try:
        cookie_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, botao_aceitar_cookies_class))
        )
        driver.execute_script("arguments[0].click();", cookie_button)
        time.sleep(2)  # Pequeno atraso para garantir que o banner desapareça
    except:
        pass  # Ignorar se o botão não existir

    # Rolar a página para baixo
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Aguardar o botão e clicar - Clicando com JavaScript
    botao = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.ID, botao_id))
    )
    driver.execute_script("arguments[0].click();", botao)

    # Aguardar o término do teste (ajuste o tempo conforme necessário)
    time.sleep(80)

    # Tentar encontrar o elemento por um tempo máximo
    max_wait_time = 30  # Tempo máximo de espera em segundos
    start_time = time.time()
    resultado_encontrado = False

    while time.time() - start_time < max_wait_time:
        try:
            resultado_elemento = driver.find_element(By.XPATH, "//*[@id='medicao']/div")
            resultado_encontrado = True
            break  # Sai do loop se o elemento for encontrado
        except:
            time.sleep(1)  # Espera 1 segundo antes de tentar novamente

    # Tirar screenshot e enviar e-mail
    if resultado_encontrado:
        try:
            # Rolar a página para garantir que todos os elementos estejam visíveis
            driver.execute_script("arguments[0].scrollIntoView(true);", resultado_elemento)
            time.sleep(2)  # Aguardar um tempo extra para garantir a renderização

            # Gerar nome de arquivo com data e hora
            agora = datetime.now()
            nome_arquivo = f"resultado_{agora.strftime('%d-%m-%Y_%H-%M')}.png"
            caminho_screenshot = os.path.join(diretorio_screenshots, nome_arquivo)

            resultado_elemento.screenshot(caminho_screenshot)
            print(f"Print salvo em: {caminho_screenshot}")

            # Verifica se o arquivo existe
            if not os.path.exists(caminho_screenshot):
                raise Exception("O arquivo de print não foi encontrado")

            # Enviar e-mail com o resultado (anexo com a imagem)
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = "suporte@postos.com.br"
            mail.Subject = "Teste de Velocidade"
            mail.Body = ("Bom dia! \n\nSegue o resultado do teste de velocidade em anexo. \n\nAtt.")
            mail.Attachments.Add(caminho_screenshot)

            mail.Save()
            mail.Send()

            # Forçar o envio de todos os e-mails na caixa de saída
            namespace = outlook.GetNamespace("MAPI")
            outbox = namespace.GetDefaultFolder(4)  # 4 é a caixa de saída
            for item in outbox.Items:
                if item.Sent == False:
                    item.Send()

        except Exception as e:
            print(f"Erro ao tirar o print: {e}")

    else:
        # Enviar e-mail indicando que o resultado não foi encontrado
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "suporte@postos.com.br"
        mail.Subject = "Resultado do Teste de Velocidade (Erro)"
        mail.Body = "O resultado do teste de velocidade não foi encontrado."
        mail.Send()

    driver.quit()


# Agendar as execuções
schedule.every().monday.at("08:00").do(executar_teste)
schedule.every().monday.at("16:30").do(executar_teste)
schedule.every().tuesday.at("08:00").do(executar_teste)
schedule.every().tuesday.at("16:30").do(executar_teste)
schedule.every().wednesday.at("08:00").do(executar_teste)
schedule.every().wednesday.at("16:30").do(executar_teste)
schedule.every().thursday.at("08:00").do(executar_teste)
schedule.every().thursday.at("16:30").do(executar_teste)
schedule.every().friday.at("08:00").do(executar_teste)
schedule.every().friday.at("15:42").do(executar_teste)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())



