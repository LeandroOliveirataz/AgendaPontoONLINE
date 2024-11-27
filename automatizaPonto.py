import tkinter as tk
from tkinter import simpledialog, messagebox
import pyautogui
import time
import schedule
import os
from datetime import datetime, timedelta
import pygetwindow as gw
import requests
import threading

class PontoRegistrarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("-== AGENDA PONTO SUAP ==-")

        # Caminho para a área de trabalho
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.log_file_path = os.path.join(desktop_path, "Registro Ponto.txt")
        
        # Criação de widgets
        self.label = tk.Label(root, text="### Registrar Ponto SUAP ###\n### Versão 1.3 ###")
        self.label.pack()
        
        self.log_text = tk.Text(root, height=10, width=79)
        self.log_text.pack()
        
        self.start_button = tk.Button(root, text="Iniciar Agendamento", command=self.start_task)
        self.start_button.pack()
        
        self.quit_button = tk.Button(root, text="Cancelar", command=self.stop_task)
        self.quit_button.pack()
        
        self.continuar_verificacao = True
        self.continuar_verificacao_internet = True
        self.task_thread = None
        self.start_time = None
        self.max_wait_time = 3600  # Tempo máximo de espera em segundos (60 minutos)

    def log_message(self, message):
        # Exibe a mensagem no widget de log
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        
        # Adiciona a mensagem ao arquivo de log
        with open(self.log_file_path, 'a') as log_file:
            log_file.write(message + "\n")

    def check_internet(self):
        try:
            response = requests.get("https://www.google.com", timeout=5)
            return response.status_code == 200
        except requests.ConnectionError:
            return False

    def wait_for_internet(self):
        if not self.start_time:
            self.start_time = datetime.now()
        
        elapsed_time = (datetime.now() - self.start_time).total_seconds()
        if elapsed_time >= self.max_wait_time:
            self.log_message("Não foi possível estabelecer conexão com a internet dentro de 60 minutos. Cancelando a execução.")
            self.stop_task()
            return
        
        if not self.check_internet():
            self.log_message(f"Sem conexão com a internet às {datetime.now()}. Tentando novamente...")
            self.root.after(10000, self.wait_for_internet)  # Tentar novamente após 10 segundos
        else:
            self.log_message(f'###########################################')
            self.log_message("Teste de conexão com a internet: CONECTADO.")
            self.log_message(f'###########################################')
            self.execute_task_if_due()

    def start_task(self):
        try:
            # Coletar inputs do usuário
            self.usuario = simpledialog.askstring("Input", "Insira seu SUAP de usuário:")
            self.senha = simpledialog.askstring("Input", "Insira sua senha:", show='*')
            self.hora_execucao = simpledialog.askstring("Input", "Insira o horário para executar o registro (HH:MM:SS):")

            if not self.usuario or not self.senha or not self.hora_execucao:
                raise ValueError("Todos os campos devem ser preenchidos.")

            # Verificar se o horário está no formato correto
            time.strptime(self.hora_execucao, '%H:%M:%S')

            # Verificar se o horário programado é menor que o horário atual
            now = datetime.now().strftime('%H:%M:%S')
            if self.hora_execucao < now:
                # Ajustar o horário para o dia seguinte
                hoje = datetime.now()
                data_futura = hoje + timedelta(days=1)
                nova_data_hora = data_futura.strftime('%Y-%m-%d') + ' ' + self.hora_execucao
                # Solicitar confirmação do usuário
                resposta = messagebox.askyesno("Confirmar Agendamento", f"O horário agendado ({self.hora_execucao}) é menor que o horário atual. A tarefa será agendada para o dia seguinte ({nova_data_hora}). Deseja confirmar?")
                if resposta:
                    self.hora_execucao = nova_data_hora
                    self.log_message(f"O horário foi ajustado para o dia seguinte: {self.hora_execucao}")
                else:
                    self.log_message("Agendamento cancelado pelo usuário.")
                    return

            else:
                # Se o horário não é menor que o horário atual, apenas ajuste para o formato correto
                self.hora_execucao = datetime.now().strftime('%Y-%m-%d') + ' ' + self.hora_execucao

            # Perguntar ao usuário se deseja desligar o computador
            self.desligar_computador = messagebox.askyesno("Desligar Computador", "Deseja DESLIGAR o computador após o registro da frequência?")

            # Logar informações fornecidas
            self.log_message(f'###########################################')
            self.log_message(f'Login de SUAP fornecido: {self.usuario}')
            self.log_message(f'Senha de usuário fornecida: {"*"*len(self.senha)}')
            self.log_message(f'A tarefa será iniciada às: {self.hora_execucao}')
            self.log_message(f'Desligar computador após a execução: {"SIM" if self.desligar_computador else "NÃO"}')
            self.log_message(f'Acesso a Internet: {"ONLINE" if self.check_internet() else "OFFLINE"}')
            self.log_message(f'###########################################')

            # Iniciar a verificação da conexão com a internet
            self.log_message("Iniciando a verificação da conexão com a internet...")
            self.start_time = datetime.now()
            self.wait_for_internet()
            
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

    def execute_task_if_due(self):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if now >= self.hora_execucao:
            self.log_message(f"Horário agendado era ({self.hora_execucao}). Executando a tarefa imediatamente.")
            self.start_background_task()
        else:
            self.start_scheduled_task()

    def start_scheduled_task(self):
        # Divida a data e a hora para agendamento
        hora = self.hora_execucao.split(' ')[1]
        schedule.every().day.at(hora).do(self.verify_and_execute_task)
        self.log_message("Aguardando o INÍCIO DE EXECUÇÃO DA TAREFA...")
        self.check_schedule()

    def check_schedule(self):
        if self.continuar_verificacao:
            schedule.run_pending()
            self.log_message(f"Verificação em ... {datetime.now()}, AGUARDANDO HORÁRIO.")
            self.root.after(20000, self.check_schedule)  # Verificar a cada 20 segundos

    def start_background_task(self):
        if self.task_thread and self.task_thread.is_alive():
            self.log_message("Tarefa já está em execução.")
            return
        self.task_thread = threading.Thread(target=self.verify_and_execute_task)
        self.task_thread.daemon = True
        self.task_thread.start()

    def verify_and_execute_task(self):
        self.log_message("Verificando a conexão com a internet antes de iniciar a tarefa...")
        self.start_time = datetime.now()
        self.root.after(10000, self.check_internet_before_task)

    def check_internet_before_task(self):
        if self.check_internet():
            self.log_message("Conexão com a internet estabelecida. Iniciando a tarefa...")
            self.job()
        else:
            elapsed_time = (datetime.now() - self.start_time).total_seconds()
            if elapsed_time >= self.max_wait_time:
                self.log_message("Não foi possível restabelecer a conexão com a internet dentro de 60 minutos. Cancelando a execução.")
                self.stop_task()
            else:
                self.log_message("Sem conexão com a internet. Tentando novamente...")
                self.root.after(10000, self.check_internet_before_task)

    def job(self):
        self.log_message(f'###########################################')
        self.log_message(f"Execução de TAREFA iniciada em {datetime.now()}...")
        
        try:
            # Passo 1 - Entrar no sistema da empresa
            pyautogui.PAUSE = 0.5  # Pausa de 0,5 segundos
            self.abrir_chrome()
            self.acessar_suap()
            self.fazer_login()
            self.registrar_frequencia()

            self.log_message(f'###########################################')
            self.log_message("REGISTRO DE FREQUÊNCIA REALIZADO COM SUCESSO.")
            self.log_message(f"TAREFA FINALIZADA ÀS {datetime.now()}.")
            self.log_message(f'###########################################')

            # Desligar o computador se o usuário escolheu
            if self.desligar_computador:
                os.system("shutdown /s /t 1")  # Desligar o computador
            
        except Exception as e:
            self.log_message(f"Erro durante a execução: {str(e)}")
        
        # Definir o flag como False para interromper o loop
        self.continuar_verificacao = False
        # Cancelar o agendamento
        schedule.clear()
        # Fechar a aplicação
        self.stop_task()
        return schedule.CancelJob

    def abrir_chrome(self):
        try:
            pyautogui.press("win")
            pyautogui.write("chrome")
            pyautogui.press("enter")
            time.sleep(2)  # Espera para garantir que o Chrome esteja aberto
            chrome_windows = gw.getWindowsWithTitle("Chrome")
            if chrome_windows:
                chrome_window = chrome_windows[0]
                chrome_window.maximize()
        except Exception as e:
            self.log_message(f"Erro ao abrir o Chrome: {str(e)}")
            raise

    def acessar_suap(self):
        try:
            pyautogui.hotkey('ctrl', 'l')  # Foca na barra de endereços do navegador
            pyautogui.write("https://suap.ifba.edu.br/ponto/registrar/")
            pyautogui.press("enter")
            time.sleep(5)  # Espera de 5 segundos
        except Exception as e:
            self.log_message(f"Erro ao acessar o SUAP: {str(e)}")
            raise

    def fazer_login(self):
        try:
            for _ in range(5):  # Ajuste o número conforme necessário
                pyautogui.press("tab")
            pyautogui.write(self.usuario)  # Usar o nome de usuário fornecido
            pyautogui.press("tab")
            pyautogui.write(self.senha)  # Usar a senha fornecida
            pyautogui.press("enter")
            time.sleep(10)  # Espera de 10 segundos
        except Exception as e:
            self.log_message(f"Erro ao fazer login: {str(e)}")
            raise

    def registrar_frequencia(self):
        try:
            # Implementar a lógica para registrar a frequência
            pass  # Placeholder
        except Exception as e:
            self.log_message(f"Erro ao registrar frequência: {str(e)}")
            raise

    def stop_task(self):
        self.continuar_verificacao = False
        self.continuar_verificacao_internet = False
        self.log_message(f'###########################################')
        self.log_message(">>>>>     Agendamento Finalizado.     <<<<<")
        self.root.after(10000, self.root.quit)  # Esperar 10 segundos antes de fechar a aplicação

root = tk.Tk()
app = PontoRegistrarApp(root)
root.mainloop()
#criar executável
#pyinstaller --onefile -w automatiz4.py

