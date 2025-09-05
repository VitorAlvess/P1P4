from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
from time import sleep 
import json
import os
from datetime import date


# imports google sheet
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class autentique:
    def iniciar(self):
        self.senha()
        self.configuracoes_selenium()
        self.autentique_login()
        self.configuracoes_sheets()
        self.sheets()
        

    def senha(self):
        with open("login_autentique.json", encoding='utf-8') as meu_json: # Importar dados de um arquivo config com usernames e passwords
            dados = json.load(meu_json)
            self.email = dados["email"]
            self.senha = dados["password"]
            print('Dados do settings.json lidos com sucesso!')

    def configuracoes_selenium(self):

        options = webdriver.ChromeOptions()

        # options.add_argument('--disable-gpu')
        # options.add_argument('--headless')
        options.add_argument('--window-size=1366,768')
        # options.add_argument('--no-sandbox')
        # options.add_argument('--start-maximized')
        # options.add_argument('--disable-setuid-sandbox')


        preferences = {"download.default_directory" : f'{os.getcwd()}\\relatorios', 'profile.default_content_setting_values.automatic_downloads': 1}
        options.add_experimental_option("prefs", preferences)
        servico = Service(ChromeDriverManager().install())
        self.navegar = webdriver.Chrome(service=servico, chrome_options=options)
    def autentique_login(self):
        pagina = self.navegar
        pagina.get('https://painel.autentique.com.br/entrar')
        sleep(5)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/label[1]/input' ).send_keys(self.email)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/label[2]/input' ).send_keys(self.senha)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/div/button' ).click()
        sleep(5)
        print('chegou até aqui')

    
    def configuracoes_sheets(self):
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = None
       
        if os.path.exists('token1.json'): #necessario pegar esse token la da api do google sheets
           self.creds = Credentials.from_authorized_user_file('token1.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token1.json', 'w') as token:
                token.write(self.creds.to_json())
        
    def sheets(self):
         # The ID and range of a sample spreadsheet.
        SAMPLE_SPREADSHEET_ID = '1eF3fhE-99ejF0K1fB7nB2WoLghAy9tVlYgseg_bvtTI'
        
        SAMPLE_RANGE_NAME = 'teste!A:Y'
       
        try:
            service = build('sheets', 'v4', credentials=self.creds)
            #Passar valores existentes para a variavel valore
            sheet = service.spreadsheets()


            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                            range=SAMPLE_RANGE_NAME).execute()
            valores = result['values']
            print(result)
         
           #Informar que o formulario ja foi gerado 
            


            # arrays_nao_vazias = list(filter(lambda x: len(x) > 0, valores[1:]))
            # arrays_nao_vazias = list(filter(lambda x: len(x) > 0, valores[1:]))
            
          


            for lista in valores:
                if lista[0] != '': 
                    print(lista)
                    if 'Carregando…' not in lista and '#N/A' not in lista and lista[-1] != 'TRUE':
                        print(lista)
                        posicao = result['values'].index(lista)
                        print(posicao + 1)
                        lista.pop()
                        self.autentique_dados_voluntarios(lista)
                        result2 = service.spreadsheets().values().update(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                        range=f'teste!Y{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [['TRUE']]
                        }
                        ).execute()

                        

                        def obter_data_atual():
                            data_atual = datetime.now()
                            data_formatada = data_atual.strftime("%d/%m/%Y")
                            return data_formatada

                        data_atual = obter_data_atual()
                        print(data_atual)
                        
        







                        result4 = service.spreadsheets().values().update(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                        range=f'teste!B{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [[f'{data_atual}']]
                        }
                        ).execute()
                        
       
        except HttpError as err:
                print(err)

    def autentique_dados_voluntarios(self, lista):
        print('Pagina do autentique')
        pagina = self.navegar
        sleep(30)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-home/app-documents-list/div[1]/app-taxonomy/aside/app-documents-taxonomy/div[1]/button').click()
        
        sleep(4)
        pagina.execute_script("window.scrollTo(0, 300);")
       
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/div[4]/div[2]/div/div[4]/a[2]/div[1]').click()
        sleep(5)








        pagina = self.navegar
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[1]/label/input').send_keys(lista[3])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[2]/label/textarea').send_keys(lista[4])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[3]/label/input').send_keys(lista[5])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[4]/label/input').send_keys(lista[6])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[5]/label/input').send_keys(lista[7])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[6]/label/input').send_keys(lista[8])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[7]/label/input').send_keys(lista[9])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[8]/label/input').send_keys(lista[10])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[9]/label/input').send_keys(lista[11])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[10]/label/input').send_keys(lista[12])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[11]/label/input').send_keys(lista[13])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[12]/label/input').send_keys(lista[14])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[13]/label/input').send_keys(lista[15])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[14]/label/input').send_keys(lista[16])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[15]/label/input').send_keys(lista[17])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[16]/label/input').send_keys(lista[18])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[17]/label/input').send_keys(lista[19])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[18]/label/input').send_keys(lista[20])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[19]/label/input').send_keys(lista[21])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[20]/label/input').send_keys(lista[22])
        data_atual = date.today()
        data_em_texto = data_atual.strftime('%d/%m/%Y')
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[21]/label/input').send_keys(data_em_texto)
        pagina.execute_script("window.scrollTo(0, 2500);")
        sleep(5)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[22]/label/textarea').send_keys(lista[22])
      
        sleep(5)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[23]/div[2]').click() #Botão de avançar
        sleep(5)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[1]/div[2]/div[1]/app-signer-input/div[1]/input').send_keys('alessandra.davanzo@opipa.org')
        sleep(2)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[1]/div[2]/div[2]/app-signer-input/div[1]/input').send_keys(f'{lista[16]}')
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[1]').click()
        sleep(1)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[2]/div[2]/app-signer-floating-fields-input/div/div[3]/div[1]').click()
        sleep(15)
        pagina.execute_script("window.scrollTo(0, 2500);")
        sleep(1)
        elem =pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/app-pdf-viewer/div[1]/div/canvas[3]')
        ac = ActionChains(pagina)
        ac.move_to_element(elem).move_by_offset(-180, 100).click().perform()
        sleep(2)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[2]/div[2]/app-signer-floating-fields-input[2]/div/div[3]/div[1]').click()
        pagina.execute_script("window.scrollTo(0, 2500);")
        sleep(2)
        elem =pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/app-pdf-viewer/div[1]/div/canvas[3]')
        ac = ActionChains(pagina)
        ac.move_to_element(elem).move_by_offset(180, 100).click().perform()
        #Clicar em avançar
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[4]').click()
        sleep(2)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[3]/app-document-configs-form/label[1]/input').send_keys(f' - {lista[0]}')
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[2]').click()
        sleep(5)
        pagina.get('https://painel.autentique.com.br/documentos/todos')
        print(lista)
        sleep(40)
start = autentique()
start.iniciar()