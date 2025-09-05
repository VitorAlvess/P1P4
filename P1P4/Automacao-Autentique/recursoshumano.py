# -*- coding: utf-8 -*-

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
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
        self.navegar = webdriver.Chrome(service=servico, options=options)
    def autentique_login(self):
        pagina = self.navegar
        pagina.get('https://painel.autentique.com.br/entrar')
        sleep(5)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/label[1]/input' ).send_keys(self.email)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/div/button' ).click()
        sleep(5)
        print('Agora vai a senha')
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/label[2]/input' ).send_keys(self.senha)
        pagina.find_element('xpath',' /html/body/app-root/app-auth-login/app-auth-container/div/div/form/div/button[2]' ).click()
        sleep(5)
        print('chegou até aqui')

    
    def configuracoes_sheets(self):
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = None
       
        if os.path.exists('token.json'): #necessario pegar esse token la da api do google sheets
           self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
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
        
        SAMPLE_RANGE_NAME = 'Termo de Voluntariado Atualizado!A:AA'
       
        try:
            service = build('sheets', 'v4', credentials=self.creds)
            #Passar valores existentes para a variavel valore
            sheet = service.spreadsheets()


            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                            range=SAMPLE_RANGE_NAME).execute()
            valores = result['values']
            # print(result)
         
           #Informar que o formulario ja foi gerado 
            


            arrays_nao_vazias = list(filter(lambda x: len(x) > 0, valores[1:]))
            
           
          


            for lista in arrays_nao_vazias:
                if lista[0] != '': 
                 
                    if 'Carregando…' not in lista and '#N/A' not in lista and lista[-1] != 'TRUE' and 'Fazer novo termo' in lista:
                        print(lista)
                        posicao = result['values'].index(lista)
                        print(posicao + 1)
                        lista.pop()
                        situacao = self.autentique_dados_voluntarios(lista)
                        result2 = service.spreadsheets().values().update(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                        range=f'Termo de Voluntariado Atualizado!AA{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [['TRUE']]
                        }
                        ).execute()
                        result6 = service.spreadsheets().values().update(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                        range=f'Termo de Voluntariado Atualizado!AB{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [[f'{lista[22]}']]
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
                        range=f'Termo de Voluntariado Atualizado!C{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [[f'{data_atual}']]
                        }
                        ).execute()

                        result5 = service.spreadsheets().values().update(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                        range=f'Termo de Voluntariado Atualizado!B{posicao + 1}',
                        valueInputOption='USER_ENTERED',
                        body={
                            'values': [['Termo adesão enviado']]
                        }
                        ).execute()

                        if situacao == 'Termo Impossível De Ser Criado Pelo P1P4':
                            result7 = service.spreadsheets().values().update(
                            spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=f'Termo de Voluntariado Atualizado!B{posicao + 1}',
                            valueInputOption='USER_ENTERED',
                            body={
                                'values': [['Termo Impossível De Ser Criado Pelo P1P4']]
                            }
                            ).execute()



                    if 'Termo adesão assinado' in lista:
                       
                        data_atual = lista[2]
                      
                        data_formatada = datetime.strptime(data_atual, "%d/%m/%Y")

                        # Adicionando um ano à data fornecida
                        um_ano_depois = data_formatada + timedelta(days=365)
                        # Obtendo a data atual
                        data_atual = datetime.now()

                        # Verificando se a data atual é posterior à data um ano após a data fornecida
                        if data_atual >= um_ano_depois:
                            print('Está um ano depois a data')
                            print(lista)
                            print(data_atual)

                            posicao = result['values'].index(lista)
                            print(posicao + 1)
                            lista.pop()
                            situacao = self.autentique_dados_voluntarios(lista)
                            result2 = service.spreadsheets().values().update(
                            spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=f'Termo de Voluntariado Atualizado!AA{posicao + 1}',
                            valueInputOption='USER_ENTERED',
                            body={
                                'values': [['TRUE']]
                            }
                            ).execute()
                            result6 = service.spreadsheets().values().update(
                            spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=f'Termo de Voluntariado Atualizado!AB{posicao + 1}',
                            valueInputOption='USER_ENTERED',
                            body={
                                'values': [[f'{lista[22]}']]
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
                            range=f'Termo de Voluntariado Atualizado!C{posicao + 1}',
                            valueInputOption='USER_ENTERED',
                            body={
                                'values': [[f'{data_atual}']]
                            }
                            ).execute()

                            result5 = service.spreadsheets().values().update(
                            spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                            range=f'Termo de Voluntariado Atualizado!B{posicao + 1}',
                            valueInputOption='USER_ENTERED',
                            body={
                                'values': [['Termo adesão enviado']]
                            }
                            ).execute()

                            if situacao == 'Termo Impossível De Ser Criado Pelo P1P4':
                                result7 = service.spreadsheets().values().update(
                                spreadsheetId=SAMPLE_SPREADSHEET_ID, 
                                range=f'Termo de Voluntariado Atualizado!B{posicao + 1}',
                                valueInputOption='USER_ENTERED',
                                body={
                                    'values': [['Termo Impossível De Ser Criado Pelo P1P4']]
                                }
                                ).execute()
                       
                        
       
        except HttpError as err:
                print(err)

    def autentique_dados_voluntarios(self, lista):
        print('Pagina do autentique')
        pagina = self.navegar
        sleep(30)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-home/app-documents-list/div[1]/app-taxonomy/aside/app-documents-taxonomy/div[1]/button').click()
        sleep(1)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/div[1]/button[2]').click()
        
        sleep(4)
        pagina.execute_script("window.scrollTo(0, 600);")
       
        elements = WebDriverWait(pagina, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "template__item-name")))

            # Texto a ser verificado
        target_text = "Atualização - Termo de Adesão Voluntariado"

        # Iterar sobre os elementos
        for element in elements:
            if target_text in element.text:
                element.click()
                break  # Para clicar apenas no primeiro elemento que corresponde








        print('cheguei aqui')
        sleep(10)
        pagina = self.navegar
        print(lista)
        

        
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[1]/label/input').send_keys(lista[0])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[2]/label/textarea').send_keys(lista[3])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[3]/label/input').send_keys(lista[4])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[4]/label/input').send_keys(lista[5])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[5]/label/input').send_keys(lista[6])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[6]/label/input').send_keys(lista[7])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[7]/label/input').send_keys(lista[8])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[8]/label/input').send_keys(lista[9])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[9]/label/input').send_keys(lista[10])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[10]/label/input').send_keys(lista[11])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[11]/label/input').send_keys(lista[12])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[12]/label/input').send_keys(lista[13])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[13]/label/input').send_keys(lista[14])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[14]/label/input').send_keys(lista[15])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[15]/label/input').send_keys(lista[16])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[16]/label/input').send_keys(lista[17])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[17]/label/input').send_keys(lista[18])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[18]/label/input').send_keys(lista[19])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[19]/label/input').send_keys(lista[20])
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[20]/label/input').send_keys(lista[21])
        
        
        data_atual = date.today()
        sleep(5)
        data_em_texto = data_atual.strftime('%d/%m/%Y')
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[21]/label/input').send_keys(data_em_texto)
        pagina.execute_script("window.scrollTo(0, 2500);")
        sleep(5)
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[22]/label/textarea').send_keys(lista[22])

        #coisa novas a partir do 


        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[23]/label/input').send_keys(lista[24])

        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[24]/label/input').send_keys(lista[23])
    
        pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[25]/label/input').send_keys(lista[25])
        

        #NOVAS INFORMACOES DE OUTRA PLANILHA

         # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = None

        if os.path.exists('token.json'): #necessario pegar esse token la da api do google sheets
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(self.creds.to_json())
            # The ID and range of a sample spreadsheet.
            SAMPLE_SPREADSHEET_ID = '1eF3fhE-99ejF0K1fB7nB2WoLghAy9tVlYgseg_bvtTI'

            SAMPLE_RANGE_NAME = 'Quadro Termo - para P1P4!I:P'

            novas_informacoes = []

        try:
            service = build('sheets', 'v4', credentials=self.creds)
            #Passar valores existentes para a variavel valore
            sheet = service.spreadsheets()


            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                            range=SAMPLE_RANGE_NAME).execute()
            valores = result['values']
            # print(result)

            #Informar que o formulario ja foi gerado 



            arrays_nao_vazias = list(filter(lambda x: len(x) > 0, valores[1:]))
            arrays_nao_vazias = list(filter(lambda x: len(x) > 0, valores[1:]))


            data = []
            print(lista[0])
            print(lista[3])
            for lista_novos in arrays_nao_vazias:
                if lista_novos[0] != '': 
                    if 'Carregando…' not in lista_novos and '#N/A' not in lista_novos:
                        lista_str = ", ".join(lista_novos)
                        cleaned_lista_str = lista_str.replace("[", "").replace("]", "").replace("'", "").replace("\n", "").replace(",", "")
                        if lista[3] in  cleaned_lista_str and '$' not in lista_novos[0]:
                            data.append(lista_novos)
                        


        except HttpError as err:
            print(err)

        print(data)
        data_nova = []
        for sublist in data:
            del(sublist[1])

        # Print the updated data
        for sublist in data:
            data_nova.append(sublist)
        print(data_nova)

        global_counter = 1
        situacao_panico = False
        for atualizado in data_nova:
            for element in atualizado:
                if situacao_panico == True:
                    break
                print("Global Counter:", global_counter)
                print("Element:", element)
                print("-----")
                global_counter += 1
                xpath_index = global_counter + 24  # Calculate the index for the xpath
                
                pagina.find_element('xpath', f'/html/body/app-root/app-documents-fill-template/aside/form/div/div[{xpath_index}]/label/input').send_keys(element)
                
                
                
                if element == 'Preencher todas as células laranjas ou amarelas para correto funcionamento do envio do termo de adesão ao voluntariado':
                    print('Algum dado está vazio com isso não é possivel criar o termo')
                    return 'Termo Impossível De Ser Criado Pelo P1P4'
                
                
                
                
                if global_counter == 50:
                    situacao_panico = True
                    break
        print(global_counter)
        if situacao_panico != True:


            sleep(15) #sleep antes do avnçar
            pagina.find_element('xpath', '/html/body/app-root/app-documents-fill-template/aside/form/div/div[96]/div[2]').click() #Botão de avançar
            sleep(5)
            
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[1]/div[6]/div/app-signer-input/div[1]/input[1]').send_keys('alessandra.davanzo@opipa.org')
            
            sleep(2)
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[1]/div[6]/div[2]/app-signer-input/div[1]/input').send_keys(f'{lista[19]}')
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[1]').click()
            sleep(1)
           


            sleep(10)
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[2]/div[2]/app-signer-floating-fields-input/div/div[3]/div[1]').click()
            sleep(3)
            pagina.execute_script("window.scrollTo(0, 3150);")
            sleep(1)
            elem =pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/app-pdf-viewer/div[1]/div/canvas[3]')
            ac = ActionChains(pagina)
            ac.move_to_element(elem).move_by_offset(-180, 300).click().perform()
            sleep(2)
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[2]/div[2]/app-signer-floating-fields-input[2]/div/div[3]/div[1]').click()
            pagina.execute_script("window.scrollTo(0, 3150);")
            sleep(2)
            elem =pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/section/app-pdf-viewer/div[1]/div/canvas[3]')
            ac = ActionChains(pagina)
            ac.move_to_element(elem).move_by_offset(180, 300).click().perform()
            sleep(10)

            #Clicar em avançar
            
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[4]').click()
            sleep(2)
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/main[3]/app-document-configs-form/label[1]/input').send_keys(f' - {lista[3]}') #Nome no final do termo
            pagina.find_element('xpath', '/html/body/app-root/app-documents-new/div/aside[2]/div/button[2]').click()
            sleep(5)
            pagina.get('https://painel.autentique.com.br/documentos/todos')
            print(lista)
            sleep(10)
        else:
            pagina.get('https://painel.autentique.com.br/documentos/todos')
            return 'Termo Impossível De Ser Criado Pelo P1P4'
start = autentique()
start.iniciar()