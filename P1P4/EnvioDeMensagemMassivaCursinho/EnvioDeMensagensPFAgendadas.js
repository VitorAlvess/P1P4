function sheets(){
  const fs = require('fs').promises;
    const path = require('path');
    const process = require('process');
    const {authenticate} = require('@google-cloud/local-auth');
    const {google} = require('googleapis');
    
    var valores = []
    // If modifying these scopes, delete token.json.
    const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    const TOKEN_PATH = path.join(process.cwd(), 'token.json');
    const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

    /**
     * Reads previously authorized credentials from the save file.
     *
     * @return {Promise<OAuth2Client|null>}
     */
    async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
    }

    /**
     * Serializes credentials to a file comptible with GoogleAUth.fromJSON.
     *
     * @param {OAuth2Client} client
     * @return {Promise<void>}
     */
    async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
    }

    /**
     * Load or request or authorization to call APIs.
     *
     */
    async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
    }

    /**
     * Prints the names and majors of students in a sample spreadsheet:
     * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
     * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
     */
    async function listMajors(auth) {
    const sheets = google.sheets({version: 'v4', auth});
    const res = await sheets.spreadsheets.values.get({
        spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
        range: 'Mensagens Agendadas!A:F',
        // ERA ATÉ AC ANTES
    });
    const contatos = await sheets.spreadsheets.values.get({
      spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
      range: 'Contatos!A:AJ',
    });

    const rows_contatos = contatos.data.values
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }
    rows.forEach((row, index) => {

        
        if (row[3] == obterDataAtual() && row[4] != 'Sim') {
            

          
            if (row[2] == 'Todos') {
           
                rows_contatos.forEach((row_contato, index_contato) => {
                    if (row_contato[2] != undefined && row_contato[2] != 'Nome da Pessoa' && row_contato[0] != "VERDADEIRO") {
                        nome = row_contato[2]
                        // numero = "11945274604"
                        numero = row_contato[3]
                        mensagem1 = row[0]
                        mensagem1 = mensagem1.replace("[nome]", nome.split(" ")[0])
                        
                        try {
                          mensagem1 = mensagem1.replace("[Nome]", nome.split(" ")[0])
                          mensagem1 = mensagem1.replace("[NOME]", nome.split(" ")[0])
                        } catch (error) {
                          console.error('Error:', error);
                        }
                        

                       
                        valores.push([[numero], [mensagem1]])
                       
                    }

                    
                })

            }



            

            if (row[2] != 'Todos') {
              
              sheets_enviar_mensagem(row[2], row)

              
              
            

          }

            adicionar_texto("E", index +1, "Sim")
            adicionar_data("F", index +1)
            adicionar_texto("G", index +1, row[1])
            adicionar_texto("H", index +1, row[2])


            // adicionar_texto("B", index +1, "FALSE")
            // adicionar_texto("G", index +1, row[2])

            


            
            
        }


            
            




      });


      function sheets_enviar_mensagem(texto_verificar,row){
        rows_contatos.forEach((row_contato, index_contato) => {
          if (row_contato[2] != undefined && row_contato[2] != 'Nome do Responsável' && row_contato[0] != "VERDADEIRO") {
              if (row_contato[5] == texto_verificar || row_contato[6] == texto_verificar || row_contato[7] == texto_verificar || row_contato[8] == texto_verificar || row_contato[9] == texto_verificar) {
                  nome = row_contato[2]
                  numero = row_contato[3]

                  mensagem1 = row[1]
                  mensagem1 = mensagem1.replace("[nome]", nome.split(" ")[0])
              

                  try {
                    mensagem1 = mensagem1.replace("[Nome]", nome.split(" ")[0])
                    mensagem1 = mensagem1.replace("[NOME]", nome.split(" ")[0])
                  } catch (error) {
                    console.error('Error:', error);
                  }
                  
                  valores.push([[numero], [mensagem1]])
                  
              }
          }

          
      })
      }
      return valores
      function adicionar_data(coluna,linha){
        let data = new Date();
        let dia = String(data.getDate()).padStart(2, '0');
        let mes = String(data.getMonth() + 1).padStart(2, '0');
        let ano = data.getFullYear();
        dataAtual = dia + '/' + mes + '/' + ano;
        let values = [
          [
          dataAtual
          ],
        ];
        const resource = {
          values,
        };
    
        try {
          const result = sheets.spreadsheets.values.update({
            spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
            range: `Mensagens Agendadas!${coluna+linha}`,
            valueInputOption: 'RAW',
            resource,
          });
    
          // Aguarde 2 segundos antes de retornar o resultado
          setTimeout(() => {
            console.log(result);
            return result;
          }, 500);
        } catch (err) {
          // TODO (Developer) - Handle exception
          throw err;
      }
        }





        function adicionar_texto(coluna,linha,texto){
        
        let values = [
          [
          texto
          ],
        ];
        const resource = {
          values,
        };
    
        try {
          const result = sheets.spreadsheets.values.update({
            spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
            range: `Mensagens Agendadas!${coluna+linha}`,
            valueInputOption: 'USER_ENTERED',
            resource,
          });
    
          setTimeout(() => {
            console.log(result);
            return result;
          }, 2000);
        } catch (err) {
          // TODO (Developer) - Handle exception
          throw err;
      }

        }
        
    }

    
    return authorize().then(listMajors).catch(console.error)
}











sheets().then((valores) => {
    console.log(valores); // aqui você pode fazer o que quiser com a variável ar
    if (valores == '') {
        console.log('Sem dados')
        valores.push([['11985848901'], [`P1P4 🤖🪁 responsável pela *Envio de Mensagem Massivo Agendadas* funcionando! ${getCurrentDateTimeBrazilian()} \n`]])
        valores.push([['11945274604'], [`P1P4 🤖🪁 responsável pela *Envio de Mensagem Massivo Agendadas* funcionando! ${getCurrentDateTimeBrazilian()} \n`]])
        whats(valores)
        
    }
    else{
        whats(valores)
        

    }
  });



  async function whats(todas_acoes) {

    const qrcode = require('qrcode-terminal');
    const { Client, LocalAuth } = require('whatsapp-web.js');
    const client = new Client({
        authStrategy: new LocalAuth(),
        webVersion: "2.2412.54",
        webVersionCache: {
        type: "remote",
        remotePath:
          "https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.2412.54.html",
      },
    });

    client.on('qr', qr => {
        qrcode.generate(qr, {small: true});
    });

    client.on('ready', () => {
        console.log('Client is ready!');

        formatado = [];
        const messagePromises = [];

        for (let index = 0; index < todas_acoes.length; index++) {
            for (let index_dentro = 0; index_dentro < todas_acoes[index].length; index_dentro++) {
                array = [];
                if (index_dentro == 0) {
                    var numero = todas_acoes[index][index_dentro];
                } else {
                    const element = todas_acoes[index][index_dentro];
                    numero_enviar = '55' + String(numero).replace(/\D/g, '') + '@c.us';
                    formatado.push([numero_enviar, element[0]]);
                }
            }
        }

        console.log(formatado);

        function enviarMensagens(index) {
            if (index >= formatado.length) {
                client.destroy();
                console.log('Todas as mensagens foram enviadas.');
                return;
            }

            const enviar = formatado[index];
            client.sendMessage(enviar[0], '')
                .then(() => client.sendMessage(enviar[0], enviar[1]))
                .then(() => console.log(enviar[0], enviar[1]));

            client.sendMessage('5511945274604@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);
            client.sendMessage('5511985848901@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);

            setTimeout(() => {
                enviarMensagens(index + 1);
            }, 8000);
        }

        enviarMensagens(0);
    });

    client.on('message', message => {
        if (message.body === '!ping') {
            client.sendMessage(message.from, 'pong');
        }
    });

    // Use await to ensure the initialization finishes before continuing
    await client.initialize();
}




function getCurrentDateTimeBrazilian() {
    const currentDate = new Date();
  
    const dayOfWeek = [
        "Domingo", "Segunda-feira", "Terça-feira",
        "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"
    ][currentDate.getDay()];
  
    const day = String(currentDate.getDate()).padStart(2, '0');
    const month = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ][currentDate.getMonth()];
  
    const year = currentDate.getFullYear();
  
    const hours = String(currentDate.getHours()).padStart(2, '0');
    const minutes = String(currentDate.getMinutes()).padStart(2, '0');
    const seconds = String(currentDate.getSeconds()).padStart(2, '0');
  
    const formattedDateTime = `${dayOfWeek}, ${day} de ${month} de ${year} ${hours}:${minutes}:${seconds}`;
    return formattedDateTime;
  }
 

function obterDataAtual() {
  const dataAtual = new Date();
  const dia = String(dataAtual.getDate()).padStart(2, '0');
  const mes = String(dataAtual.getMonth() + 1).padStart(2, '0'); // Mês começa do zero
  const ano = dataAtual.getFullYear();
  return `${dia}/${mes}/${ano}`;
}