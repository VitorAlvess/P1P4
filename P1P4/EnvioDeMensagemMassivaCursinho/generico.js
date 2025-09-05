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
        spreadsheetId: '1QjpkQAybClkH1Bq9lf-qrNIgJGRB7ZRiHeG07Iae6Ik',
        range: 'Envios de Mensagem!A:E',
    });
    

    
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }
    rows.forEach((row, index) => {
       


        if(row[0] == 'Termo adesão enviado'){
            nome = row[3]
            numero = row[21]
            mensagem1 = 'Oi, [primeiro nome], tudo certo? Acabamos de atualizar o seu termo de voluntariado no PiPA. E precisamos que você confira os dados se estão corretos e, se estiver tudo certinho, assine digitalmente o termo.'
            mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
            mensagem2 = `Se tiver alguma inconsistência ou alguma informação que precise ser atualizada, é só responder ao formulário a seguir e a gente corrige e envia outro https://curt.link/atualiza-cadastro-pipa
Um abraço 😊`

            
            let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)

           
            
            valores.push([[resultado1], [mensagem1], [mensagem2]])
            valores.push([[resultado2], [mensagem1], [mensagem2]])
            adicionar_texto("A", index + 1, "Mensagem para assinar enviada")
            adicionar_data("B",index + 1) 

        }
        
        
            




      });
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
            spreadsheetId: '1QjpkQAybClkH1Bq9lf-qrNIgJGRB7ZRiHeG07Iae6Ik',
            range: `Envios de Mensagem!A:E!${coluna+linha}`,
            valueInputOption: 'RAW',
            resource,
          });
    
          return result;
        } catch (err) {
          // TODO (Developer) - Handle exception
          throw err;
      }
        }

        function adicionar_data_termo_adesao(coluna,linha,ordem){
          let data = new Date();
          let dia = String(data.getDate()).padStart(2, '0');
          let mes = String(data.getMonth() + 1).padStart(2, '0');
          let ano = data.getFullYear();
          dataAtual = dia + '/' + mes + '/' + ano;
          data_formatada = ordem + dataAtual
          let values = [
            [
            data_formatada
            ],
          ];
          const resource = {
            values,
          };
      
          try {
            const result = sheets.spreadsheets.values.update({
              spreadsheetId: '1QjpkQAybClkH1Bq9lf-qrNIgJGRB7ZRiHeG07Iae6Ik',
              range: `Envios de Mensagem!A:E!${coluna+linha}`,
              valueInputOption: 'RAW',
              resource,
            });
      
            return result;
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
            spreadsheetId: '1eF3fhE-99ejF0K1fB7nB2WoLghAy9tVlYgseg_bvtTI',
            range: `Termo de Voluntariado Atualizado!${coluna+linha}`,
            valueInputOption: 'RAW',
            resource,
          });
    
          return result;
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
    }
    else{

        console.log('Enviando para o WhatsApp...')
        whats(valores)
        

    }
  });


  function whats(todas_acoes) {

    const qrcode = require('qrcode-terminal');
    const { Client, LocalAuth } = require('whatsapp-web.js');
    const client = new Client({
        authStrategy: new LocalAuth()
    });


    client.on('qr', qr => {
        qrcode.generate(qr, {small: true});
    });
    
    client.on('ready', () => {
        console.log('Client is ready!');
       
        formatado = []
        const messagePromises = [];

        for (let index = 0; index < todas_acoes.length; index++) {
            for (let index_dentro = 0; index_dentro < todas_acoes[index].length; index_dentro++) {
                array = []
                if (index_dentro == 0) {
                    var numero = todas_acoes[index][index_dentro]
                 
                }
                else{
                    const element = todas_acoes[index][index_dentro];
                    // console.log(`numero: ${numero}`)
                    // console.log(element)
                    if (numero.includes("+")) {
                      numero_enviar = String(numero).replace(/\D/g, '') + '@c.us'
                  } else {
                    numero_enviar = '55' + String(numero).replace(/\D/g, '') + '@c.us'
                  }
                    formatado.push([numero_enviar, element[0]])
                    
                    
                }   
            }
        }
        console.log(formatado)

        for (let enviar = 0; enviar < formatado.length; enviar++) {

            messagePromises.push(client.sendMessage(formatado[enviar][0], '')) //para não bugar a ordem de envio
            messagePromises.push(client.sendMessage(formatado[enviar][0],formatado[enviar][1]))
            console.log(formatado[enviar][0],formatado[enviar][1])
            client.sendMessage('5511945274604@c.us', `*Foi enviada com sucesso a mensagem:* \n${formatado[enviar][1]} *para o numero:*\n ${formatado[enviar][0]}`) //Mensagem informando quais mensagens foram enviadas

            client.sendMessage('5511985848901@c.us', `*Foi enviada com sucesso a mensagem:* \n${formatado[enviar][1]} *para o numero:*\n ${formatado[enviar][0]}`) //Mensagem informando quais mensagens foram enviadas
        }
        
        
        Promise.allSettled(messagePromises)
        .then(() => {
          console.log('Todas as mensagens foram enviadas!');
          // Aguarde 2 minutos antes de encerrar a instância do WhatsApp Web
          setTimeout(() => {
            client.destroy();
          }, 120000);
        })
        .catch((error) => {
          console.error('Erro ao enviar mensagem:', error);
          // Aguarde 2 minutos antes de encerrar a instância do WhatsApp Web
          setTimeout(() => {
            client.destroy();
          }, 120000);
        });





    });
    client.on('message', message => {
        if(message.body === '!ping') {
            client.sendMessage(message.from, 'pong');
        }
    });
     
    client.initialize(); 
}

function duplicanumerosporcausadonove (Numero){
  var numero = Numero.replace(/\D/g, '');
  // Verificar se o número tem 11 dígitos
  if (numero.length >= 13) {
    if (numero.substring(0, 2) === "55") {
      // Remove os dois primeiros dígitos
      numero = numero.substring(2);
      console.log("Número válido após remoção dos dígitos iniciais: " + numero);      
    }
  }
     
  
  
 
  var copia_numero = numero
  if (numero.substring(2, 4) === "99") {
    numero = numero.substring(0, 2) + numero.substring(4);
 
    
  }
  if (numero.substring(2, 4) != "99") {
    numero = numero.substring(0, 2) + "9" + numero.substring(2);
    
    
  }
  
  console.log("Número de telefone atualizado: " + numero);
  console.log("Número de telefone antigo: " + copia_numero);

  return { resultado1: copia_numero, resultado2: numero };
  
}
