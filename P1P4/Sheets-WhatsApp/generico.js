const nodemailer = require('nodemailer')
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
        range: 'Envios de Mensagem!A:M',
    });
    

    
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }
    limite = 0
    rows.forEach((row, index) => {
       


        if(row[4] == 'TRUE' && row[5] != 'TRUE' && limite < 20){
            if(row[0] != '')
            limite = limite + 1
            console.log(limite)
            nome = row[0]
            numero = row[2]
            mensagem1 = row[8]
            mensagem2 = row[9]
            mensagem3 = row[10]
            if (mensagem1.includes("[nome]")) {
              mensagem1 = mensagem1.replace("[nome]", nome.split(" ")[0]);
            }

            
            if (mensagem1.includes("[Nome]")) {
              mensagem1 = mensagem1.replace("[Nome]", nome.split(" ")[0]);
            }

            


            
            // mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
            
            // if (mensagem2.includes("[nome]")) {
            //   mensagem2 = mensagem2.replace("[nome]", nome.split(" ")[0]);
            // }
            // mensagem1 = mensagem2.replace("[primeiro nome]", nome.split(" ")[0])
            mensagem3 = row[10]
            // if (mensagem3.includes("[nome]")) {
            //   mensagem3 = mensagem3.replace("[nome]", nome.split(" ")[0]);
            // }
            // mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
            mensagem4 = row[11]
            // if (mensagem4.includes("[nome]")) {
            //   mensagem4 = mensagem4.replace("[nome]", nome.split(" ")[0]);
            // }
            // mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
            mensagem5 = row[12]
            // if (mensagem5.includes("[nome]")) {
            //   mensagem5 = mensagem5.replace("[nome]", nome.split(" ")[0]);
            // }
            // mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
           mensagem_email = row[7]
          //  if (mensagem_email.includes("[nome]")) {
          //   mensagem_email = mensagem_email.replace("[nome]", nome.split(" ")[0]);
          // }

            
            let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)

            const numeroAlterado = removerDigitoTelefone(numero);
            
            valores.push([[resultado1], [mensagem1], [mensagem2], [mensagem3], [mensagem4], [mensagem5]])
            valores.push([[resultado2], [mensagem1], [mensagem2], [mensagem3], [mensagem4], [mensagem5]])
            valores.push([[numeroAlterado], [mensagem1], [mensagem2], [mensagem3], [mensagem4], [mensagem5]])
            if( mensagem_email != ''){
              mandar_email(nome, mensagem_email, row[3],row[6])
              console.log('MANDAR EMAIL (Só que não, pq não tem email para enviar bobão')

            }
            
            adicionar_texto("F", index + 1, "TRUE")
            todas_mensagens = [[mensagem1], [mensagem2],[mensagem3],[mensagem4],[mensagem5], [mensagem_email]]
            console.log(todas_mensagens)
            
            Append(`N${index + 1}`, limpardados(todas_mensagens))
           

          

        }
        
        
            




      });
      return valores


       
      
      



      function Append(linha, texto){
       
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
            spreadsheetId: '1QjpkQAybClkH1Bq9lf-qrNIgJGRB7ZRiHeG07Iae6Ik',
            range: `Envios de Mensagem!${linha}`,
            valueInputOption: 'RAW',
            resource,
          });
        
          // Aguarde 2 segundos antes de retornar o resultado
          setTimeout(() => {
            console.log(result);
            return result;
          }, 1500);
        
        } catch (err) {
          // TODO (Developer) - Handle exception
          throw err;
        }
        }
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
            range: `Envios de Mensagem!${coluna+linha}`,
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
              range: `Envios de Mensagem!${coluna+linha}`,
              valueInputOption: 'RAW',
              resource,
            });
          
            // Aguarde 2 segundos antes de retornar o resultado
            setTimeout(() => {
              console.log(result);
              return result;
            }, 2000);
          
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
            spreadsheetId: '1QjpkQAybClkH1Bq9lf-qrNIgJGRB7ZRiHeG07Iae6Ik',
            range: `Envios de Mensagem!${coluna+linha}`,
            valueInputOption: 'USER_ENTERED',
            resource,
          });
        
          // Aguarde 2 segundos antes de retornar o resultado
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

        function enviarMensagens(index) {
          if (index >= formatado.length) {
            client.destroy();
            console.log('Todas as mensagens foram enviadas.')
            return; // Sai da função quando todas as mensagens foram enviadas
          }
      
          const enviar = formatado[index]; //Inclui mais uma parada no index 
          
          // client.sendMessage(enviar[0], '')
          if (enviar[1] != undefined) {
            client.sendMessage(enviar[0], enviar[1])
            client.sendMessage('5511945274604@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);
            client.sendMessage('5511985848901@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);
            
          }
          console.log(enviar[0], enviar[1]); 


  


          
          
         

          setTimeout(() => {
              enviarMensagens(index + 1); // Chama a função para a próxima iteração após o atraso
          }, 3500); // 2 minutos em milissegundos
      }
      
      enviarMensagens(0); // Inicia o processo de envio de mensagens





















        // for (let enviar = 0; enviar < formatado.length; enviar++) {

        //     messagePromises.push(client.sendMessage(formatado[enviar][0], '')) //para não bugar a ordem de envio
        //     messagePromises.push(client.sendMessage(formatado[enviar][0],formatado[enviar][1]))
        //     console.log(formatado[enviar][0],formatado[enviar][1])
        //     client.sendMessage('5511945274604@c.us', `*Foi enviada com sucesso a mensagem:* \n${formatado[enviar][1]} *para o numero:*\n ${formatado[enviar][0]}`) //Mensagem informando quais mensagens foram enviadas

        //     client.sendMessage('5511985848901@c.us', `*Foi enviada com sucesso a mensagem:* \n${formatado[enviar][1]} *para o numero:*\n ${formatado[enviar][0]}`) //Mensagem informando quais mensagens foram enviadas
        // }
        
        
        // Promise.allSettled(messagePromises)
        // .then(() => {
        //   console.log('Todas as mensagens foram enviadas!');
        //   // Aguarde 2 minutos antes de encerrar a instância do WhatsApp Web
        //   setTimeout(() => {
        //     client.destroy();
        //   }, 120000);
        // })
        // .catch((error) => {
        //   console.error('Erro ao enviar mensagem:', error);
        //   // Aguarde 2 minutos antes de encerrar a instância do WhatsApp Web
        //   setTimeout(() => {
        //     client.destroy();
        //   }, 120000);
        // });





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



function mandar_email(nome, mensagem, email, assunto){
  const data = require('./email.json');
  let transporter = nodemailer.createTransport({
      host: "smtp.gmail.com",
      port: 465,
      secure: true, // true for 465, false for other ports
      auth: {
        user: 'opipa@opipa.org', // generated ethereal user
        pass: data.token, // generated ethereal password
      },
  });
  console.log(mensagem)
  email = email.toString()
  mensagem = mensagem.toString();
  // mensagem = mensagem.replace("[Primeiro nome]", nome.split(" ")[0])
  mensagem = mensagem.replace("[nome]", nome.split(" ")[0])
 
  titulo = assunto
 
    // send mail with defined transport object
  // let info = await 
  transporter.sendMail({
  from: '"PiPA 🪁" <daniel.mariano@opipa.org>', // sender address
  to: email, // list of receivers
  subject: titulo, // Subject line
  text: `${mensagem}`, // plain text body
  // html: `${mensagem}`, // html body
  })
  .then(() => console.log('Email Enviado'))
  .catch((err) => console.log('Erro ao enviar o email', err))
}


function removerDigitoTelefone(numero) {
  // Remove os caracteres não numéricos do número de telefone
  const numeroLimpo = numero.replace(/\D/g, '');

  // Verifica se o número tem o formato esperado
  if (numeroLimpo.length !== 11) {
    console.log('Número de telefone inválido. Certifique-se de que o número tenha 11 dígitos.');
    return numero;
  }

  // Remove o "9" na terceira posição
  const numeroAlterado = numeroLimpo.slice(0, 2) + numeroLimpo.slice(3);

  // Retorna o número alterado com o formato "(XX) XXXXX-XXXX"
  return `(${numeroAlterado.slice(0, 2)}) ${numeroAlterado.slice(2, 7)}-${numeroAlterado.slice(7)}`;
}

function columnToLetter(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}


function limpardados(mensagem){

  var resultado = mensagem.filter(item => item[0] !== '');
  var texto = resultado.map(item => item[0]).join('\n\n');
  let dataHoraAtual = new Date();
  let dataAtual = dataHoraAtual.toLocaleDateString();
  let horaAtual = dataHoraAtual.toLocaleTimeString();
  let textohoras = `Mensagens enviada às ${horaAtual} do dia ${dataAtual}`
  let textofinal = `${textohoras} \n\n${texto}` 
  
  return textofinal
}