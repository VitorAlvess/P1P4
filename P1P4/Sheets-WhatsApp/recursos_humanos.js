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
        spreadsheetId: '1eF3fhE-99ejF0K1fB7nB2WoLghAy9tVlYgseg_bvtTI',
        range: 'Termo de Voluntariado Atualizado!A:X',
    });
    

    
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }
    rows.forEach((row, index) => {
       

       
        if(row[1] == 'Termo Impossível De Ser Criado Pelo P1P4'){
          nome = row[3]
          mensagem_impossivel = `'Olá o P1P4 🤖🪁 pede ajuda, não consigo criar o termo do humano ${nome}. Ele possui mais de *7 atividades*! `
          valores.push([['11985848901'], [mensagem_impossivel]])
          adicionar_texto("B", index + 1, "Mensagem Enviada Que o P1P4 não consegue Criar o Termo")
        }


        if(row[1] == 'Termo adesão enviado'){
            nome = row[3]
            numero = row[21]
            email = row[19]
            
            // function formatarTelefone(numero) {
            //   const numeros = numero.replace(/\D/g, '');
            
            //   // Verificar se o número tem 11 dígitos
            //   if (numeros.length === 13) {
            //     if (numero.substring(0, 2) === "55") {
            //       // Remove os dois primeiros dígitos
            //       numero = numero.substring(2);
            //       console.log("Número válido após remoção dos dígitos iniciais: " + numero);
                  
            //       console.log(numero)
            //       return numero
            //     }
            //   }
            //   // Verificar se o número tem 10 dígitos
            //   // else if (numeros.length === 10) {
            //   //   return `(${numeros.substring(0, 2)}) ${numeros.substring(2, 6)}-${numeros.substring(6)}`;
            //   // }
            //   // Caso contrário, retornar o número original
            //   else {
            //     return numero;
            //   }
            // }

            // numero_formatado = formatarTelefone(numero)

            
            mensagem1 = `Oi, [nome], tudo certo?

Quem tá falando aqui é o P1P4 - o robozinho do PiPA 🪁🤖
            
Acabamos de atualizar o seu termo de voluntariado e o enviamos para o seu e-mail cadastrado: [e-mail]. 
            
Agora, precisamos que você confira se os dados estão corretos e, se estiver tudo certinho, assine o termo digitalmente.`
            mensagem1 = mensagem1.replace("[nome]", nome.split(" ")[0])
            mensagem1 = mensagem1.replace("[e-mail]", email.split(" ")[0])

            mensagem2 = `Se tiver alguma inconsistência ou alguma informação que precise ser atualizada, é só responder ao formulário a seguir e a gente corrige e envia outro https://curt.link/atualiza-cadastro-pipa
Um abraço 😊`

            
            let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
            const numeroAlterado = removerDigitoTelefone(numero);
           
            
            valores.push([[resultado1], [mensagem1], [mensagem2]])
            valores.push([[numeroAlterado], [mensagem1], [mensagem2]])
            valores.push([[resultado2], [mensagem1], [mensagem2]])

            
           
            // mensagem2 = row[17]
            // mensagem3 = row[18]
            
            adicionar_texto("B", index + 1, "Mensagem para assinar enviada")
            adicionar_data("C",index + 1) 

        }
        
        if(row[1] == 'Mensagem para assinar enviada'){
            let data_planilha = row[1]
            let currentDate = new Date();
            let dataArray = data_planilha.split("/");
            let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
            let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
            if (diferenca > 7) { //data maior que 14 dias
                nome = row[3]
                numero = row[21]
                mensagem3 = `Oi, [primeiro nome]. Espero que esteja tudo bem por aí... 

Nós enviamos para o seu e-mail o *Termo de Adesão ao Voluntariado - PiPA* atualizado e ainda não recebemos a confirmação de sua assinatura. 
                
Poderia verificar se por acaso caiu no spam? Vai chegar pela plataforma *Autentique* é só digitar esse nome na busca do seu e-mail que você deve encontrá-lo.
                
Um abraço 😊🪁`
                mensagem3 = mensagem3.replace("[primeiro nome]", nome.split(" ")[0])
                // mensagem3 = row[17]
                // mensagem3 = row[18]

                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                const numeroAlterado = removerDigitoTelefone(numero);


                valores.push([[resultado1], [mensagem3]])
                valores.push([[numeroAlterado], [mensagem3]])
                valores.push([[resultado2], [mensagem3]])
                adicionar_texto("A", index + 1, "Segunda mensagem para assinar enviada")
                adicionar_data("B",index + 1) 
                }
        }

        if(row[1] == 'Segunda mensagem para assinar enviada'){
            let data_planilha = row[1]
            let currentDate = new Date();
            let dataArray = data_planilha.split("/");
            let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
            let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
            if (diferenca > 14) { //data maior que 14 dias
                nome = row[3]
                numero = row[21]
                mensagem3 = `Oi, [primeiro nome]. Espero que esteja tudo bem por aí... 

Nós enviamos para o seu e-mail o *Termo de Adesão ao Voluntariado - PiPA* atualizado e ainda não recebemos a confirmação de sua assinatura. 
                
Poderia verificar se por acaso caiu no spam? Vai chegar pela plataforma *Autentique* é só digitar esse nome na busca do seu e-mail que você deve encontrá-lo.
                
Um abraço 😊🪁`
                mensagem3 = mensagem3.replace("[primeiro nome]", nome.split(" ")[0])
                // mensagem3 = row[17]
                // mensagem3 = row[18]
                
                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                const numeroAlterado = removerDigitoTelefone(numero);
                valores.push([[resultado1], [mensagem3]])
                valores.push([[numeroAlterado], [mensagem3]])
                valores.push([[resultado2], [mensagem3]])

                adicionar_texto("A", index + 1, "ADM Contatar")
                adicionar_data("B",index + 1) 
                }
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
        valores.push([['11985848901'], [`P1P4 responsável pela parte de envio da planilha "Controle de Recursos Humanos" rodando, porém não existe mensagens para serem enviadas:  ${getCurrentDateTimeBrazilian()} \n`]])
        valores.push([['11945274604'], [`P1P4 responsável pela parte de envio da planilha "Controle de Recursos Humanos" rodando, porém não existe mensagens para serem enviadas:  ${getCurrentDateTimeBrazilian()} \n`]])

        whats(valores)
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
      
          const enviar = formatado[index];
          
          client.sendMessage(enviar[0], '')
          client.sendMessage(enviar[0], enviar[1])
          console.log(enviar[0], enviar[1]);


  


          client.sendMessage('5511945274604@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);
          client.sendMessage('5511985848901@c.us', `*Foi enviada com sucesso a mensagem:* \n${enviar[1]} *para o numero:*\n ${enviar[0]}`);
         

          setTimeout(() => {
              enviarMensagens(index + 1); // Chama a função para a próxima iteração após o atraso
          }, 10000); // 2 minutos em milissegundos
      }
      
      enviarMensagens(0); // Inicia o processo de envio de mensagens







    });
    client.on('message', message => {
        if(message.body === '!ping') {
            client.sendMessage(message.from, 'pong');
        }
    });
     
    client.initialize(); 
}










 // Exemplo de número de telefone
//  function duplicanumerosporcausadonove (numero_formatado){
//   var numero = String(numero_formatado)
//   var numero = numero.replace(/\D/g, "")      
  
  
 
//   var copia_numero = numero
//   if (numero.substring(2, 4) === "99") {
//     numero = numero.substring(0, 2) + numero.substring(4);
//     numero_sem_nove = numero
    
//   }
//   if (numero.substring(2, 4) != "99") {
//     numero = numero.substring(0, 2) + "9" + numero.substring(2);
//     numero_com_nove = numero
    
//   }
  
//   console.log("Número de telefone atualizado: " + numero);
//   console.log("Número de telefone antigo: " + copia_numero);

//   return { resultado1: copia_numero, resultado2: numero };
  
// }




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