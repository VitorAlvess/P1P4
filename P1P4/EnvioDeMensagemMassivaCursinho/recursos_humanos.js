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
        // if (row[6] == 'TRUE') {
            

        //     if (row[16] == 'Carregando…') {
        //         console.log('Existe algum problema no carregamento de dados da API do google sheets')
        //         return
        //     }
        //     if (row[16] == '#N/A' && row[11] != 'ADM Cursinho') {
        //       console.log('Mensagem de vaga não cadastrada', index + 1)
        //       return
              
        //   }
        //     if (row[5] == '') {
              
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem1 = row[16]
        //         mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
        //         mensagem2 = row[17]
        //         mensagem3 = row[18]
        //         valores.push([[numero], [mensagem1], [mensagem2], [mensagem3]])
        //         coluna = "F"
        //         adicionar_data(coluna,index + 1)
        //         mandar_email_inicial(nome, numero,row[8])
               
        //     }

        //     if (row[4] == "Marcar entrevista" && row[3] == '') {
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem4 = row[19]
        //         mensagem4 = mensagem4.replace("[primeiro nome]", nome.split(" ")[0])
        //         valores.push([[numero],[mensagem4]])
        //         coluna = "D"
        //         adicionar_data(coluna,index + 1)   
        //         adicionar_texto("C", index + 1, "Aguardando agendamento")
        //     }

        //     if (row[4] == "Incompatível com a vaga" && row[0] == '') {
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem5 = row[20]
        //         mensagem5 = mensagem5.replace("[primeiro nome]", nome.split(" ")[0])
        //         valores.push([[numero],[mensagem5]])
        //         coluna = "A"
        //         texto = 'Reprovado(a) e informado(a)'
        //         adicionar_texto(coluna, index + 1, texto)
        //       }

        //       if (row[5] != '' && row[4] == '') {
        //         let data_planilha = row[5]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 7) {
        //         //   console.log("A data fornecida está mais de 7 dias no passado");
        //             nome = row[7]
        //             numero = row[9]
        //             mensagem6 = row[21]
        //             texto = "Primeira semana sem responder"
        //             coluna = "E"
        //             mensagem6 = mensagem6.replace("[primeiro nome]", nome.split(" ")[0])
        //             adicionar_texto(coluna, index + 1, texto)
        //             valores.push([[numero],[mensagem6]])
        //         //   linha_adicionar_cinza.push(index + 1)
        //         //   console.log(valores_cinza)
        //         } 
        //       }
              
        //       if (row[5] != '' && row[4] == 'Primeira semana sem responder') {
        //         let data_planilha = row[5]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 14) {
        //         //   console.log("A data fornecida está mais de 21 dias no passado");
        //             nome = row[7]
        //             numero = row[9]
        //             // mensagem7 = row[22]
        //             texto = "E-mail enviado"
        //             coluna = "E"
        //             // mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])
        //             adicionar_texto(coluna, index + 1, texto)
        //             console.log('Alterei', index + 1)
        //             console.log('Nome da vaga', row[11])
        //             textos = mensagem_email.data.values;
        //             // console.log(textos[index])
                   
        //             console.log(row[8])
        //             mandar_email(nome,textos[index],row[11],numero,row[8])
              
                   
        //         } 
        //       }







        //       if (row[5] != '' && row[4] == 'E-mail enviado') {
        //         let data_planilha = row[5]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 21) {
        //         //   console.log("A data fornecida está mais de 21 dias no passado");
        //             nome = row[7]
        //             numero = row[9]
        //             mensagem7 = row[22]
        //             texto = "Terceira semana sem responder"
        //             coluna = "E"
        //             mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])
        //             adicionar_texto(coluna, index + 1, texto)

        //             adicionar_texto('A', index + 1, 'Duas tentativas sem resposta')
        //             valores.push([[numero],[mensagem7]])
        //         } 
        //       }

        //       if (row[2] == "Realizada" && row[0] == 'Aprovado(a)' && row[1] == '') {
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem8 = row[23]
        //         mensagem8 = mensagem8.replace("[Primeiro nome]", nome.split(" ")[0])
        //         valores.push([[numero],[mensagem8]])
        //         adicionar_data('B', index + 1)
        //         adicionar_texto("A", index + 1, "Responder formulário adesão")
        //       }

        //       if (row[2] == "Aguardando agendamento" && row[3] != '') { 
        //         let data_planilha = row[3]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 7) {
        //             numero = row[9]
        //             nome = row[7]
        //             mensagem9 = row[24]
        //             mensagem9 = mensagem9.replace("[Primeiro nome]", nome.split(" ")[0])
        //             mensagem9 = mensagem9.replace("[primeiro nome]", nome.split(" ")[0])

        //             valores.push([[numero],[mensagem9]])
        //             adicionar_texto("C", index + 1, "Mensagem enviada de aguardando agendamento")
        //         }
        //       }
        //       if (row[2] == 'Mensagem enviada de aguardando agendamento' && row[0] == ''){
        //         let data_planilha = row[3]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 14) {
        //             numero = row[9]
        //             nome = row[7]
        //             mensagem9 = row[24]
        //             mensagem9 = mensagem9.replace("[Primeiro nome]", nome.split(" ")[0])
        //             mensagem9 = mensagem9.replace("[primeiro nome]", nome.split(" ")[0])

        //             valores.push([[numero],[mensagem9]])
        //             adicionar_texto("C", index + 1, "Segunda mensagem de aguardando agendamento enviada")
        //         }
        //       }

        //       if (row[2] == 'Segunda mensagem de aguardando agendamento enviada' && row[0] == ''){
        //         let data_planilha = row[3]
        //         let currentDate = new Date();
        //         let dataArray = data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 21) {
        //             numero = row[9]
        //             nome = row[7]
        //             mensagem7 = row[22]
        //             mensagem7 = mensagem7.replace("[Primeiro nome]", nome.split(" ")[0])
        //             mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])

        //             valores.push([[numero],[mensagem7]])
        //             adicionar_texto("A", index + 1, "Duas tentativas sem resposta")
        //         }

        //       }
        //       if (row[2] == "Faltou primeira vez") {
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem10 = row[25]
        //         mensagem10 = mensagem10.replace("[Primeiro nome]", nome.split(" ")[0])
        //         mensagem10 = mensagem10.replace("[primeiro nome]", nome.split(" ")[0])
        //         valores.push([[numero],[mensagem10]])
        //         adicionar_texto("C", index + 1, "Segunda tentativa de agendamento")    
        //     }

        //     if (row[2] == "Faltou segunda vez" && row[0] != 'Duas faltas no agendamento') {
        //         nome = row[7]
        //         numero = row[9]
        //         mensagem11 = row[26]
        //         mensagem11 = mensagem11.replace("[Primeiro nome]", nome.split(" ")[0])
        //         mensagem11 = mensagem11.replace("[primeiro nome]", nome.split(" ")[0])
        //         valores.push([[numero],[mensagem11]])
        //         adicionar_texto("A", index + 1, "Duas faltas no agendamento")    
        //     }

        //     if(row[0] == 'P1P4 Enviar mensagem de termo de adesão'){
        //       numero = row[9]
        //       nome = row[7]
        //       mensagem_termo_assinar = 'Oi, [Primeiro nome], agradecemos por ter fornecido as informações. Acabei de enviar para o e-mail informado o termo para ser assinado digitalmente. É só seguir o passo a passo. É super simples, mas se tiver alguma dúvida, pode me chamar que eu respondo em breve'
        //       mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
        //       mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])

        //       valores.push([[numero],[mensagem_termo_assinar]])
        //       adicionar_texto('A', index + 1, 'Termo adesão enviado')
        //     }




        //     // está com Termo de adessão assinado e com uma data 7 dias no passado. 
        //     if (row[0] == "Termo adesão enviado") {
        //       let data_planilha = row[1]
        //       let currentDate = new Date();
        //       let dataArray = data_planilha.split("/");
        //       let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //       let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //       if (diferenca > 7) {
        //           numero = row[9]
        //           nome = row[7]
        //           mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])

        //           valores.push([[numero],[mensagem_termo_assinar]])
        //           adicionar_data_termo_adesao("B", index + 1, '¹')
        //       }
        //       if (row[1].charAt() == '¹'){
                
        //         let data_planilha = row[1]
        //         nova_data_planilha = data_planilha.substring(1);
        //         let currentDate = new Date();
        //         let dataArray = nova_data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 7) {
        //           numero = row[9]
        //           nome = row[7]
        //           mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])

        //           valores.push([[numero],[mensagem_termo_assinar]])
        //           adicionar_data_termo_adesao("B", index + 1, '²')
        //         }

        //       }
        //       if (row[1].charAt() == '²'){
                
        //         let data_planilha = row[1]
        //         nova_data_planilha = data_planilha.substring(1);
        //         let currentDate = new Date();
        //         let dataArray = nova_data_planilha.split("/");
        //         let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
        //         let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
        //         if (diferenca > 7) {
        //           numero = row[9]
        //           nome = row[7]
        //           mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
        //           mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])

        //           valores.push([[numero],[mensagem_termo_assinar]])
        //           adicionar_data('B', index + 1)
        //           adicionar_texto('A', index + 1, 'ADM Contatar')
        //         }

        //       }

                   
               
        //       // nome = row[7]
        //       // numero = row[9]
        //       // mensagem11 = row[26]
        //       // mensagem11 = mensagem11.replace("[Primeiro nome]", nome.split(" ")[0])
        //       // mensagem11 = mensagem11.replace("[primeiro nome]", nome.split(" ")[0])
        //       // valores.push([[numero],[mensagem11]])
        //       // adicionar_texto("A", index + 1, "Duas faltas no agendamento")    
        //   }
        // }


        if(row[0] == 'Termo adesão enviado'){
            nome = row[3]
            numero = row[21]
            
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

            
            mensagem1 = 'Oi, [primeiro nome], tudo certo? Acabamos de atualizar o seu termo de voluntariado no PiPA. E precisamos que você confira os dados se estão corretos e, se estiver tudo certinho, assine digitalmente o termo.'
            mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
            mensagem2 = `Se tiver alguma inconsistência ou alguma informação que precise ser atualizada, é só responder ao formulário a seguir e a gente corrige e envia outro https://curt.link/atualiza-cadastro-pipa
Um abraço 😊`

            
            let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)

           
            
            valores.push([[resultado1], [mensagem1], [mensagem2]])
            valores.push([[resultado2], [mensagem1], [mensagem2]])

            
           
            // mensagem2 = row[17]
            // mensagem3 = row[18]
            
            adicionar_texto("A", index + 1, "Mensagem para assinar enviada")
            adicionar_data("B",index + 1) 

        }
        
        if(row[0] == 'Mensagem para assinar enviada'){
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


                valores.push([[resultado1], [mensagem3]])
                valores.push([[resultado2], [mensagem3]])

                adicionar_texto("A", index + 1, "Segunda mensagem para assinar enviada")
                adicionar_data("B",index + 1) 
                }
        }

        if(row[0] == 'Segunda mensagem para assinar enviada'){
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
                valores.push([[resultado1], [mensagem3]])
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
