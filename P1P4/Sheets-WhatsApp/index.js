
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
        spreadsheetId: '1Jnl_PqlDJRxemLOlDP2aFawoDNo9EHG_Ma43ZvcfOyY',
        range: 'Principal!A:BB',
        // ERA ATÉ AC ANTES
    });
    const mensagem_email = await sheets.spreadsheets.values.get({
      spreadsheetId: '1Jnl_PqlDJRxemLOlDP2aFawoDNo9EHG_Ma43ZvcfOyY',
      range: 'Principal!AZ:AZ',
    });
 

   
    const rows = res.data.values;
    if (!rows || rows.length === 0) {
        console.log('No data found.');
        return;
    }
    rows.forEach((row, index) => {
        if (row[6] == 'TRUE') {
            

            if (row[16] == 'Carregando…') {
                console.log('Existe algum problema no carregamento de dados da API do google sheets')
                return
            }
            if (row[16] == '#N/A' && row[11] != 'ADM Cursinho') {
              console.log('Mensagem de vaga não cadastrada', index + 1)
              return
              
          }
            if (row[5] == '') {
              
                nome = row[7]
                numero = row[9]
                mensagem1 = row[16]
                mensagem1 = mensagem1.replace("[primeiro nome]", nome.split(" ")[0])
                mensagem2 = row[17]
                mensagem3 = row[18]
                responsavelPelaVaga = row[14]
                
                const numeroAlterado = removerDigitoTelefone(numero);
                // valores.push([[numeroAlterado], [mensagem1], [mensagem2], [mensagem3]])

                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)



                valores.push([[resultado1], [mensagem1], [mensagem2], [mensagem3]])
                valores.push([[resultado2], [mensagem1], [mensagem2], [mensagem3]])


                coluna = "F"
                adicionar_data(coluna,index + 1)
                mandar_email_inicial(nome, numero,row[8], row[51])

            }



            if (row[4] == 'Formulário Respondido') {
              let data_planilha = row[5]
              let currentDate = new Date();
              // Obtendo o dia, mês e ano
            let dia = currentDate.getDate();
            let mes = currentDate.getMonth() + 1; // Adiciona +1 porque os meses são indexados a partir de zero
            let ano = currentDate.getFullYear();
  
            // Formatando a data
            if (dia < 10) {
                dia = '0' + dia; // Adiciona um zero à esquerda se o dia for menor que 10
            }
  
            if (mes < 10) {
                mes = '0' + mes; // Adiciona um zero à esquerda se o mês for menor que 10
            }
  
            let dataFormatada = dia + '/' + mes + '/' + ano;
  
  
  
              let dataArray = data_planilha.split("/");
              let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
              let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
              nome = row[7]
            
              responsavelPelaVaga = row[14]
              
             


              coluna = "E"
              adicionar_texto(coluna, index + 1, 'Em Análise')


              
              // Trecho de codigo para avisar o responsavel pela vaga

              responsavelPelaVaga = row[14]
              funcao = row[11]
              mensagem_responsavel = `Olá [nomeresponsavel], Aqui é o P1P4🤖🪁: Estou passando para informar que *[pessoa]* da função *[função]* já respondeu o *formulário*. Agora, você precisa analisar as respostas e informar o próximo passo`
              mensagem_responsavel = mensagem_responsavel.replace("[nomeresponsavel]", responsavelPelaVaga)
              mensagem_responsavel = mensagem_responsavel.replace("[pessoa]", nome)
              mensagem_responsavel = mensagem_responsavel.replace("[função]", funcao)
              adicionar_data("D",index+1)
              if (responsavelPelaVaga == 'Daniel') {
                valores.push([[11985848901], [mensagem_responsavel]]) 
              }
              
              if (responsavelPelaVaga == 'Stella') {
                valores.push([[11972870027], [mensagem_responsavel]]) 
              }

              if (responsavelPelaVaga == 'Beatriz') {
                valores.push([[11996773018], [mensagem_responsavel]]) 
              }

              if (responsavelPelaVaga == 'Claudio') {
                valores.push([[11981691919], [mensagem_responsavel]]) 
              }

              if (responsavelPelaVaga == 'Elaine') {
                valores.push([[11984057064], [mensagem_responsavel]]) 
              }

              if (responsavelPelaVaga == 'Rafaeli') {
                valores.push([[11994077602], [mensagem_responsavel]]) 
              }


              // Fim do trecho de codigo
             
          }

          if (row[4] == 'Em Análise') {
          let data_planilha = row[5]
          let currentDate = new Date();
            // Obtendo o dia, mês e ano
          let dia = currentDate.getDate();
          let mes = currentDate.getMonth() + 1; // Adiciona +1 porque os meses são indexados a partir de zero
          let ano = currentDate.getFullYear();

          // Formatando a data
          if (dia < 10) {
              dia = '0' + dia; // Adiciona um zero à esquerda se o dia for menor que 10
          }

          if (mes < 10) {
              mes = '0' + mes; // Adiciona um zero à esquerda se o mês for menor que 10
          }

          let dataFormatada = dia + '/' + mes + '/' + ano;



            let dataArray = data_planilha.split("/");
            let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
            let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
            if (diferenca > 15) {
            //   console.log("A data fornecida está mais de 7 dias no passado");
                nome = row[7]
                numero = row[9]

                coluna = "F"
                adicionar_data(coluna,index+1)


                //codigo da parte de mandar mensagem para o responsavel
                responsavelPelaVaga = row[14]
                funcao = row[11]
                mensagem_responsavel = `Oi, [nomeresponsavel]. Tudo certinho? Aqui é o P1P4🤖🪁 ... Tô passando pra lembrar que a [pessoa] da função [função], está com o status "Em Análise" já faz um tempinho... você poderia verificar se há algo a se fazer?`

                mensagem_responsavel = mensagem_responsavel.replace("[nomeresponsavel]", responsavelPelaVaga)
                mensagem_responsavel = mensagem_responsavel.replace("[pessoa]", nome)
                mensagem_responsavel = mensagem_responsavel.replace("[função]", funcao)
                
                if (responsavelPelaVaga == 'Daniel') {
                  valores.push([[11985848901], [mensagem_responsavel]]) 
                }
                
                if (responsavelPelaVaga == 'Stella') {
                  valores.push([[11972870027], [mensagem_responsavel]]) 
                }
  
                if (responsavelPelaVaga == 'Beatriz') {
                  valores.push([[11996773018], [mensagem_responsavel]]) 
                }

                 if (responsavelPelaVaga == 'Claudio') {
                  valores.push([[11981691919], [mensagem_responsavel]]) 
                }

                if (responsavelPelaVaga == 'Elaine') {
                  valores.push([[11984057064], [mensagem_responsavel]]) 
                }
                if (responsavelPelaVaga == 'Rafaeli') {
                  valores.push([[11994077602], [mensagem_responsavel]]) 
                }




            }
          }





          
          if (row[2] == 'Em Análise') {
          let data_planilha = row[3]
          let currentDate = new Date();
            // Obtendo o dia, mês e ano
          let dia = currentDate.getDate();
          let mes = currentDate.getMonth() + 1; // Adiciona +1 porque os meses são indexados a partir de zero
          let ano = currentDate.getFullYear();

          // Formatando a data
          if (dia < 10) {
              dia = '0' + dia; // Adiciona um zero à esquerda se o dia for menor que 10
          }

          if (mes < 10) {
              mes = '0' + mes; // Adiciona um zero à esquerda se o mês for menor que 10
          }

          let dataFormatada = dia + '/' + mes + '/' + ano;



            let dataArray = data_planilha.split("/");
            let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
            let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
            if (diferenca > 15) {
            //   console.log("A data fornecida está mais de 7 dias no passado");
                nome = row[7]
                numero = row[9]

                coluna = "D"
                adicionar_data(coluna,index+1)


                //codigo da parte de mandar mensagem para o responsavel
                responsavelPelaVaga = row[14]
                funcao = row[11]
                mensagem_responsavel = `Oi, [nomeresponsavel]. Tudo certinho? Aqui é o P1P4🤖🪁 ... Tô passando pra lembrar que a [pessoa] da função [função], está com o status "Em Análise" já faz um tempinho... você poderia verificar se há algo a se fazer?`

                mensagem_responsavel = mensagem_responsavel.replace("[nomeresponsavel]", responsavelPelaVaga)
                mensagem_responsavel = mensagem_responsavel.replace("[pessoa]", nome)
                mensagem_responsavel = mensagem_responsavel.replace("[função]", funcao)
                
                if (responsavelPelaVaga == 'Daniel') {
                  valores.push([[11985848901], [mensagem_responsavel]]) 
                }
                
                if (responsavelPelaVaga == 'Stella') {
                  valores.push([[11972870027], [mensagem_responsavel]]) 
                }
  
                if (responsavelPelaVaga == 'Beatriz') {
                  valores.push([[11996773018], [mensagem_responsavel]]) 
                }

                 if (responsavelPelaVaga == 'Claudio') {
                  valores.push([[11981691919], [mensagem_responsavel]]) 
                }

                if (responsavelPelaVaga == 'Elaine') {
                  valores.push([[11984057064], [mensagem_responsavel]]) 
                }
                if (responsavelPelaVaga == 'Rafaeli') {
                  valores.push([[11994077602], [mensagem_responsavel]]) 
                }


            }
          }















          if (row[2] == "Agendada") {
            nome = row[7]
            numero = row[9]
            mensagem_agendada = row[53]
            mensagem_agendada = mensagem_agendada.replace("[primeiro nome]", nome.split(" ")[0])

            let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)

            const numeroAlterado = removerDigitoTelefone(numero);
            valores.push([[numeroAlterado],[mensagem_agendada]])
            valores.push([[resultado1],[mensagem_agendada]])
            valores.push([[resultado2],[mensagem_agendada]])

           
        
            adicionar_texto("C", index + 1, "Lembrete anotar agendamento e link enviada")
        }
            if (row[4] == "Marcar entrevista" && row[3] == '') {
                nome = row[7]
                numero = row[9]
                mensagem4 = row[19]
                mensagem4 = mensagem4.replace("[primeiro nome]", nome.split(" ")[0])

                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)

                const numeroAlterado = removerDigitoTelefone(numero);
                valores.push([[numeroAlterado],[mensagem4]])
                valores.push([[resultado1],[mensagem4]])
                valores.push([[resultado2],[mensagem4]])

                coluna = "D"
                adicionar_data(coluna,index + 1)   
                adicionar_texto("C", index + 1, "Aguardando agendamento")
            }

            if (row[4] == "Incompatível com a vaga" && row[0] == '') {
                nome = row[7]
                numero = row[9]
                motivo = row[52]
                mensagem5 = row[20]
                mensagem5 = mensagem5.replace("[primeiro nome]", nome.split(" ")[0])
                mensagem5 = mensagem5.replace("[motivo]", motivo)



                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                const numeroAlterado = removerDigitoTelefone(numero);
                valores.push([[numeroAlterado],[mensagem5]])
                valores.push([[resultado1],[mensagem5]])
                valores.push([[resultado2],[mensagem5]])

                coluna = "A"
                texto = 'Reprovado(a) e informado(a)'
                adicionar_texto(coluna, index + 1, texto)
              }





              if (row[5] != '' && row[4] == '') {
                let data_planilha = row[5]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 7) {
                //   console.log("A data fornecida está mais de 7 dias no passado");
                    nome = row[7]
                    numero = row[9]
                    mensagem6 = row[21]
                    texto = "Primeira semana sem responder"
                    coluna = "E"
                    mensagem6 = mensagem6.replace("[primeiro nome]", nome.split(" ")[0])
                    adicionar_texto(coluna, index + 1, texto)

                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    const numeroAlterado = removerDigitoTelefone(numero);
                    valores.push([[numeroAlterado], [mensagem6]])
                    valores.push([[resultado1],[mensagem6]])
                    valores.push([[resultado2],[mensagem6]])

                //   linha_adicionar_cinza.push(index + 1)
                //   console.log(valores_cinza)
                } 
              }
              
              if (row[5] != '' && row[4] == 'Primeira semana sem responder') {
                let data_planilha = row[5]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 14) {
                //   console.log("A data fornecida está mais de 21 dias no passado");
                    nome = row[7]
                    numero = row[9]
                    // mensagem7 = row[22]
                    texto = "E-mail enviado"
                    coluna = "E"
                    // mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])
                    adicionar_texto(coluna, index + 1, texto)
                    console.log('Alterei', index + 1)
                    console.log('Nome da vaga', row[11])
                    textos = mensagem_email.data.values;
                    // console.log(textos[index])
                   
                    console.log(row[8])
                    // mandar_email(nome,textos[index],row[11],numero,row[8])
              
                   
                } 
              }







              if (row[5] != '' && row[4] == 'E-mail enviado') {
                let data_planilha = row[5]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 21) {
                //   console.log("A data fornecida está mais de 21 dias no passado");
                    nome = row[7]
                    numero = row[9]
                    mensagem7 = row[22]
                    texto = "Terceira semana sem responder"
                    coluna = "E"
                    mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])
                    adicionar_texto(coluna, index + 1, texto)

                    adicionar_texto('A', index + 1, 'Duas tentativas sem resposta')
                    const numeroAlterado = removerDigitoTelefone(numero);
                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    
                    valores.push([[numeroAlterado], [mensagem7]])
                    valores.push([[resultado1],[mensagem7]])
                    valores.push([[resultado2],[mensagem7]])

                } 
              }

              if (row[2] == "Realizada" && row[0] == 'Aprovado(a)' && row[1] == '') {
                nome = row[7]
                numero = row[9]
                mensagem8 = row[23]
                mensagem8 = mensagem8.replace("[Primeiro nome]", nome.split(" ")[0])
                valores.push([[numero],[mensagem8]])
                adicionar_data('B', index + 1)
                adicionar_texto("A", index + 1, "Responder formulário adesão")
              }
              if (row[2] == "Realizada" && row[0] == 'Responder formulário adesão') {
                let data_planilha = row[1]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 7) {
                    numero = row[9]
                    nome = row[7]
                    mensagem_para_termo_adesao = row[28]
                    console.log(mensagem_para_termo_adesao)
                    mensagem_para_termo_adesao = mensagem_para_termo_adesao.replace("[Primeiro nome]", nome.split(" ")[0])
                   
                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    const numeroAlterado = removerDigitoTelefone(numero);
                    valores.push([[numeroAlterado], [mensagem_para_termo_adesao]])
                    valores.push([[resultado1],[mensagem_para_termo_adesao]])
                    valores.push([[resultado2],[mensagem_para_termo_adesao]])
  
                    adicionar_data_termo_adesao("B", index + 1, '¹')
                }



              }

              if (row[2] == "Aguardando agendamento" && row[3] != '') { 
                let data_planilha = row[3]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 7) {
                    numero = row[9]
                    nome = row[7]
                    mensagem9 = row[24]
                    mensagem9 = mensagem9.replace("[Primeiro nome]", nome.split(" ")[0])
                    mensagem9 = mensagem9.replace("[primeiro nome]", nome.split(" ")[0])


                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    const numeroAlterado = removerDigitoTelefone(numero);
                    valores.push([[numeroAlterado], [mensagem9]])
                    valores.push([[resultado1],[mensagem9]])
                    valores.push([[resultado2],[mensagem9]])

                    adicionar_texto("C", index + 1, "Mensagem enviada de aguardando agendamento")
                }
              }
              if (row[2] == 'Mensagem enviada de aguardando agendamento' && row[0] == ''){
                let data_planilha = row[3]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 14) {
                    numero = row[9]
                    nome = row[7]
                    mensagem9 = row[24]
                    mensagem9 = mensagem9.replace("[Primeiro nome]", nome.split(" ")[0])
                    mensagem9 = mensagem9.replace("[primeiro nome]", nome.split(" ")[0])
                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    const numeroAlterado = removerDigitoTelefone(numero);
                    valores.push([[numeroAlterado], [mensagem9]])
                    valores.push([[resultado1],[mensagem9]])
                    valores.push([[resultado2],[mensagem9]])

                    adicionar_texto("C", index + 1, "Segunda mensagem de aguardando agendamento enviada")
                }
              }

              if (row[2] == 'Segunda mensagem de aguardando agendamento enviada' && row[0] == ''){
                let data_planilha = row[3]
                let currentDate = new Date();
                let dataArray = data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 21) {
                    numero = row[9]
                    nome = row[7]
                    mensagem7 = row[22]
                    mensagem7 = mensagem7.replace("[Primeiro nome]", nome.split(" ")[0])
                    mensagem7 = mensagem7.replace("[primeiro nome]", nome.split(" ")[0])


                    let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                    const numeroAlterado = removerDigitoTelefone(numero);
                    valores.push([[numeroAlterado], [mensagem7]])
                    valores.push([[resultado1],[mensagem7]])
                    valores.push([[resultado2],[mensagem7]])

                    adicionar_texto("A", index + 1, "Duas tentativas sem resposta")
                }

              }
              if (row[2] == "Faltou primeira vez") {
                nome = row[7]
                numero = row[9]
                mensagem10 = row[25]
                mensagem10 = mensagem10.replace("[Primeiro nome]", nome.split(" ")[0])
                mensagem10 = mensagem10.replace("[primeiro nome]", nome.split(" ")[0])

                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                const numeroAlterado = removerDigitoTelefone(numero);
                valores.push([[numeroAlterado], [mensagem10]])
                valores.push([[resultado1],[mensagem10]])
                valores.push([[resultado2],[mensagem10]])

                adicionar_texto("C", index + 1, "Segunda tentativa de agendamento")    
                adicionar_data('D', index + 1)
            }
            if (row[2] == "Segunda tentativa de agendamento"  && row[0] == '') {
              let data_planilha = row[3]
              let currentDate = new Date();
              let dataArray = data_planilha.split("/");
              let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
              let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
              if (diferenca > 14) {
                  numero = row[9]
                  nome = row[7]
                  mensagem_termo_assinar = row[22]
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])
                  let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                  const numeroAlterado = removerDigitoTelefone(numero);
                  valores.push([[numeroAlterado], [mensagem_termo_assinar]])
                  valores.push([[resultado1],[mensagem_termo_assinar]])
                  valores.push([[resultado2],[mensagem_termo_assinar]])

                  adicionar_texto("A", index + 1, "Duas tentativas sem resposta")   
                  mandar_email_duas_tentaivas_sem_resposta(nome,row[8])
              }
            }


            if (row[2] == "Faltou segunda vez" && row[0] != 'Duas faltas no agendamento') {
                nome = row[7]
                numero = row[9]
                mensagem11 = row[26]
                mensagem11 = mensagem11.replace("[Primeiro nome]", nome.split(" ")[0])
                mensagem11 = mensagem11.replace("[primeiro nome]", nome.split(" ")[0])

                let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                const numeroAlterado = removerDigitoTelefone(numero);
                valores.push([[numeroAlterado], [mensagem11]])
                valores.push([[resultado1],[mensagem11]])
                valores.push([[resultado2],[mensagem11]])

                adicionar_texto("A", index + 1, "Duas faltas no agendamento")    
            }

            if(row[0] == 'P1P4 Enviar mensagem de termo de adesão'){
              numero = row[9]
              nome = row[7]
              mensagem_termo_assinar = 'Oi, [Primeiro nome], agradecemos por ter fornecido as informações. Acabei de enviar para o e-mail informado o termo para ser assinado digitalmente. É só seguir o passo a passo. É super simples, mas se tiver alguma dúvida, pode me chamar que eu respondo em breve'
              mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
              mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])


              let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
              const numeroAlterado = removerDigitoTelefone(numero);
              valores.push([[numeroAlterado], [mensagem_termo_assinar]])
              valores.push([[resultado1],[mensagem_termo_assinar]])
              valores.push([[resultado2],[mensagem_termo_assinar]])

              adicionar_texto('A', index + 1, 'Termo adesão enviado')
            }




            // está com Termo de adessão assinado e com uma data 7 dias no passado. 
            if (row[0] == "Termo adesão enviado") {
              let data_planilha = row[1]
              let currentDate = new Date();
              let dataArray = data_planilha.split("/");
              let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
              let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
              if (diferenca > 7) {
                  numero = row[9]
                  nome = row[7]
                  mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])
                  let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                  const numeroAlterado = removerDigitoTelefone(numero);
                  valores.push([[numeroAlterado], [mensagem_termo_assinar]])
                  valores.push([[resultado1],[mensagem_termo_assinar]])
                  valores.push([[resultado2],[mensagem_termo_assinar]])

                  adicionar_data_termo_adesao("B", index + 1, '¹')
              }
              if (row[0] == "Termo adesão enviado" && row[1].charAt() == '¹'){
                
                let data_planilha = row[1]
                nova_data_planilha = data_planilha.substring(1);
                let currentDate = new Date();
                let dataArray = nova_data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 7) {
                  numero = row[9]
                  nome = row[7]
                  mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])
                  let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                  const numeroAlterado = removerDigitoTelefone(numero);
                  valores.push([[numeroAlterado], [mensagem_termo_assinar]])
                  valores.push([[resultado1],[mensagem_termo_assinar]])
                  valores.push([[resultado2],[mensagem_termo_assinar]])

                  adicionar_data_termo_adesao("B", index + 1, '²')
                }

              }
              if (row[0] == "Termo adesão enviado" && row[1].charAt() == '²'){
                
                let data_planilha = row[1]
                nova_data_planilha = data_planilha.substring(1);
                let currentDate = new Date();
                let dataArray = nova_data_planilha.split("/");
                let novaData = new Date(dataArray[2], dataArray[1] - 1, dataArray[0]);
                let diferenca = Math.floor((currentDate.getTime() - novaData.getTime()) / (1000 * 3600 * 24));
                if (diferenca > 7) {
                  numero = row[9]
                  nome = row[7]
                  mensagem_termo_assinar = 'Olá, [Primeiro nome], espero que esteja bem... Vi aqui que você ainda não assinou o Termo de Adesão ao Voluntariado. Esse documento é super importante para nós. Você teve alguma dificuldade? Qualquer coisa, me chama aqui...'
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[Primeiro nome]", nome.split(" ")[0])
                  mensagem_termo_assinar = mensagem_termo_assinar.replace("[primeiro nome]", nome.split(" ")[0])


                  let { resultado1, resultado2} = duplicanumerosporcausadonove(numero)
                  const numeroAlterado = removerDigitoTelefone(numero);
                  valores.push([[numeroAlterado], [mensagem_termo_assinar]])
                  valores.push([[resultado1],[mensagem_termo_assinar]])
                  valores.push([[resultado2],[mensagem_termo_assinar]])

                  adicionar_data('B', index + 1)
                  adicionar_texto('A', index + 1, 'ADM Contatar')
                }

              }

                   
               
              // nome = row[7]
              // numero = row[9]
              // mensagem11 = row[26]
              // mensagem11 = mensagem11.replace("[Primeiro nome]", nome.split(" ")[0])
              // mensagem11 = mensagem11.replace("[primeiro nome]", nome.split(" ")[0])
              // valores.push([[numero],[mensagem11]])
              // adicionar_texto("A", index + 1, "Duas faltas no agendamento")    
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
            spreadsheetId: '1Jnl_PqlDJRxemLOlDP2aFawoDNo9EHG_Ma43ZvcfOyY',
            range: `Principal!${coluna+linha}`,
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
              spreadsheetId: '1Jnl_PqlDJRxemLOlDP2aFawoDNo9EHG_Ma43ZvcfOyY',
              range: `Principal!${coluna+linha}`,
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
            spreadsheetId: '1Jnl_PqlDJRxemLOlDP2aFawoDNo9EHG_Ma43ZvcfOyY',
            range: `Principal!${coluna+linha}`,
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
        valores.push([['11985848901'], [`P1P4 responsável pela parte de envio da planilha "Vagas e Candidaturas" rodando, porém não existe mensagens para serem enviadas:  ${getCurrentDateTimeBrazilian()} \n`]])
        valores.push([['11945274604'], [`P1P4 responsável pela parte de envio da planilha "Vagas e Candidaturas" rodando, porém não existe mensagens para serem enviadas:  ${getCurrentDateTimeBrazilian()} \n`]])
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






console.log('FINALLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLL')
// console.log(sheets())
// // whats(array_mensagens)



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
function mandar_email(nome, mensagem, vaga, telefone, email){
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
  mensagem = mensagem.replace("[primeiro nome]", nome.split(" ")[0])
  mensagem = mensagem.replace("[número telefone informado]", telefone)
  titulo = "Vaga de Voluntariado - [ Nome da vaga ]"
  titulo = titulo.replace("[ Nome da vaga ]", vaga)
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








function mandar_email_inicial(nome, telefone, email, mensagem){
  


  // colocar a mensagem da coluna p da base dialogos
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
  mensagem = mensagem.replace("[primeiro nome]", nome.split(" ")[0])
  mensagem = mensagem.replace("[Nome da pessoa]", nome.split(" ")[0])
  mensagem = mensagem.replace("[número celular]", telefone)
  mensagem = mensagem.replace("[número telefone informado]", telefone)


  titulo = "Voluntariado - Vai voar no PiPA com a gente?"
    // send mail with defined transport object
  // let info = await 
  transporter.sendMail({
  from: '"PiPA 🪁" <opipa@opipa.org>', // sender address
  to: email, // list of receivers
  subject: titulo, // Subject line
  text: `${mensagem}`, // plain text body
  // html: `${mensagem}`, // html body
  })
  .then(() => console.log('Email Enviado'))
  .catch((err) => console.log('Erro ao enviar o email', err))
}


function mandar_email_duas_tentaivas_sem_resposta(nome, email){
  mensagem = `Oi, [Nome da pessoa] aqui é da Associação PiPA🪁 

  Como não obtive respostas, estou finalizando sua candidatura. Caso ainda tenha interesse em participar conosco, veja as vagas de voluntariado abertas no momento em: https://atados.com.br/ong/pipa/vagas e se inscreva no processo novamente.
  
  Um abraço,
  
  Equipe PiPA 🪁`



  
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
  mensagem = mensagem.replace("[Nome da pessoa]", nome.split(" ")[0])

  titulo = "Voluntariado - Vai voar no PiPA com a gente?"
    // send mail with defined transport object
  // let info = await 
  transporter.sendMail({
  from: '"PiPA 🪁" <opipa@opipa.org>', // sender address
  to: email, // list of receivers
  subject: titulo, // Subject line
  text: `${mensagem}`, // plain text body
  // html: `${mensagem}`, // html body
  })
  .then(() => console.log('Email Enviado'))
  .catch((err) => console.log('Erro ao enviar o email', err))
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



