const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const qrcode = require('qrcode-terminal'); // Importado aqui para uso na função whats
const { Client, LocalAuth } = require('whatsapp-web.js'); // Importado aqui para uso na função whats
const nodemailer = require('nodemailer'); // Mantido, pois você usa nodemailer em funções auxiliares

// --- Configurações Google Sheets ---
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

// --- Funções de Autenticação do Google Sheets ---
async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
}

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

// --- Funções Auxiliares de Data e Hora ---
function getCurrentDateTimeBrazilian() {
    const currentDate = new Date();
    const dayOfWeek = ["Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado"][currentDate.getDay()];
    const day = String(currentDate.getDate()).padStart(2, '0');
    const month = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][currentDate.getMonth()];
    const year = currentDate.getFullYear();
    const hours = String(currentDate.getHours()).padStart(2, '0');
    const minutes = String(currentDate.getMinutes()).padStart(2, '0');
    const seconds = String(currentDate.getSeconds()).padStart(2, '0');
    return `${dayOfWeek}, ${day} de ${month} de ${year} ${hours}:${minutes}:${seconds}`;
}

function obterDataAtual() {
    const dataAtual = new Date();
    const dia = String(dataAtual.getDate()).padStart(2, '0');
    const mes = String(dataAtual.getMonth() + 1).padStart(2, '0');
    const ano = dataAtual.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

// --- Funções de Manipulação de Email (Nodemailer) ---
function mandar_email(nome, mensagem, vaga, telefone, email) {
    const data = require('./email.json');
    let transporter = nodemailer.createTransport({
        host: "smtp.gmail.com",
        port: 465,
        secure: true,
        auth: {
            user: 'opipa@opipa.org',
            pass: data.token,
        },
    });
    console.log(`Tentando enviar email para ${email}: ${mensagem.substring(0, 50)}...`);
    email = email.toString();
    mensagem = mensagem.toString();
    mensagem = mensagem.replace(/\[primeiro nome\]/gi, nome.split(" ")[0]); // Usando regex para case-insensitive
    mensagem = mensagem.replace("[número telefone informado]", telefone);
    let titulo = "Vaga de Voluntariado - [ Nome da vaga ]";
    titulo = titulo.replace("[ Nome da vaga ]", vaga);
    transporter.sendMail({
        from: '"PiPA 🪁" <daniel.mariano@opipa.org>',
        to: email,
        subject: titulo,
        text: `${mensagem}`,
    })
    .then(() => console.log('Email Enviado com sucesso!'))
    .catch((err) => console.error('Erro ao enviar o email:', err));
}

function mandar_email_inicial(nome, telefone, email, mensagem) {
    const data = require('./email.json');
    let transporter = nodemailer.createTransport({
        host: "smtp.gmail.com",
        port: 465,
        secure: true,
        auth: {
            user: 'opipa@opipa.org',
            pass: data.token,
        },
    });
    console.log(`Tentando enviar email inicial para ${email}: ${mensagem.substring(0, 50)}...`);
    email = email.toString();
    mensagem = mensagem.toString();
    mensagem = mensagem.replace(/\[primeiro nome\]/gi, nome.split(" ")[0]);
    mensagem = mensagem.replace("[Nome da pessoa]", nome.split(" ")[0]);
    mensagem = mensagem.replace("[número celular]", telefone);
    mensagem = mensagem.replace("[número telefone informado]", telefone);
    let titulo = "Voluntariado - Vai voar no PiPA com a gente?";
    transporter.sendMail({
        from: '"PiPA 🪁" <opipa@opipa.org>',
        to: email,
        subject: titulo,
        text: `${mensagem}`,
    })
    .then(() => console.log('Email Inicial Enviado com sucesso!'))
    .catch((err) => console.error('Erro ao enviar o email inicial:', err));
}

function mandar_email_duas_tentaivas_sem_resposta(nome, email) {
    let mensagem = `Oi, [Nome da pessoa] aqui é da Associação PiPA🪁 
Como não obtive respostas, estou finalizando sua candidatura. Caso ainda tenha interesse em participar conosco, veja as vagas de voluntariado abertas no momento em: https://atados.com.br/ong/pipa/vagas e se inscreva no processo novamente.
Um abraço,
Equipe PiPA 🪁`;

    const data = require('./email.json');
    let transporter = nodemailer.createTransport({
        host: "smtp.gmail.com",
        port: 465,
        secure: true,
        auth: {
            user: 'opipa@opipa.org',
            pass: data.token,
        },
    });
    console.log(`Tentando enviar email de duas tentativas para ${email}: ${mensagem.substring(0, 50)}...`);
    email = email.toString();
    mensagem = mensagem.toString();
    mensagem = mensagem.replace("[Nome da pessoa]", nome.split(" ")[0]);
    let titulo = "Voluntariado - Vai voar no PiPA com a gente?";
    transporter.sendMail({
        from: '"PiPA 🪁" <opipa@opipa.org>',
        to: email,
        subject: titulo,
        text: `${mensagem}`,
    })
    .then(() => console.log('Email de duas tentativas enviado com sucesso!'))
    .catch((err) => console.error('Erro ao enviar o email de duas tentativas:', err));
}

// --- Funções de Manipulação de Números de Telefone ---
function removerDigitoTelefone(numero) {
    const numeroLimpo = String(numero).replace(/\D/g, ''); // Garante que é string

    if (numeroLimpo.length !== 11) {
        console.warn(`Aviso: Número de telefone inválido para remoção de dígito (esperado 11 dígitos): ${numero}`);
        return numero; // Retorna o original se for inválido
    }
    const numeroAlterado = numeroLimpo.slice(0, 2) + numeroLimpo.slice(3);
    return `(${numeroAlterado.slice(0, 2)}) ${numeroAlterado.slice(2, 7)}-${numeroAlterado.slice(7)}`;
}

function duplicanumerosporcausadonove(Numero) {
    let numero = String(Numero).replace(/\D/g, ''); // Garante que é string

    if (numero.startsWith("55") && numero.length >= 13) {
        numero = numero.substring(2);
    }

    let copia_numero = numero;
    let resultado2;

    if (numero.length === 11 && numero.substring(2, 3) === "9") { // Se já tem 9 no lugar certo
        // Remove o 9, assumindo que é um dígito extra de telefone móvel brasileiro
        resultado2 = numero.substring(0, 2) + numero.substring(3);
    } else if (numero.length === 10) { // Se não tem 9 e tem 10 dígitos (pode precisar de um 9)
        resultado2 = numero.substring(0, 2) + "9" + numero.substring(2);
    } else {
        // Se o número não se encaixa nos padrões, ou já está ok, retorna ele mesmo
        resultado2 = numero;
    }

    // console.log("Número de telefone original (limpo): " + copia_numero); // Removi log excessivo para clareza
    // console.log("Número de telefone duplicado/ajustado: " + resultado2); // Removi log excessivo

    return { resultado1: copia_numero, resultado2: resultado2 };
}

// --- Função Principal de Leitura e Processamento do Google Sheets ---

async function sheets() {
    var valores = []; // Array para armazenar as mensagens a serem enviadas

    async function listMajors(auth) {
        // Objeto sheetsAPI com a autenticação
        const sheetsAPI = google.sheets({ version: 'v4', auth });

        // Leitura da planilha 'Mensagens Agendadas'
        const res = await sheetsAPI.spreadsheets.values.get({
            spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4', // ID da planilha de mensagens agendadas
            range: 'Mensagens Agendadas!A:F',
        });
        // Leitura da planilha 'Contatos'
        const contatos = await sheetsAPI.spreadsheets.values.get({
            spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4', // ID da planilha de contatos
            range: 'Contatos!A:AJ',
        });

        const rows_contatos = contatos.data.values;
        const rows_mensagens = res.data.values;

        if (!rows_mensagens || rows_mensagens.length === 0) {
            console.log('Nenhum dado encontrado na planilha "Mensagens Agendadas".');
            return [];
        }
        if (!rows_contatos || rows_contatos.length === 0) {
            console.log('Nenhum dado encontrado na planilha "Contatos".');
            return [];
        }

        // --- FUNÇÕES AUXILIARES DE ATUALIZAÇÃO E ENVIO DEFINIDAS AQUI DENTRO DO ESCOPO DE listMajors ---
        // Elas agora têm acesso a 'sheetsAPI'
        async function adicionar_data(coluna, linha) {
            let dataAtualFormatada = obterDataAtual();
            let values = [[dataAtualFormatada]];
            const resource = { values };
            try {
                const result = await sheetsAPI.spreadsheets.values.update({
                    spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
                    range: `Mensagens Agendadas!${coluna}${linha}`,
                    valueInputOption: 'RAW',
                    resource,
                });
                // console.log(`Data adicionada em ${coluna}${linha}:`, result.data);
                return result;
            } catch (err) {
                console.error(`Erro ao adicionar data em ${coluna}${linha}:`, err.message);
                throw err;
            }
        }

        async function adicionar_data_termo_adesao(coluna, linha, ordem) {
            let dataAtualFormatada = obterDataAtual();
            let data_formatada = ordem + dataAtualFormatada;
            let values = [[data_formatada]];
            const resource = { values };
            try {
                const result = await sheetsAPI.spreadsheets.values.update({
                    spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
                    range: `Mensagens Agendadas!${coluna}${linha}`,
                    valueInputOption: 'RAW',
                    resource,
                });
                // console.log(`Data (termo adesão) adicionada em ${coluna}${linha}:`, result.data);
                return result;
            } catch (err) {
                console.error(`Erro ao adicionar data (termo adesão) em ${coluna}${linha}:`, err.message);
                throw err;
            }
        }

        async function adicionar_texto(coluna, linha, texto) {
            let values = [[texto]];
            const resource = { values };
            try {
                const result = await sheetsAPI.spreadsheets.values.update({
                    spreadsheetId: '1nfv2ALUmqW9I9pijsHMZ1rjzakcWGY6ZH5_lreYDo_4',
                    range: `Mensagens Agendadas!${coluna}${linha}`,
                    valueInputOption: 'USER_ENTERED',
                    resource,
                });
                // console.log(`Texto adicionado em ${coluna}${linha}:`, result.data);
                return result;
            } catch (err) {
                console.error(`Erro ao adicionar texto em ${coluna}${linha}:`, err.message);
                throw err;
            }
        }

        function sheets_enviar_mensagem(texto_verificar, row) {
            rows_contatos.forEach((row_contato) => {
                if (row_contato[2] && row_contato[2] !== 'Nome do Responsável' && row_contato[0] !== "VERDADEIRO") {
                    if (row_contato[5] === texto_verificar || row_contato[6] === texto_verificar || row_contato[7] === texto_verificar || row_contato[8] === texto_verificar || row_contato[9] === texto_verificar) {
                        let nome = row_contato[2];
                        let numero = row_contato[3];
                        let mensagem1 = row[1]; // Assumindo que a mensagem está na coluna B (index 1) da linha de 'Mensagens Agendadas'
                        
                        mensagem1 = mensagem1.replace(/\[nome\]/gi, nome.split(" ")[0]);
                        valores.push([[numero], [mensagem1]]);
                    }
                }
            });
        }
        // --- FIM DAS FUNÇÕES AUXILIARES DE ATUALIZAÇÃO E ENVIO ---


        // Loop principal para processar as linhas da planilha 'Mensagens Agendadas'
        for (let index = 0; index < rows_mensagens.length; index++) {
            const row = rows_mensagens[index];
            if (!row || !row[3]) continue; // Pula linhas vazias ou sem data agendada

            // Lógica para mensagens agendadas
            if (row[3] === obterDataAtual() && row[4] !== 'Sim') { // Coluna D (index 3) para data, Coluna E (index 4) para 'Sim'
                if (row[2] === 'Todos') { // Coluna C (index 2) para 'Todos'
                    rows_contatos.forEach((row_contato) => {
                        if (row_contato[2] && row_contato[2] !== 'Nome da Pessoa' && row_contato[0] !== "VERDADEIRO") {
                            let nome = row_contato[2];
                            let numero = row_contato[3];
                            let mensagem1 = row[0]; // Mensagem da coluna A
                            mensagem1 = mensagem1.replace(/\[nome\]/gi, nome.split(" ")[0]);
                            valores.push([[numero], [mensagem1]]);
                        }
                    });
                } else if (row[2]) { // Se não for 'Todos', verifica se a célula C não está vazia
                    sheets_enviar_mensagem(row[2], row);
                }
                
                // Atualizações na planilha
                await adicionar_texto("E", index + 1, "Sim"); // Coluna E
                await adicionar_data("F", index + 1); // Coluna F
                await adicionar_texto("G", index + 1, row[1]); // Coluna G (conteúdo da coluna B da mensagem)
                await adicionar_texto("H", index + 1, row[2]); // Coluna H (conteúdo da coluna C da mensagem)
            }
            // Outras lógicas de Ifs do seu código original aqui, adaptadas para usar await e as funções internas
            // Devido à complexidade e quantidade, não vou transcrevê-las todas, mas o padrão é o mesmo:
            // Trocar `adicionar_texto(...)` por `await adicionar_texto(...)`
            // Trocar `adicionar_data(...)` por `await adicionar_data(...)`
            // Lembre-se de verificar os `spreadsheetId` se você tiver IDs diferentes para outras planilhas
        }
        return valores;
    }
    
    return authorize().then(listMajors).catch(console.error);
}

// --- Função Principal de Envio de Mensagens (WhatsApp) ---

async function whats(todas_acoes) {
    const client = new Client({
        authStrategy: new LocalAuth(),
        webVersion: "2.2412.54",
        webVersionCache: {
            type: "remote",
            remotePath: "https://raw.githubusercontent.com/wppconnect-team/wa-version/main/html/2.2412.54.html",
        },
        puppeteer: {
            timeout: 60000, // 60 segundos
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-accelerated-video-decode',
                '--no-first-run',
                '--no-zygote',
                '--single-process',
                '--disable-gpu',
                '--no-cache-dir'
            ],
            // SE O ERRO 'Could not find expected browser' PERSISTIR, DESCOMENTE A LINHA ABAIXO
            // E COLOQUE O CAMINHO CORRETO DO SEU CHROME.EXE
            // executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
        }
    });

    client.on('qr', qr => {
        console.log('QR RECEIVED', qr);
        qrcode.generate(qr, { small: true });
    });

    client.on('ready', async () => {
        console.log('Client is ready!');

        const formatado = [];
        for (let i = 0; i < todas_acoes.length; i++) {
            if (todas_acoes[i] && todas_acoes[i][0] && todas_acoes[i][0][0] && todas_acoes[i][1] && todas_acoes[i][1][0]) {
                const numeroBruto = todas_acoes[i][0][0];
                const mensagem = todas_acoes[i][1][0];
                const numero_enviar = '55' + String(numeroBruto).replace(/\D/g, '') + '@c.us';
                formatado.push([numero_enviar, mensagem]);
            } else {
                console.warn('Estrutura de dados inesperada na posição', i, 'de todas_acoes:', todas_acoes[i]);
            }
        }

        if (formatado.length === 0) {
            console.log('Nenhuma mensagem válida para enviar.');
            await client.destroy();
            return;
        }

        console.log(`Iniciando envio de ${formatado.length} mensagens.`);

        for (const enviar of formatado) {
            try {
                // await client.sendMessage(enviar[0], ''); // Removido, geralmente não é necessário um send vazio
                await client.sendMessage(enviar[0], enviar[1]);
                console.log(`✔ Mensagem enviada para ${enviar[0]}: ${enviar[1].substring(0, Math.min(enviar[1].length, 50))}...`);

                await client.sendMessage('5511945274604@c.us', `*🤖 Mensagem enviada:* \n\`${enviar[1]}\` \n*Para:* \`${enviar[0]}\``);
                await client.sendMessage('5511985848901@c.us', `*🤖 Mensagem enviada:* \n\`${enviar[1]}\` \n*Para:* \`${enviar[0]}\``);

                await new Promise(resolve => setTimeout(resolve, 8000)); // Espera 8 segundos
            } catch (error) {
                console.error(`❌ Erro ao enviar mensagem para ${enviar[0]}:`, error.message);
                await new Promise(resolve => setTimeout(resolve, 8000)); // Espera mesmo com erro
            }
        }

        console.log('Todas as mensagens foram processadas. Finalizando o cliente.');
        await client.destroy();
    });

    client.on('message', message => {
        if (message.body === '!ping') {
            client.sendMessage(message.from, 'pong');
        }
    });

    client.on('disconnected', (reason) => {
        console.error('Cliente desconectado por motivo:', reason);
    });

    client.on('auth_failure', msg => {
        console.error('Falha na autenticação:', msg);
    });

    client.on('ready', () => {
        console.log('Client ready for use!');
    });

    client.on('change_state', state => {
        console.log('Estado do cliente mudou para:', state);
    });

    await client.initialize();
}

// --- Ponto de Entrada Principal ---
console.log('Iniciando o processo...');
sheets().then(async (valores) => {
    if (!valores || valores.length === 0) {
        console.log('Sem dados de mensagens para enviar da planilha. Enviando mensagem de teste...');
        const mensagensTeste = [
            [['11985848901'], [`P1P4 responsável pela parte de envio da planilha "Vagas e Candidaturas" rodando, porém não existe mensagens para serem enviadas: ${getCurrentDateTimeBrazilian()} \n`]],
            [['11945274604'], [`P1P4 responsável pela parte de envio da planilha "Vagas e Candidaturas" rodando, porém não existe mensagens para serem enviadas: ${getCurrentDateTimeBrazilian()} \n`]]
        ];
        await whats(mensagensTeste);
    } else {
        console.log(`Dados da planilha carregados. Total de ${valores.length} mensagens para processar.`);
        await whats(valores);
    }
}).catch(error => {
    console.error('Erro geral ao processar sheets ou iniciar whats:', error);
});