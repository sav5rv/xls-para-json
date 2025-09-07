const { MongoClient } = require('mongodb');
const xlsx = require('xlsx');
const fs = require('fs');

// --- Configuração do seu Banco de Dados ---
const uri = 'mongodb+srv://wsavioli:briGHT1590@cluster1.1ltdolk.mongodb.net'; // Altere para a URI do seu banco
const dbName = 'bd_num_rso';
const collectionName = 'rsos';

// --- Configuração dos arquivos ---
const excelFilePath = 'rso2025xls.xlsx'; // Nome do seu arquivo Excel
const jsonOutput = 'dados_ordenados.json';

// Função principal para a conversão
async function convertExcelToOrderedJson() {
    let client;
    try {
        // 1. Conectar ao MongoDB
        client = new MongoClient(uri);
        await client.connect();
        console.log('1. Conectado ao MongoDB com sucesso!');

        const db = client.db(dbName);
        const collection = db.collection(collectionName);

        // 2. Obter a ordem dos campos de um documento existente
        const sampleDoc = await collection.findOne({});
        if (!sampleDoc) {
            console.error('ERRO: Nenhum documento encontrado na coleção. A ordem dos campos não pode ser definida.');
            return;
        }
        const fieldOrder = Object.keys(sampleDoc).filter(key => key !== '_id');
        console.log('2. Ordem dos campos obtida do MongoDB:', fieldOrder.join(', '));
        if (fieldOrder.length === 0) {
            console.error('ERRO: A lista de campos do documento de exemplo está vazia.');
            return;
        }

        // 3. Ler o arquivo Excel com os valores brutos (raw: true)
        const workbook = xlsx.readFile(excelFilePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: true });
        console.log('3. Dados lidos do Excel.');

        if (excelData.length === 0) {
            console.error('ERRO: Nenhum dado encontrado no arquivo Excel. Verifique o nome da planilha ou se ela está vazia.');
            return;
        }

        // 4. Mapear e reordenar os dados do Excel
        const orderedData = excelData.map(row => {
            const newRow = {};

            // 4.1. Mapeia os dados do Excel para um objeto com chaves normalizadas
            const normalizedExcelRow = {};
            for (const key in row) {
                const normalizedKey = key.toLowerCase().trim();
                normalizedExcelRow[normalizedKey] = row[key];
            }
            
            // 4.2. Itera sobre a ordem dos campos do MongoDB
            fieldOrder.forEach(field => {
                const normalizedField = field.toLowerCase().trim();
                let value = normalizedExcelRow[normalizedField];

                // 4.3. Formata os campos de data e hora
                if (field.toLowerCase() === 'data') {
                    if (typeof value === 'number' && !isNaN(value)) {
                        const excelDate = new Date(Math.round((value - 25569) * 86400 * 1000));
                        value = excelDate.toISOString().split('T')[0]; // YYYY-MM-DD
                    } else {
                        value = null; // Mantém null para valores inválidos
                    }
                }

                if (field.toLowerCase() === 'hrinicio') {
                    if (typeof value === 'number' && !isNaN(value)) {
                        const excelTime = new Date(Math.round((value - 25569) * 86400 * 1000));
                        const hours = String(excelTime.getUTCHours()).padStart(2, '0');
                        const minutes = String(excelTime.getUTCMinutes()).padStart(2, '0');
                        value = `${hours}:${minutes}`; // HH:MM
                    } else {
                        value = null; // Mantém null para valores inválidos
                    }
                }
                
                // Atribui o valor ao novo objeto
                newRow[field] = value;
            });

            if (Object.keys(newRow).length === 0) {
                console.warn('AVISO: Uma linha do Excel não pôde ser mapeada. Verifique se os nomes das colunas do Excel correspondem aos do MongoDB.');
            }

            return newRow;
        });

        // 5. Salvar o arquivo JSON final
        fs.writeFileSync(jsonOutput, JSON.stringify(orderedData, null, 4));
        console.log(`Conversão concluída! O arquivo '${jsonOutput}' foi criado com sucesso.`);

    } catch (error) {
        console.error('Ocorreu um erro geral:', error);
    } finally {
        if (client) {
            await client.close();
            console.log('Conexão com o MongoDB fechada.');
        }
    }
}

// Executa a função
convertExcelToOrderedJson();