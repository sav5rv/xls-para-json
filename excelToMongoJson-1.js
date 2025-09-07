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
        console.log('Conectado ao MongoDB com sucesso!');

        const db = client.db(dbName);
        const collection = db.collection(collectionName);

        // 2. Obter a ordem dos campos de um documento existente
        // Busca o primeiro documento para usar como modelo
        const sampleDoc = await collection.findOne({});
        if (!sampleDoc) {
            console.error('Nenhum documento encontrado na coleção. A ordem dos campos não pode ser definida.');
            return;
        }
        const fieldOrder = Object.keys(sampleDoc).filter(key => key !== '_id');
        console.log('Ordem dos campos obtida do MongoDB:', fieldOrder.join(', '));

// 3. Ler o arquivo Excel com a opção de manter valores originais
const workbook = xlsx.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
// O 'raw: false' faz com que o xlsx retorne os valores formatados como strings
// O 'defval: ""' define um valor padrão para células vazias
const excelData = xlsx.utils.sheet_to_json(worksheet, { raw: false, defval: "" });
console.log('3. Dados lidos do Excel.');

if (excelData.length === 0) {
    console.error('ERRO: Nenhum dado encontrado no arquivo Excel. Verifique o nome da planilha ou se ela está vazia.');
    return;
}

// 4. Mapear e reordenar os dados do Excel
const orderedData = excelData.map(row => {
    const newRow = {};

    // 4.1. Itera sobre a ordem dos campos do MongoDB
    fieldOrder.forEach(field => {
        const normalizedField = field.toLowerCase().trim();
        
        // Verifica se o campo existe na linha do Excel (ignorando a caixa)
        const excelKey = Object.keys(row).find(key => key.toLowerCase().trim() === normalizedField);
        
        if (excelKey) {
            let value = row[excelKey];
            
            // 4.2. Formata os campos de data e hora para strings
            // Verifica o nome do campo para aplicar a formatação
            if (field.toLowerCase() === 'data') {
                // Apenas converte se for um valor numérico válido
                if (typeof value === 'number' && !isNaN(value)) {
                    const excelDate = new Date(Math.round((value - 25569) * 86400 * 1000));
                    value = excelDate.toISOString().split('T')[0]; // Formato YYYY-MM-DD
                } else {
                    value = null; // Ou "" ou a mensagem que desejar
                }
            }

            if (field.toLowerCase() === 'hrinicio') {
                // Apenas converte se for um valor numérico válido
                if (typeof value === 'number' && !isNaN(value)) {
                    const excelTime = new Date(Math.round((value - 25569) * 86400 * 1000));
                    const hours = String(excelTime.getUTCHours()).padStart(2, '0');
                    const minutes = String(excelTime.getUTCMinutes()).padStart(2, '0');
                    value = `${hours}:${minutes}`; // Formato HH:MM
                } else {
                    value = null; // Ou "" ou a mensagem que desejar
                }
            }

            newRow[field] = value;
        }
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
        console.error('Ocorreu um erro:', error);
    } finally {
        if (client) {
            await client.close();
            console.log('Conexão com o MongoDB fechada.');
        }
    }
}

// Executa a função
convertExcelToOrderedJson();