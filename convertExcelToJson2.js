const ExcelJS = require('exceljs');
const csv = require('csv-parser');
const axios = require('axios');
const { Readable } = require('stream');

let excelEjemplo = 'https://07109f54-7493-4017-b768-8102e95cfb89.usrfiles.com/ugd/07109f_40fb214d7f0c45b897903965263e92ed.xlsx'
let csvEjemplo = 'https://07109f54-7493-4017-b768-8102e95cfb89.usrfiles.com/ugd/07109f_2f59dd6227194e92a53984e29b1a3be6.csv'


async function streamToBuffer(stream) {
    const chunks = [];
    for await (const chunk of stream) {
        chunks.push(chunk instanceof Buffer ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
}

async function convertExcelToJson(url) {
    const response = await axios({
        method: 'get',
        url: url,
        responseType: 'stream'
    });

    const format = url.split('.').pop().toLowerCase();
    let data = [];

    if (format === 'xlsx') {
        const buffer = await streamToBuffer(response.data);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const worksheet = workbook.worksheets[0];

        let headers = [];
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const rowData = {};
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                if (rowNumber === 1) {
                    headers[colNumber] = cell.value;
                } else {
                    // Verificar si el valor de la celda es numérico y manejarlo adecuadamente
                    const cellValue = cell.value;
                    if (cell.type === ExcelJS.ValueType.Number) {
                        rowData[headers[colNumber]] = cellValue;
                    } else if (typeof cellValue === 'object' && cellValue.result !== undefined) {
                        // Para fórmulas, puedes decidir cómo manejar el resultado
                        rowData[headers[colNumber]] = cellValue.result;
                    } else {
                        rowData[headers[colNumber]] = cellValue;
                    }
                }
            });
            if (rowNumber > 1) {
                data.push(rowData);
            }
        });
    } else if (format === 'csv') {
        const results = [];
        const stream = Readable.from(response.data);
        await new Promise((resolve, reject) => {
            stream
                .pipe(csv())
                .on('data', (rowData) => results.push(rowData))
                .on('end', () => {
                    data = results;
                    resolve();
                })
                .on('error', reject);
        });
    } else {
        throw new Error('Unsupported file format');
    }

    return data;
}

convertExcelToJson(excelEjemplo).then(data => {
    console.log(JSON.stringify(data, null, 2));
}).catch(error => {
    console.error('Error:', error);
});