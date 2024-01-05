const ExcelJS = require('exceljs');
const csv = require('csv-parser');
const axios = require('axios');
const { Readable } = require('stream');

async function streamToBuffer(stream) {
    const chunks = [];
    for await (const chunk of stream) {
        chunks.push(chunk instanceof Buffer ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
}

async function convertExcelToJson(fileUrl) { // La función ahora acepta fileUrl como parámetro
    const response = await axios({
        method: 'get',
        url: fileUrl,
        responseType: 'stream'
    });

    const format = fileUrl.split('.').pop().toLowerCase();
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
                    const cellValue = cell.value;
                    if (cell.type === ExcelJS.ValueType.Number) {
                        rowData[headers[colNumber]] = cellValue;
                    } else if (typeof cellValue === 'object' && cellValue.result !== undefined) {
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

module.exports = convertExcelToJson; // Exporta la función para usarla en otros archivos
