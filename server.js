const express = require('express');
const multer = require('multer');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

app.post('/upload', upload.array('sheets', 2), (req, res) => {
    if (req.files.length < 2) {
        return res.status(400).send('Por favor, faça o upload de duas planilhas.');
    }

    let workbook1, workbook2;

    try {
        // Tentativa de leitura dos arquivos Excel
        workbook1 = XLSX.readFile(req.files[0].path);
        workbook2 = XLSX.readFile(req.files[1].path);
    } catch (error) {
        // Se ocorrer um erro durante a leitura dos arquivos, ele será capturado aqui
        console.error('Erro ao ler os arquivos:', error);
        return res.status(500).send('Erro ao processar as planilhas. Por favor, verifique se os arquivos são válidos.');
    }

    const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
    const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

    // Definir a coluna específica para comparação, por exemplo, coluna "A"
    const targetColumn = 'A';

    // Converter a coluna específica em arrays de valores
    const columnValues1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 })
        .map(row => row[0]);  // A coluna "A" é o índice 0

    const columnValues2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 })
        .map(row => row[0]);

    // Comparar os valores da primeira planilha com os da segunda
    const missingValues = columnValues1.filter(value => !columnValues2.includes(value));

    if (missingValues.length > 0) {
        let resultHtml = '<h1>Orders da planilha 1 não encontrados na planilha 2:</h1><ul>';

        missingValues.forEach(value => {
            resultHtml += `<li>${value}</li>`;
        });

        resultHtml += '</ul>';
        res.send(resultHtml);
    } else {
        res.send('<h1>Todos os valores da Coluna A da primeira planilha estão presentes na segunda planilha.</h1>');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
});
