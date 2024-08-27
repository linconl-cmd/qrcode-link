const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const axios = require('axios');
const path = require('path');

const jsonFilePath = 'videos.json'; // Substitua pelo caminho para o seu arquivo JSON
const qrCodePageUrl = 'https://linconl-cmd.github.io/proj-qrcode/'; // Substitua pela URL da página de geração de QR code

(async () => {
    // Lê os dados do arquivo JSON
    const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));

    // Inicializa o navegador e a página
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(qrCodePageUrl);

    // Cria uma nova planilha
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('QRCode Data');

    // Adiciona os cabeçalhos à planilha
    worksheet.getCell('E1').value = 'Nome do Vídeo';
    worksheet.getCell('J1').value = 'Link do Vídeo';
    worksheet.getCell('K1').value = 'QR Code';

    // Define a altura da linha para 3,40 cm (altura em pontos, 96 dpi -> 1 ponto = 1/72 polegada)
    const rowHeight = 3.40 * 28.35; // Convertendo de cm para pontos (28.35 pontos/cm)
    
    // Define o tamanho do QR code para 3,50 cm x 3,50 cm
    const qrCodeSize = 3 * 28.35; // Convertendo 3,50 cm para pontos (28.35 pontos/cm)

    for (const [index, video] of jsonData.videos.entries()) {
        const { name, link } = video;

        // Acessa a página de geração de QR code
        await page.goto(qrCodePageUrl);

        // Preenche o campo com o link do vídeo e clica no botão para gerar o QR code
        await page.type('input#ipt', link); // Substitua '#linkField' pelo seletor correto do campo de input
        await page.click('button#btn'); // Substitua '#generateQRCode' pelo seletor correto do botão

        // Aguarda o QR code ser gerado
        await page.waitForSelector('img#qrcode-img'); // Substitua '#qrcodeImage' pelo seletor correto da imagem do QR code

        // Captura o QR code
        const qrCodeUrl = await page.$eval('img#qrcode-img', img => img.src); // Substitua '#qrcodeImage' pelo seletor correto da imagem

        // Baixa a imagem do QR code
        const imagePath = path.resolve(__dirname, `qrcode-${index}.png`);
        const response = await axios({
            url: qrCodeUrl,
            responseType: 'arraybuffer',
        });
        fs.writeFileSync(imagePath, Buffer.from(response.data, 'binary'));

        // Adiciona os dados à planilha
        const rowIndex = index + 2; // O índice da linha começa em 2 para ajustar a altura
        worksheet.getCell(`E${rowIndex}`).value = name;
        worksheet.getCell(`J${rowIndex}`).value = link;

        // Define a altura da linha
        worksheet.getRow(rowIndex).height = rowHeight;

        // Insere a imagem do QR code na planilha com tamanho 3,50 cm x 3,50 cm
        const imageId = workbook.addImage({
            filename: imagePath,
            extension: 'png',
        });

        // Ajusta as dimensões da imagem (conversão para pontos)
        worksheet.addImage(imageId, {
            tl: { col: 10, row: index + 1 }, // Posição inicial da imagem na coluna K (índice 10)
            ext: { width: qrCodeSize, height: qrCodeSize } // Tamanho da imagem em pontos
        });
    }

    // Salva a planilha em um arquivo
    await workbook.xlsx.writeFile('QRCodeData.xlsx');

    // Fecha o navegador
    await browser.close();

    console.log('Dados exportados para QRCodeData.xlsx');
})();
