const fs = require('fs');
const xlsx = require('node-xlsx').default;
const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');

async function scrapeTable() {
    let options = new chrome.Options();
    options.addArguments('--ignore-certificate-errors'); // Esse código ignora os erros de certificados de bloqueio

    let driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(options)
        .build();
    try {
        await driver.get('http://webapplayers.com/inspinia_admin-v2.9.4/table_data_tables.html'); // Faz a navegação para a página

        let data = [];

        // Determinar o número total de páginas
        let pagination = await driver.findElement(By.className('dataTables_paginate'));
        let pages = await pagination.findElements(By.tagName('a'));
        let totalPages = pages.length - 2; // Desconsidera os botões "Anterior" e "Próximo"

        // Iterar sobre todas as páginas
        for (let i = 0; i < totalPages; i++) {
            let table = await driver.findElement(By.id('DataTables_Table_0')); //Localiza o elemento HTML da tabela na página
            let rows = await table.findElements(By.css('tbody tr')); // Localiza todas as linhas da página

            for (let row of rows) {
                let cells = await row.findElements(By.css('td'));
                let rowData = [];
                for (let cell of cells) {
                    let cellText = await cell.getText();
                    rowData.push(cellText);
                }
                data.push(rowData);
            }

            // Clicar no botão "Próximo" para ir para a próxima página
            if (i < totalPages - 1) {
                let nextButton = await pagination.findElement(By.linkText('Next'));
                await nextButton.click();
                await driver.sleep(2000); // Aguarda um tempo até que a pagina carregue por completo
            }
        }

        let buffer = xlsx.build([{ name: "TableData", data: data }]);
        fs.writeFileSync('table_data.xlsx', buffer);
        console.log('Excel file created successfully.');
    } finally {
        await driver.quit();
    }
}

async function fillForm() {
    // Faz a leitura dos dados em arquivo excel.
    let data;
    try {
        data = xlsx.parse(fs.readFileSync('table_data.xlsx'))[0].data;
    } catch (error) {
        console.error('Erro ao ler o arquivo Excel:', error.message);
        return;
    }

    let options = new chrome.Options();
    options.addArguments('--ignore-certificate-errors');// Esse código ignora os erros de certificados de bloqueio.

    let driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(options)
        .build();
    try {
        await driver.get('http://webapplayers.com/inspinia_admin-v2.9.4/form_editors.html');
        console.log('Página carregada.');
        
        // faz a localização da div fornecida da página.
        let editableDiv = await driver.wait(until.elementLocated(By.xpath('//*[@id="page-wrapper"]/div[3]/div[1]/div/div/div[2]/div[2]/div[3]/div[2]')), 20000);
        await driver.wait(until.elementIsVisible(editableDiv), 20000);
        console.log('Div editável localizada.');

        let text = data.map(row => `- ${row.join(', ')};<br>`).join(''); // Faz a inclusão do traço no começo, espaço entre as palavras, ponto e vírgula no fim quebra de linha.
        console.log('Texto preparado:', text);

        await driver.executeScript('arguments[0].innerHTML = arguments[1];', editableDiv, text);
        console.log('Formulário preenchido com sucesso com os dados do Excel.');

        await driver.actions().move({ origin: editableDiv }).click().perform();
        console.log('Interação do usuário simulada.');

        await driver.sleep(5000);

        let screenshotPath = path.join(__dirname, 'filled_form.png');
        let screenshot = await driver.takeScreenshot();
        fs.writeFileSync(screenshotPath, screenshot, 'base64');
        console.log(`Captura de tela salva em ${screenshotPath}`); // Faz uma captura de tela e salva em arquivo para verificar se a edição foi feita correta.
    } catch (error) {
        console.error('Erro:', error);
    } finally {
        await driver.quit();
    }
}

async function main() {
    await scrapeTable();
    await fillForm();
}

main();
