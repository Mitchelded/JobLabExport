const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path')
const XlsxPopulate = require('xlsx-populate');

(async () => {

    const cookieFilePath = path.join(__dirname, 'cookies.json'); // Path to the cookie file

    let phpSessionId = ''; // Variable to store the PHPSESSID value

    if (fs.existsSync(cookieFilePath)) {
        // Read the PHPSESSID value from the file
        const cookieData = fs.readFileSync(cookieFilePath, 'utf8');
        const {
            PHPSESSID
        } = JSON.parse(cookieData);
        phpSessionId = PHPSESSID;
    }

    const browser = await puppeteer.launch({
        executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe',
        headless: false
    });
    const page = await browser.newPage();

    if (!phpSessionId) {
        await page.goto('https://joblab.ru/access.php');


        const emailInput = await page.$('input[type="email"]');
        await emailInput.type('ast@5092778.ru');

        const passInput = await page.$('input[type="password"]');
        await passInput.type('8014802530');

        // Click on the radio button with value="employer"
        await page.click('input[type="radio"][value="employer"]');

        await page.waitForNavigation({
            waitUntil: 'domcontentloaded'
        });
        // Retrieve the PHPSESSID from the cookies
        const cookies = await page.cookies();
        const phpSessCookie = cookies.find((cookie) => cookie.name === 'PHPSESSID');

        if (phpSessCookie) {
            phpSessionId = phpSessCookie.value;

            // Save the PHPSESSID value to the file
            const cookieData = JSON.stringify({
                PHPSESSID: phpSessionId
            });
            fs.writeFileSync(cookieFilePath, cookieData);
        } else {
            console.error('Failed to retrieve PHPSESSID from the login page.');
            await browser.close();
            return;
        }
    }

    // Set the PHPSESSID cookie
    await page.setCookie({
        name: 'PHPSESSID',
        value: phpSessionId,
        domain: 'joblab.ru',
        path: '/',
        expires: Math.floor(Date.now() / 1000) + 3600,
        httpOnly: false,
        secure: false,
        sameSite: ''
    });

    await page.goto('https://joblab.ru/employers/inbox.php');
    await page.reload();



    const arrLinks = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a[href]')).map((el) => {
            return {
                text: el.innerText.trim(),
                href: el.href.trim()
            };
        });

        const filteredLinks = links.filter((link) => link.href.includes("res"));

        return filteredLinks.map((link) => link.href);
    });


    if (arrLinks.length - 1 == 0) {
        browser.close();
        console.log("Куки устарели")
        fs.unlink('cookies.json', (error) => {
            if (error) {
                console.error('Ошибка при удалении файла:', error);
            } else {
                console.log('Файл куки успешно удален. Перезапустите программу');
            }
        });

    } else {
        console.log(`Найдено ${arrLinks.length - 1} кандидата`);

        const workbook = await XlsxPopulate.fromBlankAsync();
        const sheet = workbook.sheet(0);

        let rowIndex = 1;

        const headers = [
            'Ссылка',
            'Имя',
            'Телефон',
            'E-mail',
            'Получено на вакансию',
            'Отправлена вакансия',
            'Проживание',
            'Заработная плата',
            'График работы',
            'Образование',
            'Опыт работы',
            'Гражданство',
            'Пол',
            'Возраст',
            'Период работы',
            'Должность',
            'Компания',
            'Обязанности',
            'Образование',
            'Окончание образования',
            'Учебное заведение',
            'Специальность',
            'Иностранные языки',
            'Водительские права',
            'Командировки',
            'Курсы и тренинги',
            'Навыки и умения',
            'Обо мне'
        ];

        for (let i = 0; i < headers.length; i++) {
            sheet.cell(rowIndex, i + 1).value(headers[i]);
        }

        rowIndex++;

        for (let i = 1; i < arrLinks.length; i++) {
            const link = arrLinks[i];
            console.log(`Переход по ссылке ${link}`);
            await page.goto(link);
            console.log(`Получаем данные Кандидата_${i}`);

            const candidateData = {};
            candidateData['Ссылка'] = arrLinks[i]; // Add it to the candidateData object
            sheet.cell(rowIndex, 1)
                .value(arrLinks[i])
                .hyperlink(arrLinks[i]); // Set the value and add hyperlink
            sheet.cell(rowIndex, 1).style({
                underline: true,
                fontColor: "0000FF"
            }); // Style hyperlink
            const arrRole = await page.evaluate(() => {
                const tbodyElement = document.querySelector(
                    'body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody'
                );
                const rows = tbodyElement.querySelectorAll('tr');
                return Array.from(rows, (row) => Array.from(row.querySelectorAll('td'), (column) => column.innerText.trim()));
            });

            for (let j = 0; j < arrRole.length; j++) {
                const item = arrRole[j];

                if (item.length > 1) {
                    const columnHeader = item[0];
                    let columnValue = item[1];

                    const columnIndex = headers.findIndex((header) => columnHeader.includes(header));

                    if (columnIndex !== -1) {
                        const cell = sheet.cell(rowIndex, columnIndex + 1);
                        if (columnHeader === 'Обязанности') {
                            // Remove line breaks from the column value
                            columnValue = columnValue.replace('\n', '');
                        }
                        if (columnHeader === 'Отправлена вакансия') {
                            const arrJob2 = await page.evaluate(() => {
                                let arrJob = [];
                                for (let i = 1; i <= 15; i++) {
                                    const selector = `body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(${i}) > td:nth-child(2)`;
                                    const element = document.querySelector(selector);
                                    if (element && (element.innerText.includes('просмотрена') || element.innerText.includes('приглашен') && !element.innerText.includes('Пригласить'))) {
                                        arrJob.push(element.innerText);
                                    }
                                }
                                return arrJob;
                            });

                            columnValue = arrJob2[0].replace(/Пригласить|Написать|Подумать|Отклонить/g, '').trim();
                        }
                        if (columnHeader === 'Получено на вакансию') {
                            const arrJob1 = await page.evaluate(() => {
                                let arrJob = [];
                                for (let i = 1; i <= 15; i++) {
                                    const selector = `body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(${i}) > td:nth-child(2)`;
                                    const element = document.querySelector(selector);
                                    if (element && (element.innerText.includes('Пригласить') && element.innerText.includes('Написать') && element.innerText.includes('Подумать') && element.innerText.includes('Отклонить'))) {
                                        arrJob.push(element.innerText);
                                    }
                                }

                                return arrJob;
                            });

                            columnValue = arrJob1[0].replace(/Пригласить|Написать|Подумать|Отклонить/g, '').trim();
                        }

                        sheet.cell(rowIndex, columnIndex + 1).value(columnValue);
                        sheet.cell(rowIndex, columnIndex + 1).style({
                            wrapText: 'true',
                            horizontalAlignment: 'left',
                            verticalAlignment: 'top'
                        });
                        const column = sheet.column(columnIndex + 1);
                        const contentLength = columnValue ? columnValue.length : 0; // Ensure columnValue is defined
                        const maxContentLength = getMaxContentLength(columnIndex);
                        const columnWidth = Math.max(contentLength, maxContentLength) + 2;
                        column.width(columnWidth);
                        candidateData[columnHeader] = columnValue;
                        const headerRange = sheet.range(`A${rowIndex-1}:AA${rowIndex-1}`);
                        headerRange.style({
                            border: true,
                            horizontalAlignment: 'left',
                            verticalAlignment: 'top'
                        });
                        headerRange.forEach(cell => {
                            cell.style({
                                fill: {
                                    type: 'pattern',
                                    patternType: 'solid',
                                    fgColor: 'FFFF00'
                                }
                            });
                        });
                        cell.value(columnValue)
                            .style({
                                wrapText: 'true',
                                horizontalAlignment: 'left',
                                verticalAlignment: 'top',
                                border: true
                            });
                    }
                }
            }



            rowIndex++;

            console.log(`Данные Кандидата_${i} добавлены в файл Excel\n`);
        }
		var pathExcel = 'D:/Storage/JobLab';
        await workbook.toFileAsync(`${pathExcel}/canditates.xlsx`);
        console.log(`Файл Excel со всеми данными создан и сохранен в ${pathExcel}.`);

        await browser.close();


    }



})();

function getMaxContentLength(arrLinks, columnIndex) {
    let maxLength = 0;
    return maxLength;
}