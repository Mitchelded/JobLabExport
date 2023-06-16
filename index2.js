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

    if (arrLinks.length - 1 === 0) {
        browser.close();
        console.log("Куки устарели");
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
        const headers = ['Ссылка',
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
            'Компания',
			'Образование',
            'Окончание образования',
            'Учебное заведение',
            'Специальность',
            'Иностранные языки',
            'Водительские права',
            'Командировки',
            'Курсы и тренинги',
            'Навыки и умения',
            'Обо мне',
            'Период работы',
            'Должность',
        ]; // Updated headers

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
            sheet.cell(rowIndex, 1).value(arrLinks[i]).hyperlink(arrLinks[i]); // Set the value and add hyperlink
            sheet.cell(rowIndex, 1).style({
                underline: true,
                fontColor: "0000FF"
            }); // Style hyperlink

            const arrRole = await page.evaluate(() => {
                const tbodyElement = document.querySelector('body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody');
                const rows = tbodyElement.querySelectorAll('tr');
                return Array.from(rows, (row) => Array.from(row.querySelectorAll('td'), (column) => column.innerText.trim()));
            });

            let workHistory = ''; // Variable to store work history

            for (let j = 0; j < arrRole.length; j++) {
                const item = arrRole[j];
                if (item.length > 1) {
                    var columnHeader = item[0];
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
                        if (columnHeader === 'Период работы') {
                            workHistory += columnValue + ', '; // Append to work history
                            columnValue = "";
                        }
                        if (columnHeader === 'Должность') {
                            workHistory += columnValue + ', '; // Append to work history
                            columnValue = "";
                        }
                        if (columnHeader === 'Компания') {
                            workHistory += columnValue + '\n\n'; // Append to work history and add line break
                            columnValue = workHistory.trim(); // Update column value
                        }


                        sheet.cell(rowIndex, columnIndex + 1).value(columnValue);
                        sheet.cell(rowIndex, columnIndex + 1).style({
                            wrapText: 'true',
                            horizontalAlignment: 'left',
                            verticalAlignment: 'top'
                        });
                        const column = sheet.column(columnIndex + 1);
                        candidateData[columnHeader] = columnValue;
                        switch (columnHeader) {
                            case 'Имя':
                                column.width(22);
                                break;
                            case 'Телефон':
                                column.width(14);
                                break;
                            case 'E-mail':
                                column.width(27);
                                break;
                            case 'Получено на вакансию':
                                column.width(95.29);
                                break;
                            case 'Отправлена вакансия':
                                column.width(97.29);
                                break;
                            case 'Проживание':
                                column.width(36.29);
                                break;
                            case 'Заработная плата':
                                column.width(18.29);
                                break;
                            case 'График работы':
                                column.width(85.86);
                                break;
                            case 'Образование':
                                column.width(21.14);
                                break;
                            case 'Опыт работы':
                                column.width(16.71);
                                break;
                            case 'Гражданство':
                                column.width(11.86);
                                break;
                            case 'Пол':
                                column.width(9);
                                break;
                            case 'Возраст':
                                column.width(25.14);
                                break;
                            case 'Компания':
                                column.width(99.14);
                                break;
                            case 'Окончание образования':
                                column.width(100); // Надо будет изменить
                                break;
                            case 'Учебное заведение':
                                column.width(85);
                                break;
                            case 'Специальность':
                                column.width(84.14);
                                break;
                            case 'Иностранные языки':
                                column.width(42.14);
                                break;
                            case 'Водительские права':
                                column.width(19.14);
                                break;
                            case 'Командировки':
                                column.width(26.14);
                                break;
                            case 'Курсы и тренинги':
                                column.width(97.71);
                                break;
                            case 'Навыки и умения':
                                column.width(99.43);
                                break;
                            case 'Обо мне':
                                column.width(99.43);
                                break;
                        }

                        const headerRange = sheet.range(`A1:X${arrLinks.length}`);
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

                    }
                }
            }

            rowIndex++;


            console.log(`Данные Кандидата_${i} добавлены в файл Excel\n`);
        }
        var pathExcel = 'D:/Storage/JobLab';
        await workbook.find("Должность", "");
        await workbook.find("Период работы", "");
        await workbook.find("Компания", "Опыт работ");
        await workbook.toFileAsync(`${pathExcel}/canditates.xlsx`);
XlsxPopulate.fromFileAsync(`${pathExcel}/canditates.xlsx`)
  .then(workbook => {
    const sheet = workbook.sheet(0);
    const columnAddress = 'A';

    const column = sheet.column(columnAddress);
    column.width(32); // Установите ширину столбца в символах

    return workbook.toFileAsync(`${pathExcel}/canditates.xlsx`);
  })
  .then(() => {
  })
  .catch(err => {
    console.error('Произошла ошибка:', err);
  });
        console.log(`Файл Excel со всеми данными создан и сохранен по пути: ${pathExcel}/canditates.xlsx`);

        await browser.close();
    }
})();