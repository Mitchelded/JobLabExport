const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path')
const XlsxPopulate = require('xlsx-populate');
const ExcelJS = require('exceljs');


(async () => {
    const cookieFilePath = path.join(__dirname, 'config.json');
    let phpSessionId = '';
    if (fs.existsSync(cookieFilePath)) {

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
        await emailInput.type('EMAIL');
        const passInput = await page.$('input[type="password"]');
        await passInput.type('PASSWORD');
        await page.click('input[type="radio"][value="employer"]');
        await page.waitForNavigation({
            waitUntil: 'domcontentloaded'
        });

        const cookies = await page.cookies();
        const phpSessCookie = cookies.find((cookie) => cookie.name === 'PHPSESSID');
        if (phpSessCookie) {
            phpSessionId = phpSessCookie.value;
            const cookieData = JSON.stringify({
                PHPSESSID: phpSessionId
            });
            fs.writeFileSync(cookieFilePath, cookieData);
        } else {
            console.error('Не удалось получить PHPSESSID со страницы входа..');
            await browser.close();
            return;
        }
    }

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

    await page.goto('https://joblab.ru/search.php?r=res&submit=1');
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
        var nado = false;
        const headers = ['Ссылка',
            'Имя',
            'Телефон',
            'E-mail',
            'Получено на вакансию',
            'Отправлена вакансия',
            'Вакансия',
            'Проживание',
            'Заработная плата',
            'График работы',
            'Образование',
            'Опыт работы',
            'Гражданство',
            'Пол',
            'Возраст',
            'Компания',
            'Образования',
            'Иностранные языки',
            'Водительские права',
            'Командировки',
            'Курсы и тренинги',
            'Навыки и умения',
            'Обо мне',
            'Период работы',
            'Должность',
            'Окончание',
            'Учебное заведение',
            'Специальность'
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
            sheet.cell(rowIndex, 1).value(arrLinks[i]).hyperlink(arrLinks[i]);
            sheet.cell(rowIndex, 1).style({
                underline: true,
                fontColor: "0000FF"
            });

			

            const arrJobData = await page.evaluate(() => {
                const selectorJob = 'body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > h1';

                const tbodyElement = document.querySelector(selectorJob).innerText;
                return tbodyElement;
            });




            const arrRole = await page.evaluate(() => {
                const tbodyElement = document.querySelector('body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody');
                const rows = tbodyElement.querySelectorAll('tr');
                return Array.from(rows, (row) => Array.from(row.querySelectorAll('td'), (column) => column.innerText.trim()));
            });

            let workHistory = '';
            var eduHistory = '';
            for (let j = 0; j < arrRole.length; j++) {
                const item = arrRole[j];
                if (item.length > 1) {
                    var columnHeader = item[0];
                    let columnValue = item[1];
                    const columnIndex = headers.findIndex((header) => columnHeader.includes(header));
                    if (columnIndex !== -1) {
                        const cell = sheet.cell(rowIndex, columnIndex + 1);
                        if (columnHeader === 'Обязанности') {
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

                        if (columnHeader === 'Образование') {
                            const preferHeader = arrRole[j + 1][0];
                            if (preferHeader === 'Опыт работы') {
                                var nado = true;
                            }



                        }

                        if (nado) {
                            if (columnHeader === 'Образование') {
                                eduHistory += columnValue + ', '; // Append to work history
                            }
                            if (columnHeader === 'Окончание') {
                                eduHistory += columnValue + ', '; // Append to work history
                                columnValue = "";
                            }
                            if (columnHeader === 'Учебное заведение') {
                                eduHistory += columnValue + ', '; // Append to work history and add line break
                                columnValue = "";
                            }
                            if (columnHeader === 'Специальность') {
                                eduHistory += columnValue + '\n\n'; // Append to work history and add line break
                                columnValue = "";

                            }

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

                        workbook.sheet("Sheet1").cell(`Q${i+1}`).value(removeDuplicateWords(eduHistory.trim()));
                        sheet.cell(rowIndex, columnIndex + 1).style({
                            wrapText: 'true',
                            horizontalAlignment: 'left',
                            verticalAlignment: 'top'
                        });

                        workbook.sheet("Sheet1").cell(`G${i+1}`).value(removeDuplicateWords(arrJobData));

                        sheet.cell(rowIndex, columnIndex + 1).style({
                            wrapText: 'true',
                            horizontalAlignment: 'left',
                            verticalAlignment: 'top'
                        });
                        const column = sheet.column(columnIndex + 1);
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

                        const headerRange = sheet.range(`A1:W${arrLinks.length}`);
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
            nado = false;
            rowIndex++;

            console.log(`Данные Кандидата_${i} добавлены в файл Excel\n`);
        }
        await workbook.find("Должность", "");
        await workbook.find("Период работы", "");
        await workbook.find("Компания", "Опыт работ");
        await workbook.find("Учебное заведение", "");
        await workbook.find("Специальность", "");
        await workbook.find("Окончание", "");

        var pathExcel = 'D:/Storage/JobLab';

		// const name = readline.question("Enter vacancy ");
		// for(let i = 0; i<arrLinks.length;i++){

			// var foundOrNot = workbook.sheet("Sheet1").cell(`G${i+1}`).find(name);
			// if(foundOrNot){
				// console.log(`Найдено ${name}`)
			// }
			// else{
				// console.log(`Не найдено ${name}`)
			// }
		// }

        await workbook.toFileAsync(`${pathExcel}/candidates.xlsx`);
        XlsxPopulate.fromFileAsync(`${pathExcel}/candidates.xlsx`)
            .then(workbook => {
                const sheet = workbook.sheet(0);
                const columnAddress = 'A';
                const columnAddress1 = 'Q';
                const columnAddress2 = 'G';
                const column = sheet.column(columnAddress);
                const column2 = sheet.column(columnAddress2);
                const column1 = sheet.column(columnAddress1);
                column2.width(87);
                column1.width(142);
                column.width(32);
                column1.style({
                    wrapText: true
                });
                column.style({
                    wrapText: true
                });
                return workbook.toFileAsync(`${pathExcel}/candidates.xlsx`);
            })
            .then(() => {})
            .catch(err => {
                console.error('Произошла ошибка:', err);
            });

        await workbook.toFileAsync(`${pathExcel}/candidates.xlsx`)

        console.log(`Файл Excel со всеми данными создан и сохранен по пути: ${pathExcel}/candidates.xlsx`);

        await browser.close();
    }
})();


function removeDuplicateWords(sentence) {
    var words = sentence.split(", ");
    var result = [];

    for (var i = 0; i < words.length; i++) {
        if (i === 0 || words[i] !== words[i - 1]) {
            result.push(words[i]);
        }
    }

    return result.join(", ");
}