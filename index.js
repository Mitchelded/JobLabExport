const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const XlsxPopulate = require('xlsx-populate');
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
        await emailInput.type('ast@5092778.ru');
        const passInput = await page.$('input[type="password"]');
        await passInput.type('8014802530');
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
    await page.goto("https://joblab.ru/search.php?r=res&srprofecy=&kw_w2=1&srzpmax=&srregion=50&srcity=77&srcategory=&submit=1&srexpir=&srgender=");

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    rl.question('Введите ключевые слова: ', async (keywords) => {
        rl.close();

        const keywordsInput = await page.$('input[name="srprofecy"]');
        await keywordsInput.type(keywords);

        const submitButton = await page.$('input[type="submit"][value="Найти"]');
        await Promise.all([
            page.waitForNavigation({
                waitUntil: 'domcontentloaded'
            }),
            submitButton.click()
        ]);
        const currentURL = await page.url();
        await page.waitForTimeout(200);
        await collectLinks(currentURL);
    });
    // Создание Excel-файла
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0);

    // Заголовки столбцов
    const headers = ['Профессия', 'Имя', 'Место жительства', 'Зарплата', 'Опыт работы'];
const headerRow = sheet.row(1);
headers.forEach((header, index) => {
  const cell = headerRow.cell(index + 1);
  cell.value(header);
  cell.style("bold", true);
});

	const people = []; // Массив для хранения всех людей
    const collectLinks = async (nextPageUrl = '') => {
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

        const targetUrl = nextPageUrl;
        await page.goto(targetUrl);
        await page.reload();

        const handleCaptcha = async () => {
            console.log('Captcha detected. Please enter the captcha code:');
            const rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });
            const captchaCode = await new Promise((resolve) => {
                rl.question('', (answer) => {
                    resolve(answer);
                    rl.close();
                });
            });

            await page.evaluate((code) => {
                const captchaInput = document.querySelector('input[name="keystring"]');
                captchaInput.value = code;
            }, captchaCode);

            await Promise.all([
                page.waitForNavigation({
                    waitUntil: 'domcontentloaded'
                }),
                page.click('input[name="submit_captcha"]')
            ]);
        };

        while (true) {
            const arrLinks = await page.evaluate(() => {
                const links = Array.from(document.querySelectorAll('a[href]')).map((el) => {
                    return {
                        text: el.innerText.trim(),
                        href: el.href.trim()
                    };
                });
                const filteredLinks = links.filter((link) => link.href.includes(".html"));
                return filteredLinks.map((link) => link.href);
            });

            const loginButton = await page.$('a[href="/access.php"][rel="nofollow"]');
            if (loginButton) {
                console.log("Куки устарели");
                fs.unlink('config.json', (error) => {
                    if (error) {
                        console.error('Ошибка при удалении файла:', error);
                    } else {
                        console.log('Файл куки успешно удален. Перезапустите программу');
                    }
                });
                browser.close();
                return;
            }

            const arrLinks2 = await page.evaluate(() => {
                const links = Array.from(document.querySelectorAll('body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > p:nth-child(9) > a:nth-child(2)')).map((el) => {
                    return {
                        text: el.innerText.trim(),
                        href: el.href.trim()
                    };
                });
                const filteredLinks = links.filter((link) => link.href.includes("&page=") && link.href.includes("&submit=") && link.href.includes("search.php?r=res&"));

                return filteredLinks.map((link) => link.href);
            });

            for (let i = 0; i < arrLinks.length; i++) {
                const link = arrLinks[i];
                await page.goto(link);

                const captchaInput = await page.$('input[name="keystring"]');
                if (captchaInput) {
                    await handleCaptcha();
                }
                const arrName = await page.evaluate(() => {
                    var name = document.querySelector('body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(1) > td:nth-child(2)').innerText;

                    for (let i = 1; i < 21; i++) {
                        var job = document.querySelector("body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > h1").innerText;
                        var residenceSel = document.querySelector(`body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(${i}) > td:nth-child(2)`);
                        if (residenceSel !== null) {
                            var residence = residenceSel.innerText;
                            if (residence.includes("Москва")) {
                                var salarySel = document.querySelector(`body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(${i+1}) > td:nth-child(2)`)
                                var salary = salarySel.innerText;
                                var workExpSel = document.querySelector(`body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(${i+5}) > td:nth-child(2)`)
                                var workExp = workExpSel.innerText;
                                return {
                                    job,
                                    name,
                                    residence,
                                    salary,
                                    workExp
                                };
                            }
                        }
						// Запись содержимого из arrName в таблицу Excel

                    }

                });


people.push(arrName); // Добавляем человека в массив
                console.log(`Профессия: ${arrName.job}\n Имя: ${arrName.name}\n Место жительства: ${arrName.residence}\n Зарплата: ${arrName.salary}\n Опыт работы: ${arrName.workExp}\n\n`);


                // Wait for some time before navigating to the next link
                await page.waitForTimeout(200);
            }
			// Запись содержимого из people в таблицу Excel
people.forEach((person, index) => {
  const dataRow = sheet.row(index + 2);
  dataRow.cell(1).value(person.job);
  dataRow.cell(2).value(person.name);
  dataRow.cell(3).value(person.residence);
  dataRow.cell(4).value(person.salary);
  dataRow.cell(5).value(person.workExp);
});
// Применение границ к остальным данным в таблице
const dataRange = sheet.range(`A1:E${people.length + 1}`);
dataRange.style("border", true);
sheet.column("A").width(80); // Установка размера столбцов
sheet.column("B").width(32); // Установка размера столбцов
sheet.column("C").width(35); // Установка размера столбцов
sheet.column("D").width(19); // Установка размера столбцов
sheet.column("E").width(17); // Установка размера столбцов

            if (arrLinks2.length > 0) {
                const nextPageUrl = arrLinks2[0];

                // Update phpSessionId and collect links on the next page
                await collectLinks(nextPageUrl);
            } else {
                break;
            }
        }

        browser.close();
        // Сохранение Excel-файла
        const excelFilePath = path.join(__dirname, 'output.xlsx');
        await workbook.toFileAsync(excelFilePath);
        console.log('Excel-файл успешно создан:', excelFilePath);
    };


})();