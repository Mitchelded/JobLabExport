const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const XlsxPopulate = require('xlsx-populate');
(async () => {
    const cookieFilePath = path.join(__dirname, 'config.json');
    let phpSessionId = '';
	let maxPeople = 100; // Объявление переменной maxPeople
if (fs.existsSync(cookieFilePath)) {
  const cookieData = fs.readFileSync(cookieFilePath, 'utf8');
  const { PHPSESSID, maxPeople: configMaxPeople } = JSON.parse(cookieData);
  phpSessionId = PHPSESSID;
  if (configMaxPeople) {
    maxPeople = configMaxPeople;
  }
} else {
  const configData = JSON.stringify({
    PHPSESSID: phpSessionId,
    maxPeople: maxPeople
  });
  fs.writeFileSync(cookieFilePath, configData);
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
                PHPSESSID: phpSessionId,
				maxPeople: 100
            });
            fs.writeFileSync(cookieFilePath, cookieData);
        } else {
            console.error('Не удалось получить PHPSESSID со страницы входа..');
            await browser.close();
            return;
        }
    }
    await page.goto("https://joblab.ru/search.php?r=res&srprofecy=&kw_w2=1&srzpmax=&srregion=50&srcity=77&srcategory=&submit=1&srexpir=3&srgender=");
		            const loginButton = await page.$('a[href="/access.php"]');
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
        await page.waitForTimeout(500);
        await collectLinks(currentURL);
    });
    // Создание Excel-файла
    const workbook = await XlsxPopulate.fromBlankAsync();
    const sheet = workbook.sheet(0);

    // Заголовки столбцов
    const headers = ['Профессия', 'Имя', 'Телефон', 'Почта', 'Место жительства', 'Зарплата', 'Опыт работы', 'Ссылка'];

    const headerRow = sheet.row(1);
    headers.forEach((header, index) => {
        const cell = headerRow.cell(index + 1);
        cell.value(header);
        cell.style("bold", true);
    });

    const people = []; // Массив для хранения всех людей
	



    let peopleCount = 0; // Counter for the number of people processed
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
            console.log('Капча обнаружена. Пожалуйста, введите код проверки:');
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
				    if (peopleCount >= maxPeople) {
        break; // Прервать цикл, если достигнуто максимальное количество людей
    }
                const link = arrLinks[i];
                await page.goto(link);

                const captchaInput = await page.$('input[name="keystring"]');
                if (captchaInput) {
                    await handleCaptcha();
                }
                // Получение ссылки элемента и выполнение клика
                const linkButPhone = await page.$('#p > a');
				if(linkButPhone){
					await linkButPhone.click();
				}
                await page.waitForTimeout(300);
                
				const linkButEmail = await page.$('#m > a');
				if(linkButEmail){
					 await linkButEmail.click();
				}
				await page.waitForTimeout(300);
                const arrName = await page.evaluate(() => {
                    var name = document.querySelector('body > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr > td > table.table-to-div > tbody > tr:nth-child(1) > td:nth-child(2)').innerText;
					
					var phoneSel = document.querySelector('#p > a');
					if(phoneSel!==null){
						var phone = phoneSel.innerText;
					}
					var emailSel = document.querySelector('#m > a');
					if(emailSel!==null){
						var email = emailSel.innerText;
					}
					
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
                                    workExp,
									phone,
									email
                                };
                            }
                        }


                    }

                });


        people.push({ ...arrName, link }); // Add the person to the array
        peopleCount++; // Increment the people counter

        console.log(`Профессия: ${arrName.job}\nИмя: ${arrName.name}\nТелефон: ${arrName.phone}\nПочта: ${arrName.email}\nМесто жительства: ${arrName.residence}\nЗарплата: ${arrName.salary}\nОпыт работы: ${arrName.workExp}\n\n`);
		console.log(peopleCount);
                // Wait for some time before navigating to the next link
                await page.waitForTimeout(500);
            }
            // Запись содержимого из people в таблицу Excel
            people.forEach((person, index) => {
    const dataRow = sheet.row(index + 2);
    dataRow.cell(1).value(person.job);
    dataRow.cell(2).value(person.name);
    dataRow.cell(3).value(person.phone);
    dataRow.cell(4).value(person.email);
    dataRow.cell(5).value(person.residence);
    dataRow.cell(6).value(person.salary);
    dataRow.cell(7).value(person.workExp);
    dataRow.cell(8).value(person.link); // Write the link to the Excel cell
	
	    const linkCell = dataRow.cell(8);
    linkCell.value(person.link); // Write the link to the Excel cell
    linkCell.style("fontColor", "0563C1"); // Set the font color to blue
    linkCell.style("underline", true); // Underline the text
    linkCell.hyperlink(person.link); // Add the hyperlink to the cell
});

            // Применение границ к данным в таблице
            const dataRange = sheet.range('A1:H' + (people.length + 1)); // Update the range to include the link column
sheet.column('A').width(80);
sheet.column('B').width(32);
sheet.column('C').width(32);
sheet.column('D').width(32);
sheet.column('E').width(35);
sheet.column('F').width(19);
sheet.column('G').width(17);
sheet.column('H').width(40); // Set width for the link column


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