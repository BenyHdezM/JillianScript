const dotenv = require("dotenv");
dotenv.config();
var XLSX = require("xlsx");
const puppeteer = require("puppeteer");
const workbook = XLSX.readFile(process.env.XLSX_PATH);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const numerosGenerados = {};

(async () => {

    const browser = await puppeteer.launch({ headless: false });
    var pages = await browser.pages();
    const page = pages[pages.length - 1];

    const username = process.env.USER_EMAIL;
    const password = process.env.USER_PASSWORD;
    

    // LOGIN
    await page.goto(process.env.LOGIN_URL, { waitUntil: 'load', timeout: 0 });
    await page.setViewport({ width: 1080, height: 1024 });

    // while (await page.$('#Username') === null) {
    //     console.log('#Username Selector not found');
    //     delaytime(1000);
    // }

    if (await page.$('#Username') !== null) {
        await page.waitForSelector('#Username');
        await page.type('#Username', username);
        await page.type('#txtPassword', password);
        await page.click("#show_password");
        await page.click(".btn.btn-login");
    }

    const menuSelector = `#tenant_apps [data-title="Voyager 7S"] a`;

    await page.waitForSelector(menuSelector);
    if (await page.$(menuSelector) !== null) await page.click(menuSelector);
    else console.log(menuSelector + ' not found');

    await page.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });

    // WEB PAGE 2
    pages = await browser.pages();
    const page2 = pages[pages.length - 1];
    await page2.setViewport({ width: 1080, height: 1024 });

    await page2.waitForSelector('#cmdLogin');
    // await page2.select('#loginAlias', 'Test'); // ENABLE FOR TESTING
    if (await page2.$(`#cmdLogin`) !== null) await page2.click(`#cmdLogin`);
    else console.log('cmdLogin not found');

    page2.on('dialog', async dialog => {
        await dialog.accept();
    });

    await page2.waitForSelector('#top-menu-wrap');
    while (await page2.$(`#top-menu-wrap`) === null) {
        console.log('#top-menu-wrap not found');
        timedelay(1000);
    }

    for (let index = 2; index > 0; index++) {
        try {
            await page2.goto(process.env.ADD_CROS_ET, { waitUntil: 'load', timeout: 0 });

            if (await page2.$(`#BtnSave_Button`) !== null && await page2.$(`#BillingProperty_LookupCode`) !== null) {
                const date = worksheet[`A${index}`].w;
                const lookupCode = returnLookupCode(worksheet[`C${index}`].v);
                const notes = worksheet[`D${index}`].v;
                const chargeCode = worksheet[`E${index}`].v;
                const amount = worksheet[`F${index}`].v;
                const account = worksheet[`G${index}`].v;
                const segment = worksheet[`H${index}`].v;
                const payee = worksheet[`I${index}`].v;
                const entity = 'cso';
                const arrDate = date.split('/');
                const numInvetario = generarNumeroInventario(lookupCode, date, arrDate, chargeCode);
                // const numInvetario = `${lookupCode}${date.slice(-2)}M${arrDate[0]}${chargeCode.substring(0, 3).toUpperCase()}${randomNumber}`;

                console.log({
                    Index: index,
                    Date: date,
                    LookupCode: lookupCode,
                    Notes: notes,
                    ChargeCode: chargeCode,
                    Amount: amount,
                    Account: account,
                    Segment: segment,
                    Payee: payee,
                    Entity: entity,
                    NumInventario: numInvetario
                });

                // await page2.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });
                //Charge To
                await page2.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });
                await page2.waitForSelector('#BillingProperty_LookupCode');
                await page2.type('#BillingProperty_LookupCode', entity);
                await page2.type('#BillingPerson_LookupCode', lookupCode.toLowerCase());
                await page2.type('#ChargeCode_LookupCode', chargeCode);
                await page2.click('#ChargeNotes_TextBox');
                await page2.waitForSelector('#ChargeNotes_TextBox');
                // await page2.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });
                await page2.type('#ChargeNotes_TextBox', notes);

                //Payable To
                await page2.type('#BilledProperty_LookupCode', worksheet[`B${index}`].v.toLowerCase());
                await page2.type('#BilledPerson_LookupCode', payee);
                await page2.type('#PayableAccount_LookupCode', account);
                await page2.type('#PayableInvNum_TextBox', numInvetario);
                await page2.type('#PayableNotes_TextBox', notes);

                //Trans Batch Details
                await page2.type('#TransDateOccrred_TextBox', date);
                await page2.type('#TransPostDate_TextBox', `${arrDate[0]}/${arrDate[2]}`);
                await page2.click('#Amount_TextBox', { clickCount: 3 });
                await page2.waitForSelector('#Amount_TextBox');
                await page2.type('#Amount_TextBox', amount.toString());
                // await page2.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });
                //Payable Segment
                await page2.type('#PayableSegment2_LookupCode', segment);

                //Charge Segment
                await page2.type('#ChargeSegment1_LookupCode', lookupCode.toLowerCase());

                // Finish - Save
                await page2.click('#Body1');
                await page2.click('#BtnSave_Button');
                await page2.waitForNetworkIdle({ idleTime: 1000, timeout: 0 });
            }

        } catch (e) {
            console.log(e);
            index = -1;
        }
    }
    await browser.close();
})();

function returnLookupCode(name) {
    var arrayWords = name.toLowerCase().split(' ');
    if (arrayWords[0] == "conrac") {
        return arrayWords[arrayWords.length - 1];
    } else {
        return arrayWords[0].toLowerCase();
    }
}

function generarNumeroInventario(lookupCode, date, arrDate, chargeCode) {
    let randomNumber;
    let numInventario;

    do {
        // Generamos un número aleatorio (ajusta según sea necesario)
        randomNumber = Math.floor(Math.random() * 100);

        // Creamos el número de inventario
        numInventario = `${lookupCode}${date.slice(-2)}M${arrDate[0]}${chargeCode.substring(0, 3).toUpperCase()}${randomNumber}`;
    } while (numerosGenerados[numInventario]);

    // Si el número es único, lo almacenamos en el diccionario
    numerosGenerados[numInventario] = true;

    // Retornamos el número de inventario generado
    return numInventario;
}
