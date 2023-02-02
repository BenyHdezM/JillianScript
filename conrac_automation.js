const dotenv = require("dotenv");
dotenv.config();
const XLSX = require("xlsx");
const puppeteer = require("puppeteer");

(async () => {

    const browser = await puppeteer.launch({ headless: false });
    var pages = await browser.pages();
    const page = pages[pages.length - 1];

    const username = process.env.USER_EMAIL;
    const password = process.env.USER_PASSWORD;

    await page.goto(process.env.LOGIN_URL);

    // LOGIN
    await page.setViewport({ width: 1080, height: 1024 });

    await page.waitForSelector('#Username');
    await page.type('#Username', username);
    await page.type('#txtPassword', password);
    await page.click("#show_password");
    await page.click(".btn.btn-auth");

    await page.waitForNetworkIdle({ idleTime: 500, timeout: 10000 });
    const menuSelector = `#tenant_apps [data-title="Voyager 7S"] a`;
    await page.waitForSelector(menuSelector);

    await page.click(menuSelector);

    await page.waitForNetworkIdle({ idleTime: 500, timeout: 10000 });

    // WEB TAB 2
    pages = await browser.pages();
    const page2 = pages[pages.length - 1];
    await page2.setViewport({ width: 1080, height: 1024 });

    await page2.waitForSelector("#cmdLogin");
    await page2.click("#cmdLogin");
    page2.on('dialog', async dialog => {
        await dialog.accept();
    });



    var workbook = XLSX.readFile(process.env.XLSX_PATH);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];

    for (let index = 2; index > 0; index++) {
        try {
            await page.waitForNetworkIdle({ idleTime: 500, timeout: 10000 });
            await page2.goto(process.env.ADD_CROS_ET);
            await page2.setViewport({ width: 1080, height: 1024 });

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
            const numInvetario = `${lookupCode}${date.slice(-2)}M${arrDate[0]}${chargeCode.substring(0, 1)}`;

            console.log({
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

            //Charge To
            await page2.type('#BillingProperty_LookupCode', entity);
            await page2.type('#BillingPerson_LookupCode', lookupCode.toLowerCase());
            await page2.type('#ChargeCode_LookupCode', chargeCode);
            await page2.click('#ChargeNotes_TextBox');
            await page.waitForNetworkIdle({ idleTime: 500, timeout: 10000 });
            await page2.type('#ChargeNotes_TextBox', notes);

            //Payable To
            await page2.type('#BilledProperty_LookupCode', lookupCode.toLowerCase());
            await page2.type('#BilledPerson_LookupCode', payee);
            await page2.type('#PayableAccount_LookupCode', account);
            await page2.type('#PayableInvNum_TextBox', numInvetario);
            await page2.type('#PayableNotes_TextBox', notes);

            //Trans Batch Details
            await page2.type('#TransDateOccrred_TextBox', date);
            await page2.type('#TransPostDate_TextBox', `${arrDate[0]}/${arrDate[2]}`);
            await page2.click('#Amount_TextBox', { clickCount: 3 });
            await page2.type('#Amount_TextBox', amount.toString());

            //Payable Segment
            await page2.type('#PayableSegment2_LookupCode', segment);

            //Charge Segment
            await page2.type('#ChargeSegment1_LookupCode', lookupCode.toLowerCase());

            //Finish
            await page2.click('#Body1');
            await page2.click('#BtnSave_Button');


        } catch (e) {
            index = -1;
        }
    }
    await browser.close();
})();

function returnLookupCode(name) {
    var arrayWords = name.split(' ');
    if (arrayWords[0] === "Conrac") {
        return arrayWords[arrayWords.length - 1];
    } else {
        return arrayWords[0].toLow;
    }
}