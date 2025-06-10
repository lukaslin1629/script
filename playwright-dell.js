const { chromium } = require('playwright');

(async () => {
    const serial = process.argv[2];
    const url = `https://www.dell.com/support/home/en-us/product-support/servicetag/${serial}/overview`;

    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(4000);

    const h1 = await page.$('h1');
    const model = h1 ? await h1.innerText() : 'Not Found';

    console.log(model);

    await browser.close();
})();