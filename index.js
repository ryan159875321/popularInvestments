import puppeteer from "puppeteer";
import xlsx from "xlsx";
import moment from "moment";
import fs from "fs";
import cron from "node-cron";

const urls = {
    gold: { url: 'https://www.gold.co.uk/gold-price/', selector: 'span[name="current_price_field"]' },
    bitcoin: { url: 'https://www.bullionbypost.co.uk/bitcoin-price/bitcoin-price/', selector: 'span[name="current_price_field"]' },
    ethereum: { url: 'https://www.coingecko.com/en/coins/ethereum/gbp', selector: 'span[data-converter-target="price"]' },
    sp500: { url: 'https://www.hl.co.uk/shares/shares-search-results/v/vanguard-funds-plc-s-and-p-500-etf-usdgbp', selector: '#ls-ask-VUSA-L' },
    dowjones: { url: 'https://www.hl.co.uk/shares/shares-search-results/i/ishares-vii-plc-dow-jones-ind-avg-ucits2', selector: '#ls-ask-CIND-L' },
    nasdaq: { url: 'https://www.hl.co.uk/shares/stock-market-summary/nasdaq', selector: '#indices-val-NDX' }
};

const getPrices = async () => {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
    });

    const page = await browser.newPage();
    const prices = {};

    for (const [key, { url, selector }] of Object.entries(urls)) {
        await page.goto(url, {
            waitUntil: "domcontentloaded",
        });

        try {
            const acceptCookiesButton = await page.$('#acceptCookieButton');
            if (acceptCookiesButton) {
                await acceptCookiesButton.click();
                console.log('Accepted cookies on ', url);
            }
        } catch (error) {
            console.log('No cookie button appeared on ', url);
        }

        try {
            await page.waitForSelector(selector);

            const price = await page.$eval(selector, el => el.textContent);
            prices[key] = price.replace(/[^0-9.-]+/g, "");

            console.log(`${key} Price: `, prices[key]);
        } catch (error) {
            console.log(`Error getting price for ${key}`, error);
        }
    }

    await browser.close();
    return prices;
};

const saveToExcel = (prices) => {
    const filePath = 'prices.xlsx';
    let workbook;

    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
    }

    const sheetName = 'Prices';
    let worksheet;

    if (workbook.Sheets[sheetName]) {
        worksheet = workbook.Sheets[sheetName];
    } else {
        worksheet = xlsx.utils.aoa_to_sheet([['Date', 'Gold', 'Bitcoin', 'Ethereum', 'S&P 500', 'Dow Jones', 'Nasdaq']]);
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
    }

    const date = moment().format('YYYY-MM-DD');
    const newRow = [date, prices.gold, prices.bitcoin, prices.ethereum, prices.sp500, prices.dowjones, prices.nasdaq];

    const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    existingData.push(newRow);

    const updatedWorksheet = xlsx.utils.aoa_to_sheet(existingData);
    workbook.Sheets[sheetName] = updatedWorksheet;

    xlsx.writeFile(workbook, filePath);
};

const main = async () => {
    const prices = await getPrices();
    console.log("Prices: ", prices);
    saveToExcel(prices);
};

//cron.schedule('*/1 * * * *', () => { Use this to test every min
cron.schedule('0 18 * * *', () => { 
    console.log('Running the script at 6pm');
    main();
})

main(); 



//0 18 * * * ....... 6pm every day