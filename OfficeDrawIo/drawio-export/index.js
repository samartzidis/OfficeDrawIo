const yargs = require('yargs');
const puppeteer = require('puppeteer');
const fs = require('fs');
var path = require('path');
const url = require('url');

process.on('unhandledRejection', (e) => {
    throw e;
});

process.on('uncaughtException', (e) => {
    throw e;
});

const { argv } = yargs;

const readFile = (file) => new Promise((resolve, reject) => {
    fs.readFile(file, 'utf-8', (err, res) => {
        if (err) {
            reject(err);
        } else {
            resolve(res);
        }
    });
});

const interceptedHosts = {
    'www.draw.io': 'drawio',
    'math.draw.io': 'mathjax'
};

function interceptRequest(interceptedRequest)
{
    const reqUrl = interceptedRequest.url();
    const reqUrlObj = url.parse(reqUrl);                  
    const reqHostname = reqUrlObj.hostname.toLowerCase();
    const reqPathWithoutQuery = reqUrlObj.path.split('?')[0];

    if(interceptedHosts.hasOwnProperty(reqHostname)) {
        const localFolder = interceptedHosts[reqHostname];
        let localFilePath = path.join(__dirname, localFolder, reqPathWithoutQuery);
        fs.readFile(localFilePath, (err, res) => {                
            if (err) {                                                
                //console.log(`[${reqUrl}] -> Failed ${err.message}, using original URL.`);
                //interceptedRequest.continue();
                console.log(`[${reqUrl}] -> Failed: ${err.message}`);
                interceptedRequest.abort();
                throw new Error(err.message);
            } else {
                console.log(`[${reqUrl}] -> [${localFilePath}]`);
                interceptedRequest.respond({status: 200, body: res});
            }
        });
    }            
    else {
        console.log(`${interceptedRequest.url()} -> ${interceptedRequest.url()}`);
        interceptedRequest.continue();
    }
}

const main = async () => {

    const infile = argv._[0];
    const outfile = argv._[1];

    const xml = await readFile(infile);
    const browser = await puppeteer.launch({ headless: true });

    try {
        const page = await browser.newPage();

        await page.setRequestInterception(true);

        page.on('request', (interceptedRequest) => interceptRequest(interceptedRequest));

        await page.goto('https://www.draw.io/export3.html', { waitUntil: 'networkidle0' });
        await page.evaluate((obj) => render(obj), { xml, format: 'png', w: 0, h: 0, border: 0, bg: 'none', scale: 1 });
        await page.waitForSelector('#LoadingComplete');

        const boundsJson = await page.mainFrame().$eval('#LoadingComplete', (div) => div.getAttribute('bounds'));
        const bounds = JSON.parse(boundsJson);

        const fixingScale = 1; // 0.959;
        const w = Math.ceil(bounds.width * fixingScale);
        const h = Math.ceil(bounds.height * fixingScale);

        await page.setViewport({ width: w, height: h });

        await page.screenshot({ omitBackground: true, type: 'png', fullPage: true, path: outfile });


    } 
    finally {
        await browser.close();
    }
};

main();

