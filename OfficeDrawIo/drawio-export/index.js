const yargs = require('yargs');
const puppeteer = require('puppeteer');
const fs = require('fs');
var path = require('path');

process.on('unhandledRejection', (e) => {
  throw e;
});

process.on('uncaughtException', (e) => {
  console.error(e);
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

function fileUrl(str) {
  if (typeof str !== 'string') {
      throw new Error('Expected a string');
  }

  var pathName = path.resolve(str).replace(/\\/g, '/');

  // Windows drive letter must be prefixed with a slash
  if (pathName[0] !== '/') {
      pathName = '/' + pathName;
  }

  return encodeURI('file://' + pathName);
};

const main = async () => {

  const infile = argv._[0];
  const outfile = argv._[1];

  const xml = await readFile(infile);
  const browser = await puppeteer.launch({
    headless: true,
  });

  try {
    const page = await browser.newPage();

    const exportUrl = fileUrl(__dirname) + '/drawio/export3.html';
    await page.goto(exportUrl, {
      waitUntil: 'networkidle0'
    });

    await page.evaluate((obj) => render(obj), {
      xml,
      format: 'png',
      w: 0,
      h: 0,
      border: 0,
      bg: 'none',
      scale: 1,
    });

    await page.waitForSelector('#LoadingComplete');

    const boundsJson = await page.mainFrame().$eval('#LoadingComplete', (div) => div.getAttribute('bounds'));
    const bounds = JSON.parse(boundsJson);

    const fixingScale = 1; // 0.959;
    const w = Math.ceil(bounds.width * fixingScale);
    const h = Math.ceil(bounds.height * fixingScale);

    await page.setViewport({
      width: w,
      height: h
    });

    await page.screenshot({
      omitBackground: true,
      type: 'png',
      fullPage: true,
      path: outfile,
    });


  } finally {
    await browser.close();
  }
}

main();

