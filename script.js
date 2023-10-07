const puppeteer = require('puppeteer')
const ExcelJS = require('exceljs');
const TelegramBot = require('node-telegram-bot-api');

require('dotenv').config()

const { _centered, _price, _header } = require('./styles');

const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: true });

bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  await sendFile(chatId)
});

async function sendFile(chatId) {
  const buffer = await getXlsxBuffer()
  const today = new Date().toLocaleDateString('en-GB').split('/').join('-')
  const options = { filename: `${today}.xlsx`, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }

  await bot.sendDocument(chatId, buffer, {}, options);
}

async function getXlsxBuffer() {
  const workbook = new ExcelJS.Workbook();

  const categories = process.env.CATEGORIES.split(',')
  console.log(categories)

  for (const category of categories) {
    const products = await getProductsFromCategory(category)

    console.log(`${products.length} ${category}`)

    const sheet = workbook.addWorksheet(category);

    sheet.columns = [
      { header: 'Producto', key: 'name', width: 50 },
      { header: 'Stock', key: 'stock', width: 20, style: _centered  },
      { header: 'kg.', key: 'quantity', width: 15, style: _centered },
      { header: 'Precio', key: 'price', width: 20, style: _price },
      { header: 'Precio por kilo', key: 'pricePerKg', width: 20, style: _price },
      { header: '', key: 'divider', width: 15 },
      { header: 'U. de Venta (kg)', key: 'sellUnitKg', width: 23, style: _centered },
      { header: 'Ganancia (%)', key: 'feePercent', width: 20, style: _centered },
      { header: 'Ganancia ($)', key: 'feeNominal', width: 20, style: _price },
      { header: 'Precio final', key: 'finalPrice', width: 20, style: _price },
    ];

    const header = sheet.getRow(1)
    header.font = _header.font
    header.alignment = _header.alignment
    header.fill = _header.fill
    header.height = 20

    for (const product of products) {
      const isAvailable = !product.offers.availability?.includes('OutOfStock')
      const pricePerGraim = product.offers.price / product.weight.value
      const sellUnitKg = calcProductSellUnit(product, category)

      sheet.addRow({
        name: product.name,
        stock: isAvailable ? 'Sí' : 'No',
        quantity: Number(product.weight.value),
        price: Number(product.offers?.price),
        pricePerKg: pricePerGraim * 1,
        sellUnitKg,
        feePercent: 100,
        feeNominal: (pricePerGraim * sellUnitKg) * 1,
        finalPrice: (pricePerGraim * sellUnitKg) * 2,
      });
    }
  }

  const buffer = await workbook.xlsx.writeBuffer()

  return buffer;
}

function calcProductSellUnit(product, type) {
  if (type === 'oxidos') {
    const sellsByFiftyGraims = ['Oxido Hierro Amarillo', 'Oxido Hierro Rojo', 'Oxido de Manganeso'].includes(product.name)

    if (sellsByFiftyGraims) return .050

    return .020
  }

  if (type === 'pigmentos-puros') return .010

  if (type === 'esmaltes-ceramicos') return .250
}

async function getProductsFromCategory(category) {
  const TWO_MINUTES = 2 * 60 * 1000

  let browser

  try {
    browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] })

    const page = await browser.newPage()

    await page.setViewport({ width: 1920, height: 1080 })

    page.setDefaultNavigationTimeout(TWO_MINUTES)

    await page.goto(`${process.env.PROVIDER_URL}/${category}`)

    const products = []

    let elementHandles = []

    let shouldKeepScrolling = true;

    while (shouldKeepScrolling) {
      await page.evaluate((y) => { document.scrollingElement.scrollBy(0, y); }, 5000);

      await page.waitForTimeout(3000)

      const loadMoreButton = await page.$('.js-load-more')
      const isLoadMoreButtonHidden = await loadMoreButton?.isHidden()

      if (loadMoreButton && !isLoadMoreButtonHidden) {
        await loadMoreButton.click('.js-load-more')
      }

      const elementHandlesOnScroll = await page.$$('script[type="application/ld+json"]')

      if (elementHandles.length === elementHandlesOnScroll.length) shouldKeepScrolling = false;

      elementHandles = elementHandlesOnScroll
    }

    for (const handle of elementHandles) {
      const elementJSON = await handle.evaluate(el => el.innerHTML)
      console.log(elementJSON)
      const literal = JSON.parse(elementJSON)

      if (literal['@type'] === 'Product') {
        products.push(literal)
      }
    }

    return products
  } catch (error) {
    console.log(error)
  } finally {
    await browser?.close()
  }
}
