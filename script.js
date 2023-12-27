const puppeteer = require('puppeteer')
const ExcelJS = require('exceljs')
const TelegramBot = require('node-telegram-bot-api')
const cron = require('node-cron')

require('dotenv').config()

const healtcheck = require('./healtcheck');
const { _centered, _price, _header } = require('./styles');

const authorizedChatIds = [process.env.CHAT_ID, process.env.ADMIN_CHAT_ID]
const PORT = process.env.PORT || 3000;

const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: true })

bot.on('message', async (msg) => {
  if (msg.text.startsWith('/')) return

  const chatId = String(msg.chat.id)
  const isAuthorized = authorizedChatIds.includes(chatId)

  console.log(`Received ${isAuthorized ? 'authorized' : 'unauthorized'} message from ${chatId}: ${msg.text}`)

  if (!isAuthorized) {
    await bot.sendMessage(chatId, 'Acceso no autorizado.')
    return
  }

  await bot.sendMessage(chatId, 'Hola amorcito, usá alguno de los comandos ❤')
})

bot.onText(/\/tetragramaton/, async (msg) => {
  const chatId = String(msg.chat.id)
  const isAuthorized = authorizedChatIds.includes(chatId)

  console.log(`Received ${isAuthorized ? 'authorized' : 'unauthorized'} message from ${chatId}: ${msg.text}`)

  if (!isAuthorized) {
    await bot.sendMessage(chatId, 'Acceso no autorizado.')
    return
  }

  await bot.sendMessage(chatId, 'Consiguiendo los productos...')

  try {
    await sendFile(chatId)
  } catch (error) {
    await bot.sendMessage(chatId, 'Hubo un error, comunicate con tu amorcito para que lo arregle ❤')
    console.log(error)
  }
})

bot.onText(/\/amorcito/, async function onLoveText(msg) {
  const chatId = String(msg.chat.id)
  const isAuthorized = authorizedChatIds.includes(chatId)

  console.log(`Received ${isAuthorized ? 'authorized' : 'unauthorized'} message from ${chatId}: ${msg.text}`)

  if (!isAuthorized) {
    await bot.sendMessage(chatId, 'Acceso no autorizado.')
    return
  }

  const opts = {
    reply_to_message_id: msg.message_id,
    reply_markup: JSON.stringify({
      keyboard: [
        ['Sí amorcito ❤'],
        ['Mucho mucho ❤']
      ]
    })
  }

  bot.sendMessage(msg.chat.id, 'Me querés mucho?', opts)
})

cron.schedule('0 12 * * 1-5', async () => {
  console.log('Running cron...')
  await bot.sendMessage(process.env.CHAT_ID, 'Hola amorcito ❤')
  await bot.sendMessage(process.env.CHAT_ID, 'Te voy a mandar la nueva lista de precios...')
  await bot.sendMessage(process.env.ADMIN_CHAT_ID, `Running CRON.`)

  try {
    await sendFile(process.env.CHAT_ID)
    await bot.sendMessage(process.env.CHAT_ID, 'Que tengas un lindo día en el trabajo amorcito ❤')
  } catch (error) {
    await bot.sendMessage(process.env.CHAT_ID, 'Hubo un error, comunicate con tu amorcito para que lo arregle ❤')
    await bot.sendMessage(process.env.ADMIN_CHAT_ID, `Error running CRON.`)
    console.log('Error running CRON:')
    console.error(error)
  }
})

async function sendFile(chatId) {
  const { buffer, summary } = await getXlsxBuffer()
  const today = new Date().toLocaleDateString('en-GB').split('/').join('-')
  const options = { filename: `${today}.xlsx`, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }

  const summaryString = summary.join(', ')

  await bot.sendMessage(chatId, `Completado con: ${summaryString}.`)
  await bot.sendDocument(chatId, buffer, {}, options)
  await bot.sendMessage(chatId, `Gracias por usar el bot amorcito, te amo ❤`)
  await bot.sendMessage(process.env.ADMIN_CHAT_ID, `Delivered XLSX file with ${summaryString} to ${chatId}.`)
}

async function getXlsxBuffer() {
  const workbook = new ExcelJS.Workbook()

  const categories = process.env.CATEGORIES.split(',')

  const summary = []

  console.log('Generating XLSX...')

  for (const category of categories) {
    const products = await getProductsFromCategory(category)

    const categorySummary = `${products.length} ${category}`
    console.log(`Retrieved ${categorySummary}`)
    summary.push(categorySummary)

    const columns = [
      { header: 'Stock', key: 'stock', width: 10, style: _centered  },
      { header: 'Producto', key: 'name', width: 50 },
      { header: 'U. de Venta (kg)', key: 'sellUnitKg', width: 23, style: _centered },
      { header: 'Ganancia ($)', key: 'feeNominal', width: 20, style: _price },
      { header: 'Precio final', key: 'finalPrice', width: 20, style: _price },
      { header: '', key: 'divider', width: 35 },
      { header: 'U. de Compra (kg).', key: 'buyUnitKg', width: 25, style: _centered },
      { header: 'Precio', key: 'price', width: 20, style: _price },
      { header: 'Precio por kilo', key: 'pricePerKg', width: 20, style: _price },
    ]

    const sheet = workbook.addWorksheet(category, { views: [{ state: 'frozen', ySplit: 2 }] })

    sheet.mergeCells('C1:E1');
    sheet.getCell('C1').value = 'VENTA';

    sheet.mergeCells('G1:I1');
    sheet.getCell('G1').value = 'COMPRA';

    sheet.getRow(2).values = columns.map((col) => col.header)

    sheet.columns = columns.map((col) => ({ key: col.key, width: col.width, style: col.style  }))

    const headers = sheet.getRows(1, 2)

    for (const row of headers) {
      row.font = _header.font
      row.alignment = _header.alignment
      row.fill = _header.fill
      row.height = 20
    }

    for (const product of products) {
      const isAvailable = !product.offers.availability?.includes('OutOfStock')
      const pricePerGraim = product.offers.price / product.weight.value
      const sellUnitKg = calcProductSellUnit(product, category)

      sheet.addRow({
        stock: isAvailable ? '✓' : '✕',
        name: product.name,
        sellUnitKg,
        feeNominal: (pricePerGraim * sellUnitKg) * 1,
        finalPrice: (pricePerGraim * sellUnitKg) * 2,
        buyUnitKg: Number(product.weight.value),
        price: Number(product.offers?.price),
        pricePerKg: pricePerGraim * 1,
      })
    }
  }

  const buffer = await workbook.xlsx.writeBuffer()

  return { buffer, summary }
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

    let shouldKeepScrolling = true

    while (shouldKeepScrolling) {
      await page.evaluate((y) => { document.scrollingElement.scrollBy(0, y) }, 5000)

      await page.waitForTimeout(3000)

      const loadMoreButton = await page.$('.js-load-more')
      const isLoadMoreButtonHidden = await loadMoreButton?.isHidden()

      if (loadMoreButton && !isLoadMoreButtonHidden) {
        await loadMoreButton.click('.js-load-more')
      }

      const elementHandlesOnScroll = await page.$$('script[type="application/ld+json"]')

      if (elementHandles.length === elementHandlesOnScroll.length) shouldKeepScrolling = false

      elementHandles = elementHandlesOnScroll
    }

    for (const handle of elementHandles) {
      const elementJSON = await handle.evaluate(el => el.innerHTML)
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

healtcheck.listen(PORT, () => {
  console.log(`Healtcheck listening on port ${PORT}`);
});
