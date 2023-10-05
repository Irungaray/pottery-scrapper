const fs = require('fs')
const jsdom = require("jsdom")

const ExcelJS = require('exceljs');
const { _centered, _price, _header } = require('./styles');

const today = new Date().toLocaleDateString('en-GB').split('/').join('-')
const files = ['oxidos', 'pigmentos-puros', 'esmaltes-ceramicos']

const workbook = new ExcelJS.Workbook();

for (const file of files) {
  const fileProducts = getProductsFromFile(file)

  const sheet = workbook.addWorksheet(file);

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

  for (const product of fileProducts) {
    const isAvailable = !product.offers.availability?.includes('OutOfStock')
    const pricePerGraim = product.offers.price / product.weight.value
    const sellUnitKg = getProductSellUnit(product, file)

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

workbook.xlsx.writeFile(`./xlsx/${today}.xlsx`)

function getProductsFromFile(file) {
  const html = fs.readFileSync(`./html/${today}/${file}.html`, { encoding: 'utf-8' })
  const dom = new jsdom.JSDOM(html);
  const elements = dom.window.document.querySelectorAll('script[type="application/ld+json"]')

  const products = []

  for (const element of elements) {
    const literal = JSON.parse(element.innerHTML);

    if (literal['@type'] === 'Product') {
      products.push(literal)
    }
  }

  fs.writeFileSync(`./json/${today}/${file}.json`, JSON.stringify(products, null, 4))

  return products
}

function getProductSellUnit(product, type) {
  if (type === 'oxidos') {
    const sellsByFiftyGraims = ['Oxido Hierro Amarillo', 'Oxido Hierro Rojo', 'Oxido de Manganeso'].includes(product.name)

    if (sellsByFiftyGraims) return .050

    return .020
  }

  if (type === 'pigmentos-puros') return .010

  if (type === 'esmaltes-ceramicos') return .250
}