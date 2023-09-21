const fs = require('fs')
const jsdom = require("jsdom")

const ExcelJS = require('exceljs');

const today = new Date().toLocaleDateString('en-GB').split('/').join('-')
const files = [ 'oxidos', 'pigmentos-puros', 'esmaltes-ceramicos']

const workbook = new ExcelJS.Workbook();

for (const file of files) {
  const fileProducts = getProductsFromFile(file)

  const sheet = workbook.addWorksheet(file);

  sheet.columns = [
    { header: 'Producto', key: 'name', width: 50 },
    { header: 'Precio', key: 'price', width: 20 },
    { header: 'Stock', key: 'stock', width: 20 },
  ];

  for (const product of fileProducts) {
    const isAvailable = !product.offers.availability?.includes('OutOfStock')

    sheet.addRow({
      name: product.name,
      price: isAvailable ? product.offers.price : '-',
      stock: isAvailable ? 'Sí' : 'No'
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
