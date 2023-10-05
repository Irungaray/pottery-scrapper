const _centered = {
  alignment: {
    vertical: 'middle',
    horizontal: 'center',
  }
}

const _header = {
  ..._centered,
  font: {
    name: 'Liberation Sans',
    family: 2,
    size: 14,
    bold: true,
  },
  fill: {
    type: 'pattern',
    pattern:'solid',
    fgColor: { argb: '729fcf' },
  }
}

const _price = {
  ..._centered,
  numFmt: '"$"#,##0.00;[Red]\-"$"#,##0.00'
}

module.exports = { _centered, _header, _price }
