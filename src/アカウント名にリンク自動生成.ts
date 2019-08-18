function addLink() {
  addAccountLink();
  // addAuthorLink()
  // addCircleLink()
}

function addAccountLink() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const accCol = sheet.getRange("B2:B39");
  for (let i = 1; i <= accCol.getLastRow(); i++) {
    const cell = accCol.getCell(i, 1);
    if (cell.getValue() == "") {
      break;
    }

    const val = cell.getValue();
    const formula = '=HYPERLINK("https://vocalodon.net/' + val + '","' + val + '")';
    cell.setFormula(formula);
  }
}

function addAuthorLink() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(3, 1, 30, 10);
  for (let i = 1; i <= data.getLastRow(); i++) {
    const nameCel = data.getCell(i, 1);
    const name = nameCel.getValue();
    if (name == "") {
      break;
    }
    const authorCell = data.getCell(i, 9);

    const val = authorCell.getValue();
    if (val == "") {
      continue;
    }
    const formula = '=HYPERLINK("https://webcatalog-free.circle.ms/Circle/List?keyword=' + val + '","' + val + '")';
    authorCell.setFormula(formula);
  }
}

function addCircleLink() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(3, 1, 30, 10);
  for (let i = 1; i <= data.getLastRow(); i++) {
    const nameCel = data.getCell(i, 1);
    const name = nameCel.getValue();
    if (name == "") {
      break;
    }
    const circleCell = data.getCell(i, 7);

    const val = circleCell.getValue();
    if (val == "") {
      continue;
    }
    const formula = '=HYPERLINK("https://webcatalog-free.circle.ms/Circle/List?keyword=' + val + '","' + val + '")';
    circleCell.setFormula(formula);
  }
}
