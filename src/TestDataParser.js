import xlsx from 'xlsx';

export default class TestDataParser {
  constructor(path) {
    this.workbook = xlsx.readFile(path, { cellDates: true });
  }
  parseAllSheets() {
    const result = {};
    this.workbook.SheetNames.forEach((sheetName) => {
      result[sheetName] = this.parseSheet(sheetName);
    });
    return result;
  }
  parseSheet(sheetName) {
    const sheet = this.workbook.Sheets[sheetName];
    const blocks = parseBlockNames(sheet);
    const result = {};
    blocks.forEach((line, blockName) => {
      result[blockName] = this.parseBlock(sheetName, line + 2, blockName.startsWith('#'));
    });
    return result;
  }

  parseBlock(sheetName, line, table) {
    return table ? this.parseTableBlock(sheetName, line) : this.parseArrayBlock(sheetName, line);
  }
  /**Распарсить вертикальный блок */
  parseArrayBlock(sheetName, line) {
    const sheet = this.workbook.Sheets[sheetName];
    const result = {};
    while (sheet['A' + line]) {
      result[sheet['A' + line].w] = getCellValue(sheet['B' + line]);
      line++;
    }
    return result;
  }
  parseTableBlock(sheetName, line) {
    // const sheet = this.workbook.Sheets[sheetName];
    const result = [];
    const header = this.parseTableRow(sheetName, line);
    let row = this.parseTableRow(sheetName, ++line);
    while (row.length) {
      const map = new Map(header.map((colName, i) => [colName, row[i]]));
      result.push(Object.fromEntries(map));
      row = this.parseTableRow(sheetName, ++line);
    }
    return result;
  }
  parseTableRow(sheetName, line) {
    const sheet = this.workbook.Sheets[sheetName];
    const row = [];
    let char = 'A';
    while (sheet[char + line]) {
      row.push(getCellValue(sheet[char + line]));
      char = String.fromCharCode(char.charCodeAt() + 1); // A->B->C ...
    }
    return row;
  }
}

/**
 * Получить значение ячейки
 * @param {*} cell
 */
function getCellValue(cell) {
  // глюк? сдвоенные ячейки (две даты равны в соседних ячейках) во второй ячейке выводятся строкой
  if (cell.t == 'd' && !(cell.v instanceof Date)) {
    return new Date(cell.v);
  }
  return cell.v;
}
/**
 * поиск блоков имя блока=>номер строки
 * {
    '/request/' => 2,
    '#balancesByDate#' => 8,
    '#ratesByDate#' => 16,
    '#interestsStatement#' => 25
  }
 * @param {*} sheet
 * @returns Map
 */
function parseBlockNames(sheet) {
  return new Map(
    Object.keys(sheet)
      .filter((cell) => cell.match(/^A/)) // ячейки A
      .map((cell) => [cell, sheet[cell].w.match(/^[/#].*[/#]/)]) // начинаются с / или $
      .filter((cell) => cell[1])
      .map((cell) => [cell[1][0], +cell[0].substring(1)]) // request,2
  );
}
