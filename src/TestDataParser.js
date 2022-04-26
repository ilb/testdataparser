import xlsx from 'xlsx';

export default class TestDataParser {
  constructor(path) {
    this.workbook = xlsx.readFile(path, { cellDates: true });
  }
  parseAllSheets() {
    const result = new Map();
    this.workbook.SheetNames.forEach((sheetName) => {
      result.set(sheetName, this.parseSheet(sheetName));
    });
    return result;
  }
  parseSheet(sheetName) {
    const sheet = this.workbook.Sheets[sheetName];
    const blocks = parseBlockNames(sheet);
    const result = new Map();
    blocks.forEach((line, blockName) => {
      result.set(blockName, this.parseBlock(sheetName, line + 2, blockName.startsWith('#')));
    });
    return result;
  }

  parseBlock(sheetName, line, table) {
    return table ? this.parseTableBlock(sheetName, line) : this.parseArrayBlock(sheetName, line);
  }
  /**Распарсить вертикальный блок */
  parseArrayBlock(sheetName, line) {
    const sheet = this.workbook.Sheets[sheetName];
    const result = new Map();
    while (sheet['A' + line]) {
      result.set(sheet['A' + line].w, sheet['B' + line].v);
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
      result.push(map);
      row = this.parseTableRow(sheetName, ++line);
    }
    return result;
  }
  parseTableRow(sheetName, line) {
    const sheet = this.workbook.Sheets[sheetName];
    const row = [];
    let char = 'A';
    while (sheet[char + line]) {
      row.push(sheet[char + line].v);
      char = String.fromCharCode(char.charCodeAt() + 1); // A->B->C ...
    }
    return row;
  }
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
