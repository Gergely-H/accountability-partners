const getSpreadsheet = () => {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID);
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  return spreadsheet;
};

const getHabitSheet = () => {
  const sheetName = PropertiesService.getScriptProperties().getProperty(SHEET_TAB_NAME);
  const sheet = getSpreadsheet().getSheetByName(sheetName);

  return sheet;
};

/**
 * Convert column letter to number (A=1, B=2, etc.)
 */
const convertColumnNameToIndex = (columnName) => columnName.charCodeAt(0) - 'A'.charCodeAt(0) + 1;

const getCellRelativeToCell = (baseCell, rowOffset, columnOffset, sheet = getHabitSheet()) => {
  const baseRow = baseCell.getRow();
  const baseColumn = baseCell.getColumn();

  const targetRow = baseRow + rowOffset;
  const targetColumn = baseColumn + columnOffset;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const isTargetCellWithinSheet = targetRow >= 0 && targetRow <= lastRow && targetColumn >= 0 && targetColumn <= lastColumn;
  
  if (isTargetCellWithinSheet) {
    return sheet.getRange(targetRow, targetColumn);
  }

  return null;
};

const getFirstEmptyCellInColumn = (fromRowIndex, columnIndex, sheet = getHabitSheet()) => {
  const lastRow = sheet.getLastRow();

  const range = sheet.getRange(fromRowIndex, columnIndex, lastRow, 1);
  const values = range.getValues();

  for (let rowIndex = 0; rowIndex <= values.length - 1; rowIndex++) {
    const isFirstEmptyCellInRow = values[rowIndex][0] === "";
    if (isFirstEmptyCellInRow) {
      return sheet.getRange(rowIndex + fromRowIndex, columnIndex);
    }
  }

  return null;
};

const getRowDistanceBetweenCells = (cell1, cell2) => {
  const row1 = cell1.getRow();
  const row2 = cell2.getRow();

  const distance = Math.abs(row2 - row1);

  return distance;
};
